using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using log4net.Core;
using Microsoft.Office.Interop.Excel;
using Internal.Framework.IO;
using Internal.Framework.Storage;

namespace Spreadsheet.Handler
{
    internal static class WorksheetUtilities
    {
        public enum ResizeType
        {
            Height,
            Width,
            HeightAndWidth
        }

        public enum TestType
        {
            AssayLevel,
            ImpurityLevelStd,
            Impurities,
            WaterContent,
            Dissolution,
            Volatiles,
            CleaningVerification,
            AssayLevel_Impurities,
            LimitTest_LC,
            ReportingLimit,
            LimitTest_Volatiles,
            DetectionLimit,
            Identification,
            Volatile_Impurities
        }

        public static readonly Dictionary<TestType, string> TestTypeDisplayMap = new Dictionary<TestType, string>
        {
            { TestType.AssayLevel, "Assay Level" },
            { TestType.ImpurityLevelStd, "Impurity Level Std" },
            { TestType.Impurities, "Impurities" },
            { TestType.WaterContent, "Water Content" },
            { TestType.Dissolution, "Dissolution" },
            { TestType.Volatiles, "Volatiles" },
            { TestType.CleaningVerification, "Cleaning Verification" },
            { TestType.AssayLevel_Impurities, "Assay Level_Impurities" },
            { TestType.LimitTest_LC, "Limit Test (LC)" },
            { TestType.ReportingLimit, "Reporting Limit" },
            { TestType.LimitTest_Volatiles, "Limit Test (Volatiles)" },
            { TestType.DetectionLimit, "Detection Limit" },
            { TestType.Identification, "Identification" },
            { TestType.Volatile_Impurities, "Volatile_Impurities" }
        };

        private static Application _app;
        [DllImport("user32.dll")]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        public static Application GetExcelApp()
        {
            if (_app == null)
            {
                _app = new Application();
                // We don't want the user to get ANY dialogs/prompts from Excel!
                _app.Interactive = false;
            }
            return _app;
        }

        public static void ReleaseExcelApp()
        {
            if (_app != null)
            {
                // ReSharper disable RedundantAssignment
                uint excelProcessId = 0;
                // ReSharper restore RedundantAssignment
                GetWindowThreadProcessId(new IntPtr(_app.Hwnd), out excelProcessId);

                try
                {
                    _app.Quit();
                }
                catch (Exception ex)
                {
                    Logger.LogMessage(
                        "Error executing WorksheetUtilities.ReleaseExcelApp. Could not quit the application.\r\n" +
                        ex.Message, Level.Error);
                }
                finally
                {
                    // Release the app
                    ReleaseComObject(_app);
                    _app = null;

                    try
                    {
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        GC.Collect();
                    }
                    catch (Exception ex)
                    {
                        Logger.LogMessage(
                            "Error executing WorksheetUtilities.ReleaseExcelApp. Garbage collection failed.\r\n" +
                            ex.Message, Level.Error);
                    }
                }

                // Let's do one final check to make sure Excel closed
                try
                {
                    int pId = (int) excelProcessId;
                    Process process = Process.GetProcessById(pId);
                    process.Kill();
                }
                catch
                {
                    return;
                }
            }
        }


        /// <summary>
        /// Copies the passed source file path to a new temporary destination file in a specified folder. Uses the
        /// Internal.Framework.Storage.LocalFilePaths class.
        /// </summary>
        /// <param name="sourcePath">The full path to the source file.</param>
        /// <param name="destDirName">The destination directory name.</param>
        /// <param name="destFileName">The destination file name.</param>
        /// <returns></returns>
        public static string CopyWorkbook(string sourcePath, string destDirName, string destFileName)
        {
            string savePath = "";
            if (String.IsNullOrEmpty(sourcePath) || String.IsNullOrEmpty(destDirName) || String.IsNullOrEmpty(destFileName))return savePath;
            
            if (!File.Exists(sourcePath))
            {
                Logger.LogMessage("Error in call to WorksheetUtilitites.CopyWorkbook. Invalid source file path specified.", Level.Error);
                return "";
            }

            // Generate an random temp path to save new workbook
            TempDirectory tempDir = new TempDirectory(Path.Combine(LocalFilePaths.CommonStorageRoot, destDirName));
            savePath = Path.Combine(tempDir.Path, destFileName);

            tempDir.DeleteDirectory = false;
            tempDir.Dispose();
            try
            {
                File.Copy(sourcePath, savePath);
            }
            catch (Exception ex)
            {
                Logger.LogMessage("Copy of source file in WorksheetUtilitites.CopyWorkbook errored!\r\n" + ex.Message, Level.Error);
            }
            if (!File.Exists(savePath))
            {
                Logger.LogMessage("Copy of source file in WorksheetUtilitites.CopyWorkbook failed. Destination file does not exist!", Level.Error);
                savePath = "";
            }
            return savePath;
        }


        /// <summary>
        /// A "loose" method for checking if a sheet is password protected.
        /// </summary>
        /// <param name="sheet">The worksheet.</param>
        /// <returns>True if sheet is protected else false</returns>
        public static bool IsSheetPasswordProtected(_Worksheet sheet)
        {
            try
            {
                sheet.Unprotect(String.Empty);
                return false;
            }
            catch
            {
                return true;
            }
        }


        /// <summary>
        /// Protects or unprotects the passed sheet. Currently only supports setting of contents (e.g. cells).
        /// </summary>
        /// <param name="sheet">The sheet.</param>
        /// <param name="password">The password (can be null).</param>
        /// <param name="protect">False to unprotect. True to protect the sheet.</param>
        /// <returns>True if protection was unset or set successfully.</returns>
        public static bool SetSheetProtection(_Worksheet sheet, string password, bool protect)
        {
            if (sheet == null) return false;

            if (!protect)
            {
                try
                {
                    if (!String.IsNullOrEmpty(password))
                        sheet.Unprotect(password);
                    else
                        sheet.Unprotect(Type.Missing);
                }
                catch
                {
                    Logger.LogMessage("Protect of sheet failed in WorksheetUtilities.SetSheetProtection!", Level.Error);
                    return false;
                }
            }
            else
            {
                try
                {
                    if (!String.IsNullOrEmpty(password))
                        sheet.Protect(password, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    else
                        sheet.Protect(Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                }
                catch
                {
                    Logger.LogMessage("Protect of sheet failed in WorksheetUtilities.SetSheetProtection!", Level.Error);
                    return false;
                }
            }

            return true;
        }

        public static void ScrollToTopLeft(Worksheet sheet)
        {
            try
            {
                _app.Goto(sheet.Cells[1, 1], true);
            }
            catch
            {
                Logger.LogMessage("Scroll of sheet failed", Level.Error);
            }
        }

        /// <summary>
        /// Sets the border edges around the specified named range.
        /// </summary>
        /// <param name="sheet">The worksheet.</param>
        /// <param name="namedRangeName">The named range.</param>
        /// <param name="lineStyle">The line style.</param>
        /// <param name="borderWeight">The border (line) weight).</param>
        public static void SetNamedRangeBorderEdges(_Worksheet sheet, string namedRangeName, XlLineStyle lineStyle, XlBorderWeight borderWeight)
        {
            if (sheet == null || String.IsNullOrEmpty(namedRangeName)) return;

            Name namedRange = null;
            Range range = null;
            Border borderTop = null;
            Border borderBottom = null;
            Border borderLeft = null;
            Border borderRight = null;

            // Get the range by name
            object objNamedRange;
            try
            {
                objNamedRange = sheet.Names.Item(namedRangeName, Type.Missing, Type.Missing);
            }
            catch
            {
                objNamedRange = null;
                Logger.LogMessage(String.Format("Error executing WorksheetUtilities.SetNamedRangeBorderEdges. \r\nRange {0} doesn't exist in the sheet.", namedRangeName), Level.Error);
            }
            if (objNamedRange == null || !(objNamedRange is Name)) goto CleanUp_SetNamedRangeBorder;

            namedRange = objNamedRange as Name;
            range = namedRange.RefersToRange;

            borderTop = range.Borders.get_Item(XlBordersIndex.xlEdgeTop);
            borderTop.LineStyle = lineStyle;
            borderTop.Weight = borderWeight;
            borderBottom = range.Borders.get_Item(XlBordersIndex.xlEdgeBottom);
            borderBottom.LineStyle = lineStyle;
            borderBottom.Weight = borderWeight;
            borderLeft = range.Borders.get_Item(XlBordersIndex.xlEdgeLeft);
            borderLeft.LineStyle = lineStyle;
            borderLeft.Weight = borderWeight;
            borderRight = range.Borders.get_Item(XlBordersIndex.xlEdgeRight);
            borderRight.LineStyle = lineStyle;
            borderRight.Weight = borderWeight;

            CleanUp_SetNamedRangeBorder:
            {
                ReleaseComObject(borderTop);
                ReleaseComObject(borderBottom);
                ReleaseComObject(borderLeft);
                ReleaseComObject(borderRight);
                ReleaseComObject(range);
                ReleaseComObject(namedRange);
                ReleaseComObject(objNamedRange);

                // ReSharper disable RedundantAssignment
                sheet = null;
                // ReSharper restore RedundantAssignment
            }
        }

        /// <summary>
        /// Returns the row count for the named range.
        /// </summary>
        /// <param name="sheet">The worksheet.</param>
        /// <param name="namedRangeName">The named range.</param>
        /// <returns></returns>
        public static int GetNamedRangeRowCount(_Worksheet sheet, string namedRangeName)
        {
            int numRows = 0;

            if (sheet == null || String.IsNullOrEmpty(namedRangeName)) return numRows;

            Name namedRange = null;
            Range range = null;

            // Get the range by name
            object objNamedRange;
            try
            {
                objNamedRange = sheet.Names.Item(namedRangeName, Type.Missing, Type.Missing);
            }
            catch
            {
                objNamedRange = null;
                Logger.LogMessage(String.Format("Error executing WorksheetUtilities.GetNamedRangeRowCount. \r\nRange {0} doesn't exist in the sheet.", namedRangeName), Level.Error);
            }
            if (objNamedRange == null || !(objNamedRange is Name)) goto CleanUp_GetNamedRangeRowCount;

            namedRange = objNamedRange as Name;
            range = namedRange.RefersToRange;

            numRows = range.Rows.Count;

            CleanUp_GetNamedRangeRowCount:
            {
                ReleaseComObject(objNamedRange);
                ReleaseComObject(range);
                ReleaseComObject(namedRange);

                // ReSharper disable RedundantAssignment
                sheet = null;
                // ReSharper restore RedundantAssignment
            }

            return numRows;
        }

        public static int GetNamedRangeStartRow(Worksheet sheet, string rangeName)
        {
            return sheet.Range[rangeName].Row;
        }

        public static int GetNamedRangeStartColumn(Worksheet sheet, string rangeName)
        {
            return sheet.Range[rangeName].Column;
        }

        public static Range GetNamedRange(_Worksheet sheet, string namedRangeName)
        {
            Name namedRange = null;
            Range range = null;
            if (sheet == null || String.IsNullOrEmpty(namedRangeName)) return range;

            // Get the range by name
            object objNamedRange;
            try
            {
                objNamedRange = sheet.Names.Item(namedRangeName, Type.Missing, Type.Missing);
            }
            catch
            {
                objNamedRange = null;
                Logger.LogMessage(String.Format("Error executing WorksheetUtilities.GetNamedRangeRowCount. \r\nRange {0} doesn't exist in the sheet.", namedRangeName), Level.Error);
            }
            if (objNamedRange == null || !(objNamedRange is Name)) goto CleanUp_GetNamedRangeRowCount;

            namedRange = objNamedRange as Name;
            range = namedRange.RefersToRange;

           
        CleanUp_GetNamedRangeRowCount:
            {
                ReleaseComObject(objNamedRange);
                // ReleaseComObject(range);
                ReleaseComObject(namedRange);
                // ReSharper disable RedundantAssignment
                // sheet = null;
                // ReSharper restore RedundantAssignment
            }

            return range;
        }


        /// <summary>
        /// Inserts rows for a named range. Can fill the new range from a named range.
        /// </summary>
        /// <param name="sheet">The worksheet.</param>
        /// <param name="namedRangeName">The named range to use for the row inserts (i.e. inserts rows above that range since use EntireRow.Insert call)</param>
        /// <param name="numRows">The number of rows as offset to fill if fill is done.</param>
        /// <param name="fillFromRangeName">The range to fill from.</param>
        /// <param name="fillRows">Flag indicating whether to fill the newly inserted rows.</param>
        /// <param name="fillType">The fill type.</param>
        public static void InsertRowsForNamedRange(_Worksheet sheet, string namedRangeName, int numRows, string fillFromRangeName, bool fillRows, XlPasteType fillType)
        {
            if (sheet == null || numRows <= 0 || (fillRows && String.IsNullOrEmpty(fillFromRangeName))) return;

            Name namedRange = null;
            Range range = null;
            Range destRange = null;
            Name namedRangeFill = null;
            Range rangeFill = null;
            object objNamedRange;

            // Get the range by name
            try
            {
                objNamedRange = sheet.Names.Item(namedRangeName, Type.Missing, Type.Missing);
            }
            catch
            {
                objNamedRange = null;
                Logger.LogMessage(String.Format("Error executing WorksheetUtilities.InsertRowsForNamedRange. \r\nRange {0} doesn't exist in the sheet.", namedRangeName), Level.Error);
            }
            if (objNamedRange == null || !(objNamedRange is Name)) goto CleanUp_InsertRowsForNamedRange;

            namedRange = objNamedRange as Name;
            range = namedRange.RefersToRange;

            if (range != null)
            {
                range.EntireRow.Insert(Type.Missing, Type.Missing);
                //destRange = range.get_Offset(numRows, 0);
                //destRange.EntireRow.Insert(Type.Missing, Type.Missing);

                if (fillRows && !String.IsNullOrEmpty(fillFromRangeName))
                {
                    destRange = range.get_Offset(-numRows, 0);

                    object objFillRange;
                    try
                    {
                        objFillRange = sheet.Names.Item(fillFromRangeName, Type.Missing, Type.Missing);
                    }
                    catch
                    {
                        objFillRange = null;
                        Logger.LogMessage(String.Format("Error executing WorksheetUtilities.InsertRowsForNamedRange. \r\nRange {0} doesn't exist in the sheet.", fillFromRangeName), Level.Error);
                    }
                    if (objFillRange == null || !(objFillRange is Name)) goto CleanUp_InsertRowsForNamedRange;
                    namedRangeFill = objFillRange as Name;
                    rangeFill = namedRangeFill.RefersToRange;

                    rangeFill.Copy(Type.Missing);
                    destRange.PasteSpecial(fillType, XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                }
            }

            // The com objects must always be released
            CleanUp_InsertRowsForNamedRange:
            {
                ReleaseComObject(objNamedRange);
                ReleaseComObject(range);
                ReleaseComObject(namedRange);
                ReleaseComObject(destRange);
                ReleaseComObject(rangeFill);
                ReleaseComObject(namedRangeFill);

                // ReSharper disable RedundantAssignment
                sheet = null;
                // ReSharper restore RedundantAssignment
            }
        }

        /// <summary>
        /// Inserts rows for a named range after named range. Can fill the new range from a named range.
        /// </summary>
        /// <param name="sheet">The worksheet.</param>
        /// <param name="namedRangeName">The named range to use for the row inserts (i.e. inserts rows above that range since use EntireRow.Insert call)</param>
        /// <param name="numRows">The number of rows as offset to fill if fill is done.</param>
        /// <param name="fillFromRangeName">The range to fill from.</param>
        /// <param name="fillRows">Flag indicating whether to fill the newly inserted rows.</param>
        /// <param name="fillType">The fill type.</param>
        public static void InsertRowsAfterForNamedRange(_Worksheet sheet, string namedRangeName, int numRows, string fillFromRangeName, bool fillRows, XlPasteType fillType, string destNamedRange="", int rowPos =0)
        {
            if (sheet == null || numRows <= 0 || (fillRows && String.IsNullOrEmpty(fillFromRangeName))) return;

            Name namedRange = null;
            Range range = null;
            Range destRange = null;
            Name namedRangeFill = null;
            Range rangeFill = null;
            Range destRange2 = null;
            Range destRangeForName = null;
            object objNamedRange;

            // Get the range by name
            try
            {
                objNamedRange = sheet.Names.Item(namedRangeName, Type.Missing, Type.Missing);
            }
            catch
            {
                objNamedRange = null;
                Logger.LogMessage(String.Format("Error executing WorksheetUtilities.InsertRowsForNamedRange. \r\nRange {0} doesn't exist in the sheet.", namedRangeName), Level.Error);
            }
            if (objNamedRange == null || !(objNamedRange is Name)) goto CleanUp_InsertRowsForNamedRange;

            namedRange = objNamedRange as Name;
            range = namedRange.RefersToRange;
            if (range != null)
            {
                if (fillRows && !String.IsNullOrEmpty(fillFromRangeName))
                {
                    destRange = range.get_Offset(numRows-1, 0);
                    //if (rowPos == 0) 
                    destRange.EntireRow.Insert(XlInsertShiftDirection.xlShiftDown, Type.Missing);
                    //destRange[rowPos, 1].EntireRow.Insert(XlInsertShiftDirection.xlShiftDown, Type.Missing);

                    object objFillRange;
                    try
                    {
                        objFillRange = sheet.Names.Item(fillFromRangeName, Type.Missing, Type.Missing);
                    }
                    catch
                    {
                        objFillRange = null;
                        Logger.LogMessage(String.Format("Error executing WorksheetUtilities.InsertRowsForNamedRange. \r\nRange {0} doesn't exist in the sheet.", fillFromRangeName), Level.Error);
                    }
                    if (objFillRange == null || !(objFillRange is Name)) goto CleanUp_InsertRowsForNamedRange;
                    namedRangeFill = objFillRange as Name;
                    rangeFill = namedRangeFill.RefersToRange;

                    rangeFill.Copy(Type.Missing);
                    //if (rowPos == 0)
                    rowPos = -numRows + 2;
                    
                    destRange[rowPos, 1].PasteSpecial(fillType, XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                    if (!String.IsNullOrEmpty(destNamedRange))
                    {
                        destRangeForName = sheet.Range[sheet.Cells[destRange.Row - numRows + 1, destRange.Column], sheet.Cells[destRange.Row - 1, destRange.Column + destRange.Columns.Count - 1]].Cells;
                        //// Name the new range (pasted to)
                        ////destRangeForName = sheet.Application.Selection as Range;
                        //destRange2 = destRange.Cells[range.Rows.Count, range.Columns.Count] as Range;
                        ////string destRange2Address = destRange2.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                        //if (destRange2 != null)
                        //{
                        //    //destRangeForName = sheet.get_Range(destRange, destRange2);
                        //    //destRangeForName = (Range) sheet.get_Range(sheet.Cells[destRange.Row - numRows, destRange.Column], sheet.Cells[destRange.Row, destRange.Column + destRange.Columns.Count-1]);
                        //    destRangeForName = sheet.Range[sheet.Cells[destRange.Row - numRows+1, destRange.Column], sheet.Cells[destRange.Row-1, destRange.Column + destRange.Columns.Count - 1]].Cells;
                        //}
                        if (destRangeForName != null)
                        {
                            //string destNamedRangeAddress = destRangeForName.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

                            string refersToLocal = "='" + sheet.Name + "'!" + destRangeForName.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

                            sheet.Names.Add(destNamedRange, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing, refersToLocal, Type.Missing, Type.Missing, Type.Missing);
                        }
                    }
                }
            }

        // The com objects must always be released
        CleanUp_InsertRowsForNamedRange:
            {
                ReleaseComObject(objNamedRange);
                ReleaseComObject(range);
                ReleaseComObject(namedRange);
                ReleaseComObject(destRange);
                ReleaseComObject(rangeFill);
                ReleaseComObject(namedRangeFill);
                ReleaseComObject(destRange2);
                ReleaseComObject(destRangeForName);

                // ReSharper disable RedundantAssignment
                sheet = null;
                // ReSharper restore RedundantAssignment
            }
        }

        /// <summary>
        /// Inserts rows for a named range after named range. Can fill the new range from a named range.
        /// </summary>
        /// <param name="sheet">The worksheet.</param>
        /// <param name="namedRangeName">The named range to use for the row inserts (i.e. inserts rows above that range since use EntireRow.Insert call)</param>
        /// <param name="numRows">The number of rows as offset to fill if fill is done.</param>
        /// <param name="fillFromRangeName">The range to fill from.</param>
        /// <param name="fillRows">Flag indicating whether to fill the newly inserted rows.</param>
        /// <param name="fillType">The fill type.</param>
        public static void CopyNamedRangeToNewLocation(_Worksheet sheet, string srcNamedRangeName, string NameForTargetNamedRange, int rowOffset, int colOffset, XlPasteType fillType)
        {
            if (sheet == null || rowOffset <= 0 || String.IsNullOrEmpty(srcNamedRangeName) || String.IsNullOrEmpty(NameForTargetNamedRange)) return;

            Name namedRange = null;
            Range range = null;
            Range destRange = null;
            Name namedRangeFill = null;
            Range rangeFill = null;
            Range destRange2 = null;
            Range destRangeForName = null;
            object objNamedRange;

            // Get the range by name
            try
            {
                objNamedRange = sheet.Names.Item(srcNamedRangeName, Type.Missing, Type.Missing);
            }
            catch
            {
                objNamedRange = null;
                Logger.LogMessage(String.Format("Error executing WorksheetUtilities.InsertRowsForNamedRange. \r\nRange {0} doesn't exist in the sheet.", srcNamedRangeName), Level.Error);
            }
            if (objNamedRange == null || !(objNamedRange is Name)) goto CleanUp_InsertRowsForNamedRange;

            namedRange = objNamedRange as Name;
            range = namedRange.RefersToRange;
            if (range != null)
            {
                destRange = range.get_Offset(rowOffset, colOffset);
                destRange.EntireRow.Insert(XlInsertShiftDirection.xlShiftDown, Type.Missing);

                object objFillRange;
                try
                {
                    objFillRange = sheet.Names.Item(srcNamedRangeName, Type.Missing, Type.Missing);
                }
                catch
                {
                    objFillRange = null;
                    Logger.LogMessage(String.Format("Error executing WorksheetUtilities.InsertRowsForNamedRange. \r\nRange {0} doesn't exist in the sheet.", srcNamedRangeName), Level.Error);
                }
                if (objFillRange == null || !(objFillRange is Name)) goto CleanUp_InsertRowsForNamedRange;
                namedRangeFill = objFillRange as Name;
                rangeFill = namedRangeFill.RefersToRange;

                rangeFill.Copy(Type.Missing);

                destRange[0,1].PasteSpecial(fillType, XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                destRangeForName = sheet.Range[sheet.Cells[destRange.Row-1, destRange.Column], sheet.Cells[destRange.Row-1, destRange.Column + destRange.Columns.Count - 1]].Cells;

                if (destRangeForName != null)
                {
                    
                    string refersToLocal = "='" + sheet.Name + "'!" + destRangeForName.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

                    sheet.Names.Add(NameForTargetNamedRange, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, refersToLocal, Type.Missing, Type.Missing, Type.Missing);
                }

            }

        // The com objects must always be released
        CleanUp_InsertRowsForNamedRange:
            {
                ReleaseComObject(objNamedRange);
                ReleaseComObject(range);
                ReleaseComObject(namedRange);
                ReleaseComObject(destRange);
                ReleaseComObject(rangeFill);
                ReleaseComObject(namedRangeFill);
                ReleaseComObject(destRange2);
                ReleaseComObject(destRangeForName);

                // ReSharper disable RedundantAssignment
                sheet = null;
                // ReSharper restore RedundantAssignment
            }
        }

        /// <summary>
        /// Inserts a specified number of rows into a named range starting after the first row
        /// </summary>
        /// <param name="numRowsToInsert">The number of rows to insert</param>
        /// <param name="sheet">The sheet containing the named range</param>
        /// <param name="namedRangeName">The name of the named range</param>
        /// <param name="fillRows">Flag indicating whether to fill the new rows with content from the first row (e.g. formulas)</param>
        /// <param name="insertDirection">The direction to insert, only Up or Down are supported.</param>
        /// <param name="fillType">The fill type as the paste operation from the first row</param>
        public static void InsertRowsIntoNamedRange(int numRowsToInsert, _Worksheet sheet, string namedRangeName, bool fillRows, XlDirection insertDirection, XlPasteType fillType)
        {
            if (sheet == null || numRowsToInsert <= 0 || String.IsNullOrEmpty(namedRangeName)) return;

            List<Range> destRows = new List<Range>(0);
            Name namedRange = null;
            Range range = null;
            Range srcRow = null;
            Range row = null;
            Range destRow = null;
            object objNamedRange;
            
            // Get the range by name
            try
            {
                objNamedRange = sheet.Names.Item(namedRangeName, Type.Missing, Type.Missing);
            }
            catch
            {
                objNamedRange = null;
                Logger.LogMessage(String.Format("Error executing WorksheetUtilities.InsertRowsIntoNamedRange. \r\nRange {0} doesn't exist in the sheet.", namedRangeName), Level.Error);
            }
            if (!(objNamedRange is Name)) goto Cleanup_InsertRowsIntoNamedRange;

            namedRange = objNamedRange as Name;
            range = namedRange.RefersToRange;

            if (range != null)
            {
                switch (insertDirection)
                {
                    case XlDirection.xlDown:
                        srcRow = range.Rows[1, Type.Missing] as Range;
                        break;
                    case XlDirection.xlUp:
                        srcRow = range.Rows[range.Rows.Count, Type.Missing] as Range;
                        break;
                }

                for (int i = numRowsToInsert; i > 0; i--)
                {
                    switch (insertDirection)
                    {
                        case XlDirection.xlDown:
                            row = range.Rows[2, Type.Missing] as Range;
                            break;
                        case XlDirection.xlUp:
                            row = range.Rows[range.Rows.Count, Type.Missing] as Range;
                            break;
                        default:
                            continue;
                    }
                    if (row == null) continue;

                    row.EntireRow.Insert(Type.Missing, Type.Missing);
                    destRow = row.get_Offset(-1, 0);

                    destRows.Add(destRow);

                    // Clean up
                    ReleaseComObject(row);
                }

                // Fill the rows AFTER inserting so to get the correct fill (e.g. formulas)
                if (destRows.Count > 0 && fillRows)
                {
                    if (srcRow != null)
                    {
                        foreach (Range destRowFromList in destRows)
                        {
                            srcRow.Copy(Type.Missing);
                            destRowFromList.PasteSpecial(fillType, XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                        }

                        // Clean up
                        ReleaseComObject(srcRow);
                    }
                }
            }

            // The com objects must always be released
            Cleanup_InsertRowsIntoNamedRange:
            {
                try
                {
                    ReleaseComObject(objNamedRange);
                    ReleaseComObject(namedRange);
                    ReleaseComObject(range);
                    ReleaseComObject(srcRow);
                    ReleaseComObject(row);
                    ReleaseComObject(destRow);

                    if (destRows.Count > 0)
                    {
                        foreach (var dest in destRows)
                        {
                            ReleaseComObject(dest);
                        }
                        destRows.Clear();
                    }

                    // ReSharper disable RedundantAssignment
                    sheet = null;
                    // ReSharper restore RedundantAssignment
                }
                catch
                {
                    return;
                }
            }
        }

        /// <summary>
        /// Inserts a specified number of rows into a named range starting from the parameter Row. Direction must be xlDown
        /// </summary>
        /// <param name="numRowsToInsert">The number of rows to insert</param>
        /// <param name="sheet">The sheet containing the named range</param>
        /// <param name="namedRangeName">The name of the named range</param>
        /// <param name="fillRows">Flag indicating whether to fill the new rows with content from the first row (e.g. formulas)</param>
        /// <param name="insertDirection">The direction to insert, only Up or Down are supported.</param>
        /// <param name="fillType">The fill type as the paste operation from the first row</param>
        /// <param name="rowIndex">row from where the insert starts</param>
        public static void InsertRowsIntoNamedRangeFromRow(int numRowsToInsert, _Worksheet sheet, string namedRangeName, bool fillRows, XlDirection insertDirection, XlPasteType fillType,int rowIndex)
        {
            if (sheet == null || numRowsToInsert <= 0 || String.IsNullOrEmpty(namedRangeName)) return;

            List<Range> destRows = new List<Range>(0);
            Name namedRange = null;
            Range range = null;
            Range srcRow = null;
            Range row = null;
            Range destRow = null;
            object objNamedRange;

            // Get the range by name
            try
            {
                objNamedRange = sheet.Names.Item(namedRangeName, Type.Missing, Type.Missing);
            }
            catch
            {
                objNamedRange = null;
                Logger.LogMessage(String.Format("Error executing WorksheetUtilities.InsertRowsIntoNamedRange. \r\nRange {0} doesn't exist in the sheet.", namedRangeName), Level.Error);
            }
            if (!(objNamedRange is Name)) goto Cleanup_InsertRowsIntoNamedRange;

            namedRange = objNamedRange as Name;
            range = namedRange.RefersToRange;

            if (range != null)
            {
                switch (insertDirection)
                {
                    case XlDirection.xlDown:
                        srcRow = range.Rows[rowIndex, Type.Missing] as Range;
                        break;
                    case XlDirection.xlUp:
                        srcRow = range.Rows[range.Rows.Count, Type.Missing] as Range;
                        break;
                }

                for (int i = numRowsToInsert; i > 0; i--)
                {
                    switch (insertDirection)
                    {
                        case XlDirection.xlDown:
                            row = range.Rows[rowIndex, Type.Missing] as Range;
                            break;
                        case XlDirection.xlUp:
                            row = range.Rows[range.Rows.Count, Type.Missing] as Range;
                            break;
                        default:
                            continue;
                    }
                    if (row == null) continue;

                    row.EntireRow.Insert(Type.Missing, Type.Missing);
                    destRow = row.get_Offset(-1, 0);                    

                    destRows.Add(destRow);

                    // Clean up
                    ReleaseComObject(row);
                }

                // Fill the rows AFTER inserting so to get the correct fill (e.g. formulas)
                if (destRows.Count > 0 && fillRows)
                {
                    if (srcRow != null)
                    {
                        foreach (Range destRowFromList in destRows)
                        {
                            srcRow.Copy(Type.Missing);
                            destRowFromList.PasteSpecial(fillType, XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                        }

                        // Clean up
                        ReleaseComObject(srcRow);
                    }
                }
            }

        // The com objects must always be released
        Cleanup_InsertRowsIntoNamedRange:
            {
                try
                {
                    ReleaseComObject(objNamedRange);
                    ReleaseComObject(namedRange);
                    ReleaseComObject(range);
                    ReleaseComObject(srcRow);
                    ReleaseComObject(row);
                    ReleaseComObject(destRow);

                    if (destRows.Count > 0)
                    {
                        for (int i = 0; i < destRows.Count; i++)
                        {
                            ReleaseComObject(destRows[i]);
                        }
                        destRows.Clear();
                    }

                    // ReSharper disable RedundantAssignment
                    sheet = null;
                    // ReSharper restore RedundantAssignment
                }
                catch
                {
                    return;
                }
            }

        }

        /// <summary>
        /// Deletes a specified number of rows from a named range. Shift is up by default.
        /// </summary>
        /// <param name="numRowsToRemove">The number of rows to remove. Should be less than the number of rows in the range!</param>
        /// <param name="sheet">The sheet containing the named range.</param>
        /// <param name="namedRangeName">The name of the named range to delete the row(s) from.</param>
        /// <param name="insertDirection">The direction in which to delete (only xlUp and xlDown supported!)</param>
        public static void DeleteRowsFromNamedRange(int numRowsToRemove, _Worksheet sheet, string namedRangeName, XlDirection insertDirection)
        {
            if (sheet == null || String.IsNullOrEmpty(namedRangeName)) return;

            Name namedRange = null;
            Range range = null;
            Range row;
            object objNamedRange;

            // Get the range by name
            try
            {
                objNamedRange = sheet.Names.Item(namedRangeName, Type.Missing, Type.Missing);
            }
            catch
            {
                objNamedRange = null;
                Logger.LogMessage(String.Format("Error executing WorksheetUtilities.DeleteRowsFromNamedRange. \r\nRange {0} doesn't exist in the sheet.", namedRangeName), Level.Error);
            }
            if (!(objNamedRange is Name)) goto Cleanup_DeleteRowsFromNamedRange;

            namedRange = objNamedRange as Name;
            range = namedRange.RefersToRange;

            if (range != null)
            {
                for (int i = numRowsToRemove; i > 0; i--)
                {
                    // Remove from range based on the direction
                    switch (insertDirection)
                    {
                        case XlDirection.xlDown:
                            row = range.Rows[1, Type.Missing] as Range;
                            break;
                        case XlDirection.xlUp:
                            row = range.Rows[range.Rows.Count, Type.Missing] as Range;
                            break;
                        default:
                            continue;
                    }

                    if (row == null) continue;

                    //MessageBox.Show("row address = " + row.get_Address(Type.Missing, Type.Missing, XlReferenceStyle.xlA1, Type.Missing, Type.Missing));
                    try
                    {
                        row.EntireRow.Delete(XlDeleteShiftDirection.xlShiftUp);
                    }
                    catch (Exception ex)
                    {
                        Logger.LogMessage("Error executing WorksheetUtilities.DeleteRowsFromNamedRange. " + ex.Message, Level.Error);
                    }

                    // Clean up
                    ReleaseComObject(row);
                }
            }

            // The com objects must always be released
            Cleanup_DeleteRowsFromNamedRange:
            {
                try
                {
                    ReleaseComObject(objNamedRange);
                    ReleaseComObject(range);
                    ReleaseComObject(namedRange);

                    // ReSharper disable RedundantAssignment
                    sheet = null;
                    // ReSharper restore RedundantAssignment
                }
                catch
                {
                    return;
                }
            }
        }


        /// <summary>
        /// Inserts a specified number of columns into a named range starting at beginning or end of the range (as defined by the direction).
        /// </summary>
        /// <param name="numColsToInsert">The number of columns to insert</param>
        /// <param name="sheet">The sheet containing the named range</param>
        /// <param name="namedRangeName">The name of the named range</param>
        /// <param name="insertDirection">The direction to insert, only Left or Right are supported.</param>
        public static void InsertColumnsIntoNamedRange(int numColsToInsert, _Worksheet sheet, string namedRangeName, XlDirection insertDirection)
        {
            if (sheet == null || numColsToInsert <= 0 || String.IsNullOrEmpty(namedRangeName)) return;

            Name namedRange = null;
            Range range = null;
            Range col = null;
            object objNamedRange;

            // Get the range by name
            try
            {
                objNamedRange = sheet.Names.Item(namedRangeName, Type.Missing, Type.Missing);
            }
            catch
            {
                objNamedRange = null;
                Logger.LogMessage(String.Format("Error executing WorksheetUtilities.InsertColumnsIntoNamedRange. \r\nRange {0} doesn't exist in the sheet.", namedRangeName), Level.Error);
            }
            if (!(objNamedRange is Name)) goto Cleanup_InsertRowsIntoNamedRange;

            namedRange = objNamedRange as Name;
            range = namedRange.RefersToRange;

            if (range != null)
            {
                for (int i = numColsToInsert; i > 0; i--)
                {
                    switch (insertDirection)
                    {
                        case XlDirection.xlToLeft:
                            col = range.Columns[1, Type.Missing] as Range;
                            break;
                        case XlDirection.xlToRight:
                            col = range.Columns[range.Columns.Count, Type.Missing] as Range;
                            break;
                        default:
                            continue;
                    }
                    if (col == null) continue;

                    //string address = col.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                    col.EntireColumn.Insert(Type.Missing, Type.Missing);

                    // Clean up
                    ReleaseComObject(col);
                }
            }

            // The com objects must always be released
            Cleanup_InsertRowsIntoNamedRange:
            {
                try
                {
                    ReleaseComObject(objNamedRange);
                    ReleaseComObject(namedRange);
                    ReleaseComObject(range);
                    ReleaseComObject(col);

                    // ReSharper disable RedundantAssignment
                    sheet = null;
                    // ReSharper restore RedundantAssignment
                }
                catch
                {
                    return;
                }
            }
        }

        /// <summary>
        /// Deletes a named range from the sheet or workbook. Does NOT delete the cells the named range names.
        /// </summary>
        /// <param name="sheet">The worksheet.</param>
        /// <param name="namedRangeName">The name of the named range.</param>
        public static void DeleteNamedRange(_Worksheet sheet, string namedRangeName)
        {
            if (sheet == null || String.IsNullOrEmpty(namedRangeName)) return;

            Names names = null;
            object wbkObj = null;
            _Workbook wbook = null;

            try
            {
                // Try a delete at sheet level first
                names = sheet.Names;
                names.Item(namedRangeName, Type.Missing, Type.Missing).Delete();
            }
            catch
            {
                try
                {
                    // Perhaps the name is at the workbook level
                    wbkObj = sheet.Names.Parent;
                    if (wbkObj is _Workbook)
                    {
                        wbook = wbkObj as _Workbook;
                        names = wbook.Names;
                        names.Item(namedRangeName, Type.Missing, Type.Missing).Delete();
                    }
                }
                catch
                {
                    goto Cleanup_DeleteNamedRange;
                }
            }

            // The com objects must always be released
            Cleanup_DeleteNamedRange:
            {
                try
                {
                    ReleaseComObject(names);
                    ReleaseComObject(wbook);
                    ReleaseComObject(wbkObj);

                    // ReSharper disable RedundantAssignment
                    sheet = null;
                    // ReSharper restore RedundantAssignment
                }
                catch
                {
                    return;
                }
            }

        }

        public static void DeleteInvalidNamedRanges(Worksheet sheet)
        {
            var ranges = sheet.Parent.Names; // Get the named ranges from the workbook
            for (int i = ranges.Count; i >= 1; i--) // Iterate in reverse to avoid skipping items after deletion
            {
                var currentName = ranges.Item(i, Type.Missing, Type.Missing);
                var refersTo = currentName.RefersTo.ToString();
                if (refersTo.Contains("REF!"))
                {
                    DeleteNamedRange(sheet, currentName.Name);
                }
            }
        }

        /// <summary>
        /// Moves a named range within the sheet (includes moving the range contents).
        /// </summary>
        /// <param name="sheet">The worksheet.</param>
        /// <param name="namedRangeName">The name of the named range.</param>
        /// <param name="rowOffset">The row offset from the sheet to move the range to.</param>
        /// <param name="colOffset">The column offset from the sheet to move the range to.</param>
        public static void MoveNamedRange(_Worksheet sheet, string namedRangeName, int rowOffset, int colOffset)
        {
            if (sheet == null || String.IsNullOrEmpty(namedRangeName)) return;

            Name namedRange = null;
            Range range = null;
            object objNamedRange;
            Range destRange = null;

            // Get the range by name
            try
            {
                objNamedRange = sheet.Names.Item(namedRangeName, Type.Missing, Type.Missing);
            }
            catch
            {
                objNamedRange = null;
                Logger.LogMessage(String.Format("Error executing WorksheetUtilities.MoveNamedRange. \r\nRange {0} doesn't exist in the sheet.", namedRangeName), Level.Error);
            }
            if (!(objNamedRange is Name)) goto Cleanup_MoveNamedRange;
            namedRange = objNamedRange as Name;
            range = namedRange.RefersToRange;

            destRange = sheet.Cells[rowOffset, colOffset] as Range;
            //string destAddress = destRange.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

            range.Cut(destRange);

            // The com objects must always be released
           Cleanup_MoveNamedRange:
            {
                try
                {
                    ReleaseComObject(destRange);
                    ReleaseComObject(range);
                    ReleaseComObject(namedRange);
                    ReleaseComObject(objNamedRange);

                    // ReSharper disable RedundantAssignment
                    sheet = null;
                    // ReSharper restore RedundantAssignment
                }
                catch
                {
                    return;
                }
            }

        }


        /// <summary>
        /// Deletes a named range rows from a sheet. Essentially deletes the named range.
        /// </summary>
        /// <param name="sheet">The sheet containing the named range.</param>
        /// <param name="namedRangeName">The name of the range containing the rows to delete.</param>
        public static void DeleteNamedRangeRows(_Worksheet sheet, string namedRangeName)
        {
            if (sheet == null || String.IsNullOrEmpty(namedRangeName)) return;

            Name namedRange = null;
            Range range = null;
            object objNamedRange;

            // Get the range by name
            try
            {
                objNamedRange = sheet.Names.Item(namedRangeName, Type.Missing, Type.Missing);
            }
            catch
            {
                objNamedRange = null;
                Logger.LogMessage(String.Format("Error executing WorksheetUtilities.DeleteNamedRangeRows. \r\nRange {0} doesn't exist in the sheet.", namedRangeName), Level.Error);
            }
            if (!(objNamedRange is Name)) goto Cleanup_DeleteNamedRangeRows;
            namedRange = objNamedRange as Name;
            range = namedRange.RefersToRange;

            try
            {
                range.EntireRow.Delete(XlDeleteShiftDirection.xlShiftUp);
            }
            catch( Exception ex)
            {
                Logger.LogMessage("Error executing WorksheetUtilities.DeleteNamedRangeRows.\r\n" + ex.Message, Level.Error);
            }

            // The com objects must always be released
            Cleanup_DeleteNamedRangeRows:
            {
                try
                {
                    ReleaseComObject(objNamedRange);
                    ReleaseComObject(range);
                    ReleaseComObject(namedRange);

                    // ReSharper disable RedundantAssignment
                    sheet = null;
                    // ReSharper restore RedundantAssignment
                }
                catch
                {
                    return;
                }
            }
        }


        /// <summary>
        /// Inserts a specified number of rows into a named range that has a single row defined. Insert direction is always down.
        /// </summary>
        /// <param name="numRowsToInsert">The number of rows to insert</param>
        /// <param name="sheet">The sheet containing the named range</param>
        /// <param name="namedRangeName">The name of the named range</param>
        /// <param name="fillRows">Flag indicating whether to fill the new rows with content from the first row (e.g. formulas)</param>
        /// <param name="fillType">The fill type as the paste operation from the first row</param>
        public static void InsertRowsIntoSingleRowedNamedRange(int numRowsToInsert, _Worksheet sheet, string namedRangeName, bool fillRows, XlPasteType fillType)
        {
            if (sheet == null || numRowsToInsert <= 0 || String.IsNullOrEmpty(namedRangeName)) return;

            List<Range> destRows = new List<Range>(0);
            Name namedRange = null;
            Range range = null;
            Range srcRow = null;
            Range row = null;
            Range destRow = null;
            Range startCell = null;
            Range endCell = null;
            object objNamedRange;

            // Get the range by name
            try
            {
                objNamedRange = sheet.Names.Item(namedRangeName, Type.Missing, Type.Missing);
            }
            catch
            {
                objNamedRange = null;
                Logger.LogMessage(String.Format("Error executing WorksheetUtilities.InsertRowsIntoSingleRowedNamedRange. \r\nRange {0} doesn't exist in the sheet.", namedRangeName), Level.Error);
            }
            if (!(objNamedRange is Name)) goto Cleanup_InsertRowsIntoNamedRange;

            namedRange = objNamedRange as Name;
            range = namedRange.RefersToRange;

            if (range != null)
            {
                srcRow = range.Rows[1, Type.Missing] as Range;

                for (int i = numRowsToInsert; i > 0; i--)
                {
                    // Insert at the second row
                    row = range.Rows[2, Type.Missing] as Range;
                    if (row == null) continue;

                    //string rowAddress = row.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

                    row.EntireRow.Insert(Type.Missing, Type.Missing);
                    destRow = row.get_Offset(-1, 0);

                    destRows.Add(destRow);

                    // Clean up
                    ReleaseComObject(row);
                }

                // Update the named range to point to the new address (since single rowed named ranges don't expand dynamically)
                startCell = namedRange.RefersToRange.Cells[1, 1] as Range;
                endCell = namedRange.RefersToRange.Cells[numRowsToInsert + 1, range.Columns.Count] as Range;
                if (startCell != null && endCell != null)
                {
                    //string startCellAddress = startCell.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                    //string endCellAddress = endCell.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

                    string refersToLocal = "='" + sheet.Name + "'!" +
                                     startCell.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing) + ":" +
                                     endCell.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

                    // Update the range
                    sheet.Names.Add(namedRangeName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, refersToLocal, Type.Missing, Type.Missing, Type.Missing);
                    //namedRange.RefersToLocal = refersToLocal;
                }


                // Fill the rows AFTER inserting so to get the correct fill (e.g. formulas)
                if (destRows.Count > 0 && fillRows)
                {
                    if (srcRow != null)
                    {
                        foreach (Range destRowFromList in destRows)
                        {
                            srcRow.Copy(Type.Missing);
                            destRowFromList.PasteSpecial(fillType, XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                        }

                        // Clean up
                        ReleaseComObject(srcRow);
                    }
                }
            }

            // The com objects must always be released
            Cleanup_InsertRowsIntoNamedRange:
            {
                try
                {
                    ReleaseComObject(objNamedRange);
                    ReleaseComObject(namedRange);
                    ReleaseComObject(range);
                    ReleaseComObject(srcRow);
                    ReleaseComObject(row);
                    ReleaseComObject(destRow);

                    if (destRows.Count > 0)
                    {
                        foreach (var dest in destRows)
                        {
                            ReleaseComObject(dest);
                        }
                        destRows.Clear();
                    }

                    ReleaseComObject(endCell);
                    ReleaseComObject(startCell);

                    // ReSharper disable RedundantAssignment
                    sheet = null;
                    // ReSharper restore RedundantAssignment
                }
                catch
                {
                    return;
                }
            }

        }

        /// <summary>
        /// Deletes a row within a specified range.
        /// </summary>
        /// <param name="sheet">The sheet containing the named range.</param>
        /// <param name="namedRangeName">The name of the range.</param>
        /// <param name="rowNum">The row number within the range to delete.</param>
        public static void DeleteRowFromNamedRange(_Worksheet sheet, string namedRangeName, int rowNum)
        {
            if (sheet == null || String.IsNullOrEmpty(namedRangeName) || rowNum < 1) return;

            Name namedRange = null;
            Range range = null;
            Range row = null;
            object objNamedRange;

            // Get the range by name
            try
            {
                objNamedRange = sheet.Names.Item(namedRangeName, Type.Missing, Type.Missing);
            }
            catch
            {
                objNamedRange = null;
                Logger.LogMessage(String.Format("Error executing WorksheetUtilities.DeleteRowFromNamedRange. \r\nRange {0} doesn't exist in the sheet.", namedRangeName), Level.Error);
            }
            if (!(objNamedRange is Name)) goto CleanUp_DeleteRowFromNamedRange;
            namedRange = objNamedRange as Name;
            range = namedRange.RefersToRange;
            
            try
            {
                row = range.Rows[rowNum, Type.Missing] as Range;
                if (row == null) goto CleanUp_DeleteRowFromNamedRange;

                row.EntireRow.Delete(XlDeleteShiftDirection.xlShiftUp);

                ReleaseComObject(row);
            }
            catch (Exception ex)
            {
                Logger.LogMessage("Error executing WorksheetUtilities.DeleteRowFromNamedRange. \r\n" + ex.Message, Level.Error);
            }

            // The com objects must always be released
            CleanUp_DeleteRowFromNamedRange:
            {
                try
                {
                    ReleaseComObject(objNamedRange);
                    ReleaseComObject(range);
                    ReleaseComObject(namedRange);
                    ReleaseComObject(row);

                    // ReSharper disable RedundantAssignment
                    sheet = null;
                    // ReSharper restore RedundantAssignment
                }
                catch
                {
                    return;
                }
            }
        }

        public static void DeleteSheet(Worksheet sheetToDelete)
        {
            if (sheetToDelete != null)
            {
                _app.DisplayAlerts = false;
                sheetToDelete.Delete();
                _app.DisplayAlerts = true;
            }
        }

        /// <summary>
        /// Replaces ALL occurrences of a string within a sheet.
        /// </summary>
        /// <param name="sheet">The sheet in which to perform the replacement</param>
        /// <param name="toReplace">The string to replace.</param>
        /// <param name="replaceWith">The replacement string.</param>
        public static void ReplaceInSheet(_Worksheet sheet, string toReplace, string replaceWith)
        {
            if (sheet == null || String.IsNullOrEmpty(toReplace) || String.IsNullOrEmpty(replaceWith)) return;

            try
            {
                sheet.Cells.Replace(toReplace, replaceWith, XlLookAt.xlPart, XlSearchOrder.xlByRows, false, false, false, false);
            }
            catch (Exception ex)
            {
                Logger.LogMessage("Error executing WorksheetUtilities.ReplaceInSheet. \r\n" + ex.Message, Level.Error);
                return;
            }
        }

        public static void PostProcessSheet(Worksheet sheet)
        {
            DeleteInvalidNamedRanges(sheet);

            // sheet.Columns.AutoFit();

            ScrollToTopLeft(sheet);

            SetSheetProtection(sheet, null, true);

            ReleaseComObject(sheet);
        }

        public static void ReplicateSheetAndDeleteOriginal(Workbook book, Worksheet originalSheet, int numSamples)
        {
            if (originalSheet == null || numSamples <= 0)
                return;

            string baseName = originalSheet.Name;

            for (int i = 1; i <= numSamples; i++)
            {
                // Copy the original sheet to the end of the workbook
                originalSheet.Copy(After: book.Sheets[book.Sheets.Count]);
                Worksheet copiedSheet = (Worksheet)book.Sheets[book.Sheets.Count];

                // Unprotect if needed
                bool wasProtected = SetSheetProtection(copiedSheet, null, false);

                // Rename the copied sheet
                copiedSheet.Name = $"{baseName} {i}";

                // Optional: scroll to top-left
                ScrollToTopLeft(copiedSheet);

                // Re-protect if it was originally protected
                if (wasProtected)
                    SetSheetProtection(copiedSheet, null, true);

                ReleaseComObject(copiedSheet);
            }

            // Delete the original sheet
            DeleteSheet(originalSheet);
            ReleaseComObject(originalSheet);
        }

        public static string ProcessSampleInfo(Worksheet sheet, string strcmbProductType, int numSamples)
        {
            const int DefaultNumRowsSampleInfo = 2;

            // delete unused Sample Info sections.
            List<string> allSections = new List<string>
            {
                "DrugSubstanceSection",
                "SDDSection",
                "DrugProductSection"
            };

            // Determine which section to keep
            var sectionMap = new Dictionary<string, string>
            {
                { "Drug Substance", "DrugSubstanceSection" },
                { "SDD", "SDDSection" },
                { "Drug Product", "DrugProductSection" }
            };

            string sectionToKeep = sectionMap.TryGetValue(strcmbProductType, out var section) ? section : null;

            // Delete all other sections
            foreach (string sectionToDel in allSections)
            {
                if (sectionToDel != sectionToKeep)
                {
                    DeleteNamedRangeRows(sheet, sectionToDel);
                }
            }

            string rowsToKeep = sectionToKeep?.Replace("Section", "Rows");

            if (numSamples > DefaultNumRowsSampleInfo)
            {
                InsertRowsIntoNamedRange(numSamples - 2, sheet, rowsToKeep, true, XlDirection.xlUp, XlPasteType.xlPasteAll);
            }
            else if (numSamples < DefaultNumRowsSampleInfo)
            {
                DeleteRowsFromNamedRange(1, sheet, rowsToKeep, XlDirection.xlDown);
            }

            // set row numbers
            List<string> list = new List<string>(0);
            for (int i = 1; i <= numSamples; i++)
            {
                list.Add(i.ToString());
            }

            SetNamedRangeValues(sheet, sectionToKeep?.Replace("Section", "RowNums"), list);

            return rowsToKeep;
        }

        /// <summary>
        /// Updates chart category axis title by replacing the specified string with a new string
        /// </summary>
        /// <param name="sheet">The sheet that contains the chart.</param>
        /// <param name="chartName">The name of the chart to update.</param>
        /// <param name="replace">The string to replace.</param>
        /// <param name="replaceWith">The replacement string.</param>
        public static void UpdateChartCategoryAxisTitle(_Worksheet sheet, string chartName, string replace, string replaceWith)
        {
            if (sheet == null || String.IsNullOrEmpty(chartName) || String.IsNullOrEmpty(replace) || String.IsNullOrEmpty(replaceWith)) return;

            object chartObj;

            try
            {
                chartObj = sheet.ChartObjects(chartName);
            }
            catch
            {
                chartObj = null;
                Logger.LogMessage(String.Format("Error executing WorksheetUtilities.UpdateChartCategoryAxisTitle. \r\nChart '{0}' doesn't exist in the sheet.", chartName), Level.Error);
            }
            if (chartObj != null && chartObj is ChartObject)
            {
                try
                {
                    ChartObject chartObject = chartObj as ChartObject;
                    Chart chart = chartObject.Chart;
                    Axis axis = chart.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary) as Axis;
                    if (axis == null) return;
                    axis.AxisTitle.Text = axis.AxisTitle.Text.Replace(replace, replaceWith);

                    // Clean up
                    ReleaseComObject(axis);
                    ReleaseComObject(chart);
                    ReleaseComObject(chartObject);
                }
                catch (Exception ex)
                {
                    Logger.LogMessage("Error executing WorksheetUtilities.UpdateChartCategoryAxisTitle.\r\n" + ex.Message, Level.Error);
                }
                // ReSharper disable RedundantAssignment
                sheet = null;
                // ReSharper restore RedundantAssignment
            }
        }



        /// <summary>
        /// Sets a list of values into a named range. Expects that passed range is a singe dimensional list of cells (i.e. a single column of cells)
        /// Note: This method is likely limited to filling named ranges that are a column of cells. See SetNamedRangeValues2 method.
        /// </summary>
        /// <param name="sheet">The sheet containing the named range.</param>
        /// <param name="namedRangeName">The name of the named range.</param>
        /// <param name="valList">The list of values (as strings).</param>
        public static void SetNamedRangeValues(_Worksheet sheet, string namedRangeName, List<string> valList)
        {
            if (sheet == null || String.IsNullOrEmpty(namedRangeName) || valList.Count <= 0) return;

            Name name = null;
            Range range = null;
            Range row = null;
            object objNamedRange;

            // Get the range by name
            try
            {
                objNamedRange = sheet.Names.Item(namedRangeName, Type.Missing, Type.Missing);
            }
            catch
            {
                objNamedRange = null;
                Logger.LogMessage(String.Format("Error executing WorksheetUtilities.SetNamedRangeValues. \r\nRange {0} doesn't exist in the sheet.", namedRangeName), Level.Error);
            }
            if (objNamedRange == null || !(objNamedRange is Name)) goto CleanUp_SetNamedRangeValues;
            name = objNamedRange as Name;
            range = name.RefersToRange;

            int i = 1;
            foreach (string val in valList)
            {


                if (i > range.Rows.Count) break;
                row = range.Rows[i, Type.Missing] as Range;
                try
                {
                    if (row != null) row.Value2 = val;
                }
                catch (Exception ex)
                {
                    Logger.LogMessage("Error executing WorksheetUtilities.SetNamedRangeValues to set cell value. \r\n" + ex.Message, Level.Error);
                }

                // Clean up
                ReleaseComObject(row);
                i++;
            }

            // The com objects must always be released
            CleanUp_SetNamedRangeValues:
            {
                try
                {
                    ReleaseComObject(objNamedRange);
                    ReleaseComObject(range);
                    ReleaseComObject(name);
                    ReleaseComObject(row);

                    // ReSharper disable RedundantAssignment
                    sheet = null;
                    // ReSharper restore RedundantAssignment
                }
                catch
                {
                    return;
                }
            }
        }

        public static void SetNamedRangeValuesByCol(_Worksheet sheet, string namedRangeName, List<string> valList, int colIndex)
        {
            if (sheet == null || String.IsNullOrEmpty(namedRangeName) || valList.Count <= 0) return;

            Name name = null;
            Range range = null;
            Range cell = null;
            object objNamedRange;

            // Get the range by name
            try
            {
                objNamedRange = sheet.Names.Item(namedRangeName, Type.Missing, Type.Missing);
            }
            catch
            {
                objNamedRange = null;
                Logger.LogMessage(String.Format("Error executing WorksheetUtilities.SetNamedRangeValues. \r\nRange {0} doesn't exist in the sheet.", namedRangeName), Level.Error);
            }
            if (objNamedRange == null || !(objNamedRange is Name)) goto CleanUp_SetNamedRangeValues;
            name = objNamedRange as Name;
            range = name.RefersToRange;

            int i = 1;
            foreach (string val in valList)
            {
                if (i > range.Rows.Count) break;
                cell = range.Cells[i, colIndex] as Range;
                try
                {
                    if (cell != null) cell.Value2 = val;
                }
                catch (Exception ex)
                {
                    Logger.LogMessage("Error executing WorksheetUtilities.SetNamedRangeValues to set cell value. \r\n" + ex.Message, Level.Error);
                }

                // Clean up
                ReleaseComObject(cell);
                i++;
            }

           // The com objects must always be released
        CleanUp_SetNamedRangeValues:
            {
                try
                {
                    ReleaseComObject(objNamedRange);
                    ReleaseComObject(range);
                    ReleaseComObject(name);
                    ReleaseComObject(cell);

                    // ReSharper disable RedundantAssignment
                    sheet = null;
                    // ReSharper restore RedundantAssignment
                }
                catch
                {
                    return;
                }
            }

        }

        public static void SetNamedRangeValuesByRow(_Worksheet sheet, string namedRangeName, List<string> valList, int rowIndex)
        {
            if (sheet == null || String.IsNullOrEmpty(namedRangeName) || valList.Count <= 0) return;

            Name name = null;
            Range range = null;
            Range cell = null;
            object objNamedRange;

            // Get the range by name
            try
            {
                objNamedRange = sheet.Names.Item(namedRangeName, Type.Missing, Type.Missing);
            }
            catch
            {
                objNamedRange = null;
                Logger.LogMessage(String.Format("Error executing WorksheetUtilities.SetNamedRangeValues. \r\nRange {0} doesn't exist in the sheet.", namedRangeName), Level.Error);
            }
            if (objNamedRange == null || !(objNamedRange is Name)) goto CleanUp_SetNamedRangeValues;
            name = objNamedRange as Name;
            range = name.RefersToRange;

            int i = 1;
            foreach (string val in valList)
            {
                if (i > range.Columns.Count) break;
                cell = range.Cells[rowIndex, i] as Range;
                try
                {
                    if (cell != null) cell.Value2 = val;
                }
                catch (Exception ex)
                {
                    Logger.LogMessage("Error executing WorksheetUtilities.SetNamedRangeValues to set cell value. \r\n" + ex.Message, Level.Error);
                }

                // Clean up
                ReleaseComObject(cell);
                i++;
            }

          // The com objects must always be released
        CleanUp_SetNamedRangeValues:
            {
                try
                {
                    ReleaseComObject(objNamedRange);
                    ReleaseComObject(range);
                    ReleaseComObject(name);
                    ReleaseComObject(cell);

                    // ReSharper disable RedundantAssignment
                    sheet = null;
                    // ReSharper restore RedundantAssignment
                }
                catch
                {
                    return;
                }
            }

        }


        /// <summary>
        /// Sets a list of values into a named range. Expects that passed range is a singe dimensional list of cells (i.e. a single column of cells).
        /// NOTE: This version of the method is able to fill a range that is either a column of cells, a row of cells, or columns and rows of cells. 
        /// </summary>
        /// <param name="sheet">The sheet containing the named range.</param>
        /// <param name="namedRangeName">The name of the named range.</param>
        /// <param name="valList">The list of values (as strings).</param>
        public static void SetNamedRangeValues2(_Worksheet sheet, string namedRangeName, List<string> valList)
        {
            if (sheet == null || String.IsNullOrEmpty(namedRangeName) || valList.Count <= 0) return;

            Name name = null;
            Range range = null;
            Range cell = null;
            object objNamedRange;

            // Get the range by name
            try
            {
                objNamedRange = sheet.Names.Item(namedRangeName, Type.Missing, Type.Missing);
            }
            catch
            {
                objNamedRange = null;
                Logger.LogMessage(String.Format("Error executing WorksheetUtilities.SetNamedRangeValues. \r\nRange {0} doesn't exist in the sheet.", namedRangeName), Level.Error);
            }
            if (objNamedRange == null || !(objNamedRange is Name)) goto CleanUp_SetNamedRangeValues;
            name = objNamedRange as Name;
            range = name.RefersToRange;

            int rangeRowCount = range.Rows.Count;
            int rangeColCount = range.Columns.Count;

            // Fill the range top down, left to right
            int rowIdx = 1;
            int colIdx = 1;
            foreach (string val in valList)
            {
                if (rowIdx > rangeRowCount)
                {
                    // Move to the next column, first row
                    colIdx++;
                    rowIdx = 1;
                }

                // Has the range been completely filled?
                if (colIdx > rangeColCount) break;

                cell = range.Cells[rowIdx, colIdx] as Range;
                try
                {
                    if (cell != null) cell.Value2 = val;
                }
                catch (Exception ex)
                {
                    Logger.LogMessage("Error executing WorksheetUtilities.SetNamedRangeValues to set cell value. \r\n" + ex.Message, Level.Error);
                }

                // Clean up
                ReleaseComObject(cell);
                rowIdx++;
            }

            // The com objects must always be released
        CleanUp_SetNamedRangeValues:
            {
                try
                {
                    ReleaseComObject(objNamedRange);
                    ReleaseComObject(range);
                    ReleaseComObject(name);
                    ReleaseComObject(cell);

                    // ReSharper disable RedundantAssignment
                    sheet = null;
                    // ReSharper restore RedundantAssignment
                }
                catch
                {
                    return;
                }
            }

        }


        /// <summary>
        /// Sets a values into a specific location in a named range.
        /// </summary>
        /// <param name="sheet">The sheet containing the named range.</param>
        /// <param name="namedRangeName">The name of the named range.</param>
        /// <param name="value">The value to set.</param>
        /// <param name="rowOffset">The row offset within the range to set the value.</param>
        /// <param name="colOffset">The column offset within the range to set the value.</param>
        public static void SetNamedRangeValue(_Worksheet sheet, string namedRangeName, string value, int rowOffset, int colOffset)
        {
            if (sheet == null || String.IsNullOrEmpty(namedRangeName) || String.IsNullOrEmpty(value)) return;

            Name name = null;
            Range range = null;
            Range cell = null;
            object objNamedRange;

            // Get the range by name
            try
            {
                objNamedRange = sheet.Names.Item(namedRangeName, Type.Missing, Type.Missing);
            }
            catch
            {
                objNamedRange = null;
                Logger.LogMessage(String.Format("Error executing WorksheetUtilities.SetNamedRangeValue. \r\nRange {0} doesn't exist in the sheet.", namedRangeName), Level.Error);
            }
            if (objNamedRange == null || !(objNamedRange is Name)) goto CleanUp_SetNamedRangeValue;
            name = objNamedRange as Name;
            range = name.RefersToRange;

            try
            {
                cell = range.Cells[rowOffset, colOffset] as Range;
                //string cellAddress = cell.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                if (cell != null) cell.Value2 = value;
            }
            catch (Exception ex)
            {
                Logger.LogMessage("Error executing WorksheetUtilities.SetNamedRangeValue to set cell value.\r\n" + ex.Message, Level.Error);
            }

            // The com objects must always be released
            CleanUp_SetNamedRangeValue:
            {
                try
                {
                    ReleaseComObject(objNamedRange);
                    ReleaseComObject(range);
                    ReleaseComObject(name);
                    ReleaseComObject(cell);

                    // ReSharper disable RedundantAssignment
                    sheet = null;
                    // ReSharper restore RedundantAssignment
                }
                catch
                {
                    return;
                }
            }

        }



        /// <summary>
        /// Sets the font style specific location in a named range.
        /// </summary>
        /// <param name="sheet">The sheet containing the named range.</param>
        /// <param name="namedRangeName">The name of the named range.</param>
        /// <param name="rowOffset">The row offset within the range to set the value.</param>
        /// <param name="colOffset">The column offset within the range to set the value.</param>
        public static void SetNamedRangeCellFontBold(_Worksheet sheet, string namedRangeName, int rowOffset, int colOffset)
        {
            if (sheet == null || String.IsNullOrEmpty(namedRangeName)) return;

            Name name = null;
            Range range = null;
            Range cell = null;
            object objNamedRange;

            // Get the range by name
            try
            {
                objNamedRange = sheet.Names.Item(namedRangeName, Type.Missing, Type.Missing);
            }
            catch
            {
                objNamedRange = null;
                Logger.LogMessage(String.Format("Error executing WorksheetUtilities.SetNamedRangeCellFontBold. \r\nRange {0} doesn't exist in the sheet.", namedRangeName), Level.Error);
            }
            if (objNamedRange == null || !(objNamedRange is Name)) goto CleanUp_SetNamedRangeCellFontBold;
            name = objNamedRange as Name;
            range = name.RefersToRange;

            try
            {
                cell = range.Cells[rowOffset, colOffset] as Range;
                if (cell != null) cell.Font.Bold = true;
            }
            catch (Exception ex)
            {
                Logger.LogMessage("Error executing WorksheetUtilities.SetNamedRangeCellFontBold to set cell value.\r\n" + ex.Message, Level.Error);
            }

            // The com objects must always be released
        CleanUp_SetNamedRangeCellFontBold:
            {
                try
                {
                    ReleaseComObject(objNamedRange);
                    ReleaseComObject(range);
                    ReleaseComObject(name);
                    ReleaseComObject(cell);

                    // ReSharper disable RedundantAssignment
                    sheet = null;
                    // ReSharper restore RedundantAssignment
                }
                catch
                {
                    return;
                }
            }

        }



        /// <summary>
        /// Creates a series of named ranges based on a parent named range and offset rows. For example, MyRange1..MyRangeN.
        /// </summary>
        /// <param name="sheet">The sheet in which to create the named ranges.</param>
        /// <param name="namedRangeName">The first named range in the series (i.e. MUST exist).</param>
        /// <param name="rowOffset">The offset in rows to which the next named range is set.</param>
        /// <param name="colOffset">The offset in columns to which the next named range is set.</param>
        /// <param name="numRanges">The number of ranges to create.</param>
        public static void CreateNamedRangeSeriesFromParent(_Worksheet sheet, string namedRangeName, int rowOffset, int colOffset, int numRanges)
        {
            if (sheet == null || String.IsNullOrEmpty(namedRangeName) || numRanges <= 0) return;

            Name namedRange = null;
            Range range = null;
            Range destRange = null;
            object objNamedRange;

            // Get the range by name for the FIRST range in the series
            try
            {
                objNamedRange = sheet.Names.Item(namedRangeName + "1", Type.Missing, Type.Missing);
            }
            catch
            {
                objNamedRange = null;
                Logger.LogMessage(String.Format("Error executing WorksheetUtilities.CreateNamedRangeSeriesFromParent. \r\nRange {0} doesn't exist in the sheet.", namedRangeName), Level.Error);
            }
            if (objNamedRange == null || !(objNamedRange is Name)) goto CleanUp_CreateNamedRangeSeriesFromParent;

            namedRange = objNamedRange as Name;
            range = namedRange.RefersToRange;

            for (int i = 1; i <= numRanges; i++)
            {
                destRange = range.get_Offset(rowOffset * i, colOffset * i);
                string destAddress = destRange.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

                string refersToLocal = "='" + sheet.Name + "'!" + destAddress;

                sheet.Names.Add(namedRangeName + (i + 1), Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing,
                                refersToLocal, Type.Missing, Type.Missing, Type.Missing);

                ReleaseComObject(destRange);
            }

            // The com objects must always be released
            CleanUp_CreateNamedRangeSeriesFromParent:
            {
                try
                {
                    ReleaseComObject(objNamedRange);
                    ReleaseComObject(range);
                    ReleaseComObject(namedRange);
                    ReleaseComObject(destRange);

                    // ReSharper disable RedundantAssignment
                    sheet = null;
                    // ReSharper restore RedundantAssignment
                }
                catch
                {
                    return;
                }
            }
        }


        /// <summary>
        /// Copies a named range to a destination determined as a row and column index (offset) from the first cell in the source range.
        /// </summary>
        /// <param name="sheet">The worksheet.</param>
        /// <param name="srcNamedRange">The source named range name.</param>
        /// <param name="destNamedRange">The name to name the destination range.</param>
        /// <param name="rowIndex">The row index to which to paste the range.</param>
        /// <param name="colIndex">The column index to which to paste the range.</param>
        /// <param name="pasteType">The paste type.</param>
        public static void CopyNamedRangeToNewNamedRange(_Worksheet sheet, string srcNamedRange, string destNamedRange, int rowIndex, int colIndex, XlPasteType pasteType)
        {
            if (sheet == null || String.IsNullOrEmpty(srcNamedRange)) return;

            Name srcName = null;
            Range srcRange = null;
            Range destRange = null;
            Range destRange2 = null;
            Range destRangeForName = null;
            object objSrcNamedRange;

            // Get the range by name
            try
            {
                objSrcNamedRange = sheet.Names.Item(srcNamedRange, Type.Missing, Type.Missing);
            }
            catch
            {
                objSrcNamedRange = null;
                Logger.LogMessage(String.Format("Error executing WorksheetUtilities.CopyNamedRangeToNewNamedRange. \r\nRange {0} doesn't exist in the sheet.", srcNamedRange), Level.Error);
            }
            if (!(objSrcNamedRange is Name)) goto CleanUp_CopyNamedRangeToNewRange;

            srcName = objSrcNamedRange as Name;
            srcRange = srcName.RefersToRange;
            
            destRange = srcRange.Cells[rowIndex, colIndex] as Range;
            if (destRange == null) goto CleanUp_CopyNamedRangeToNewRange;

            srcRange.Copy(Type.Missing);
            //string destAddress = destRange.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
            destRange.PasteSpecial(pasteType, XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

            if (!String.IsNullOrEmpty(destNamedRange))
            {
                // Name the new range (pasted to)
                //destRangeForName = sheet.Application.Selection as Range;
                destRange2 = destRange.Cells[srcRange.Rows.Count, srcRange.Columns.Count] as Range;
                //string destRange2Address = destRange2.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                if (destRange2 != null) destRangeForName = sheet.get_Range(destRange, destRange2);
                if (destRangeForName != null)
                {
                    //string destNamedRangeAddress = destRangeForName.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

                    string refersToLocal = "='" + sheet.Name + "'!" + destRangeForName.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

                    sheet.Names.Add(destNamedRange, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, refersToLocal, Type.Missing, Type.Missing, Type.Missing);
                }
            }

            // The com objects must always be released
            CleanUp_CopyNamedRangeToNewRange:
            {
                try
                {
                    ReleaseComObject(objSrcNamedRange);
                    ReleaseComObject(srcRange);
                    ReleaseComObject(srcName);
                    ReleaseComObject(destRange);
                    ReleaseComObject(destRange2);
                    ReleaseComObject(destRangeForName);

                    // ReSharper disable RedundantAssignment
                    sheet = null;
                    // ReSharper restore RedundantAssignment
                }
                catch
                {
                    return;
                }
            }
        }



        /// <summary>
        /// Copies a named range to a destination determined as a row and column index (offset) from the first cell in the source range.
        /// </summary>
        /// <param name="sheet">The worksheet.</param>
        /// <param name="srcNamedRange">The source named range name.</param>
        /// <param name="destNamedRange">The name to name the destination range.</param>
        /// <param name="rowIndex">The row index in the same sheet to which to paste the range.</param>
        /// <param name="colIndex">The column index in the same sheet to which to paste the range.</param>
        /// <param name="pasteType">The paste type.</param>
        public static void CopyNamedRangeToNewLocationWithNewNamedRange(_Worksheet sheet, string srcNamedRange, string destNamedRange, int rowIndex, int colIndex, XlPasteType pasteType)
        {
            if (sheet == null || String.IsNullOrEmpty(srcNamedRange)) return;

            Name srcName = null;
            Range srcRange = null;
            Range destRange = null;
            Range destRange2 = null;
            Range destRangeForName = null;
            object objSrcNamedRange;

            // Get the range by name
            try
            {
                objSrcNamedRange = sheet.Names.Item(srcNamedRange, Type.Missing, Type.Missing);
            }
            catch
            {
                objSrcNamedRange = null;
                Logger.LogMessage(String.Format("Error executing WorksheetUtilities.CopyNamedRangeToNewNamedRange. \r\nRange {0} doesn't exist in the sheet.", srcNamedRange), Level.Error);
            }
            if (!(objSrcNamedRange is Name)) goto CleanUp_CopyNamedRangeToNewRange;

            srcName = objSrcNamedRange as Name;
            srcRange = srcName.RefersToRange;

            destRange = sheet.Cells[rowIndex, colIndex] as Range;
            if (destRange == null) goto CleanUp_CopyNamedRangeToNewRange;

            srcRange.Copy(Type.Missing);
            //string destAddress = destRange.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
            destRange.PasteSpecial(pasteType, XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

            if (!String.IsNullOrEmpty(destNamedRange))
            {
                // Name the new range (pasted to)
                //destRangeForName = sheet.Application.Selection as Range;
                destRange2 = destRange.Cells[srcRange.Rows.Count, srcRange.Columns.Count] as Range;
                //string destRange2Address = destRange2.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                if (destRange2 != null) destRangeForName = sheet.get_Range(destRange, destRange2);
                if (destRangeForName != null)
                {
                    //string destNamedRangeAddress = destRangeForName.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

                    string refersToLocal = "='" + sheet.Name + "'!" + destRangeForName.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

                    sheet.Names.Add(destNamedRange, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, refersToLocal, Type.Missing, Type.Missing, Type.Missing);
                }
            }

            // The com objects must always be released
        CleanUp_CopyNamedRangeToNewRange:
            {
                try
                {
                    ReleaseComObject(objSrcNamedRange);
                    ReleaseComObject(srcRange);
                    ReleaseComObject(srcName);
                    ReleaseComObject(destRange);
                    ReleaseComObject(destRange2);
                    ReleaseComObject(destRangeForName);

                    // ReSharper disable RedundantAssignment
                    sheet = null;
                    // ReSharper restore RedundantAssignment
                }
                catch
                {
                    return;
                }
            }
        }

        
        /// <summary>
        /// Resizes a chart by specified index. Can resize height, width, or both.
        /// </summary>
        /// <param name="sheet">The worksheet containing the chart.</param>
        /// <param name="index">The index of the chart in the sheet.</param>
        /// <param name="height">The height to set, must be >0 if height is specified for resize.</param>
        /// <param name="width">The width to set, must be >0 if width is specified for resize.</param>
        /// <param name="resizeType">Either Height, Width, or HeightAndWidth.</param>
        public static void ResizeChart(_Worksheet sheet, int index, double height, double width, ResizeType resizeType)
        {
            if (sheet == null) return;

            // Consider these invalid
            if (resizeType == ResizeType.Height && height <= 0) return;
            if (resizeType == ResizeType.Width && width <= 0) return;
            if (resizeType == ResizeType.HeightAndWidth && (height <= 0 || width <= 0)) return;

            ChartObject chartObj;
            try
            {
                // This is actually a ChartObject instance and NOT the interface IChartObject
                chartObj = (ChartObject) sheet.ChartObjects(index);
            }
            catch
            {
                chartObj = null;
            }

            if (chartObj == null)
            {
                // ReSharper disable RedundantAssignment
                sheet = null;
                // ReSharper restore RedundantAssignment
                return;
            }

            // Set height
            if ((resizeType == ResizeType.Height || resizeType == ResizeType.HeightAndWidth))
            {
                chartObj.Height = height;
            }

            // Set width
            if ((resizeType == ResizeType.Width || resizeType == ResizeType.HeightAndWidth))
            {
                chartObj.Width = width;
            }

            // The com objects must always be released
            try
            {
                ReleaseComObject(chartObj);
                // ReSharper disable RedundantAssignment
                sheet = null;
                // ReSharper restore RedundantAssignment
            }
            catch
            {
                return;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="namedRangeName"></param>
        /// <param name="rowChange"></param>
        /// <param name="colChange"></param>
        public static void ResizeNamedRange(_Worksheet sheet, string namedRangeName, int rowChange, int colChange)
        {
            if (sheet == null ) return;
            Name namedRange = null;
            Range range = null;
            Range startCell = null;
            Range endCell = null;
            Name objNamedRange = null;
            // Get the range by name
            try
            {
                namedRange = sheet.Names.Item(namedRangeName, Type.Missing, Type.Missing);
                range = namedRange.RefersToRange;
                if (range != null)
                {
                    // Update the named range to point to the new address (since single rowed named ranges don't expand dynamically)
                    startCell = namedRange.RefersToRange.Cells[1, 1] as Range;
                    endCell = namedRange.RefersToRange.Cells[range.Rows.Count + rowChange, range.Columns.Count + colChange] as Range;

                    if (startCell != null && endCell != null)
                    {
                        string refersToLocal = "='" + sheet.Name + "'!" +
                                         startCell.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing) + ":" +
                                         endCell.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

                        // Update the range
                        sheet.Names.Add(namedRangeName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, refersToLocal, Type.Missing, Type.Missing, Type.Missing);
                        //namedRange.RefersToLocal = refersToLocal;
                    }

                }
            }
            catch(Exception ex)
            {
                objNamedRange = null;
                Logger.LogMessage(String.Format("Error executing WorksheetUtilities.ResizeNamedRange. \r\n " + ex.Message), Level.Error);
            }
            finally
            {
                try
                {
                    ReleaseComObject(objNamedRange);
                    ReleaseComObject(namedRange);
                    ReleaseComObject(range);
                    ReleaseComObject(endCell);
                    ReleaseComObject(startCell);

                    // ReSharper disable RedundantAssignment
                    sheet = null;
                    // ReSharper restore RedundantAssignment
                }
                catch
                {
                    
                }

            }
        }

        public static bool NamedRangeExist(_Worksheet sheet, string namedRangeName)
        {
            bool ret = false;
            Name namedRange = null;
            Range range = null;
            if (sheet == null) return ret;
            // Get the range by name
            try
            {
                namedRange = sheet.Names.Item(namedRangeName, Type.Missing, Type.Missing);
                range = namedRange.RefersToRange;
                if (range != null)
                {
                    ret = true;
                }
            }
            catch (Exception ex)
            {
                Logger.LogMessage(String.Format("Error executing WorksheetUtilities.ResizeNamedRange. \r\n " + ex.Message), Level.Error);
                ret = false;
            }
            finally
            {
                try
                {
                    ReleaseComObject(namedRange);
                    ReleaseComObject(range);
                    sheet = null;
                }
                catch { }

            }

            return ret;
        }

        public static void SetSequentialNumbersInNamedRange(Worksheet sheet, string namedRange, int count)
        {
            List<string> values = new List<string>(count);

            for (int i = 1; i <= count; i++)
                values.Add(i.ToString());

            SetNamedRangeValues(sheet, namedRange, values);
        }

        public static void ResizeNamedRangeRows(Worksheet sheet, string namedRange, int desiredRowCount, XlPasteType pasteType = XlPasteType.xlPasteAll)
        {
            int currentRowCount = GetNamedRangeRowCount(sheet, namedRange);

            if (desiredRowCount > currentRowCount)
            {
                // Insert the difference
                InsertRowsIntoNamedRange(desiredRowCount - currentRowCount, sheet, namedRange, true, XlDirection.xlDown, pasteType);
            }
            else if (desiredRowCount < currentRowCount)
            {
                // Delete rows until the named range shrinks to match
                DeleteRowsFromNamedRange(currentRowCount - desiredRowCount, sheet, namedRange, XlDirection.xlDown);
            }
        }

        public static double GetColumnWidth(Worksheet sheet, string namedRange, int colIndex)
        {
            Range range = sheet.Range[namedRange];
            Range cell = range.Columns[colIndex];
            double width = cell.ColumnWidth;

            ReleaseComObject(cell);
            ReleaseComObject(range);

            return width;
        }

        public static void SetColumnWidth(Worksheet sheet, string namedRange, int colIndex, double width)
        {
            Range range = sheet.Range[namedRange];
            Range cell = range.Columns[colIndex];

            cell.EntireColumn.ColumnWidth = width;

            ReleaseComObject(cell);
            ReleaseComObject(range);
        }

        public static void CopyColumnWidthBetweenNamedRanges(Worksheet sheet, string sourceNamedRange, string destinationNamedRange)
        {
            double width = GetColumnWidth(sheet, sourceNamedRange, 1);
            SetColumnWidth(sheet, destinationNamedRange, 1, width);
        }

        public static void SetMetadataValues(Worksheet sheet, string validationType, string productType, string testType = null)
        {
            SetNamedRangeValue(sheet, "MetaData", validationType, 1, 1);
            SetNamedRangeValue(sheet, "MetaData", productType, 2, 1);

            if (testType != null)
            {
                SetNamedRangeValue(sheet, "MetaData", testType, 3, 1);
            }
        }

        public static string GetSimpleReferenceFormula(string sourceAddress)
        {
            return $"=IF({sourceAddress}=\"\", \"\", {sourceAddress})";
        }

        public static string GetDifferenceConditionsLinkingFormula(string sourceAddress, string decimalCellAddress)
        {
            return $"=IF({sourceAddress}=\"\", \"\", IF(ISNUMBER({sourceAddress}), FIXED({sourceAddress}, {decimalCellAddress}), {sourceAddress}))";
        }

        public static void LinkDifferenceCondition(Worksheet sheet, string sourceNamedRange, string targetNamedRange, string decimalsCellAbsoluteAddress, int row, int col)
        {
            string sourceAddress = GetCellAddress(sheet, sourceNamedRange, row, col);
            string formula = GetDifferenceConditionsLinkingFormula(sourceAddress, decimalsCellAbsoluteAddress);
            SetNamedRangeFormula(sheet, targetNamedRange, formula, row, col);
        }

        public static string GetCellAddress(Worksheet sheet, string namedRange, int row, int col, bool rowAbsolute = false, bool colAbsolute = false)
        {
            Range sourceRange = sheet.Range[namedRange];
            Range sourceCell = sourceRange.Cells[row, col] as Range;
            return sourceCell.Address[rowAbsolute, colAbsolute, XlReferenceStyle.xlA1];
        }

        public static void SetNamedRangeFormula(Worksheet sheet, string namedRange, string formula, int rowIndex, int colIndex)
        {
            if (sheet == null || string.IsNullOrEmpty(namedRange) || string.IsNullOrEmpty(formula))
                return;

            try
            {
                Range range = sheet.Range[namedRange];
                Range targetCell = range.Cells[rowIndex, colIndex] as Range;
                if (targetCell != null)
                {
                    targetCell.Formula = formula;
                }
            }
            catch (Exception ex)
            {
                Logger.LogMessage($"Failed to set formula in named range '{namedRange}': {ex.Message}", Level.Error);
            }
        }

        // make the first of destinationNamedRange refer to first cell of sourceNamedRange.
        public static void LinkFirstCell(Worksheet sheet, string sourceNamedRange, string destinationNamedRange)
        {
            // Retrieve the named ranges
            Range sourceRange = sheet.Range[sourceNamedRange];
            Range destinationRange = sheet.Range[destinationNamedRange];

            // Ensure both ranges have at least one cell
            if (sourceRange.Cells.Count > 0 && destinationRange.Cells.Count > 0)
            {
                // Get the first cell of each range
                Range sourceCell = sourceRange.Cells[1, 1] as Range;
                Range destinationCell = destinationRange.Cells[1, 1] as Range;

                // Get the address of the source cell in A1 format
                string sourceAddress = sourceCell.Address[false, false, XlReferenceStyle.xlA1];

                // Set the formula in the destination cell to reference the source cell
                destinationCell.Formula = $"=IF({sourceAddress}=\"\", \"\", {sourceAddress})";

                // Release COM objects
                ReleaseComObject(sourceCell);
                ReleaseComObject(destinationCell);
            }

            // Release COM objects
            ReleaseComObject(sourceRange);
            ReleaseComObject(destinationRange);
        }

        // make each cell of destinationNamedRange refer to corresponding cell of sourceNamedRange.
        public static void LinkVerticalNamedRanges(Worksheet sheet, string sourceNamedRange, string destinationNamedRange)
        {
            int rowCount = GetNamedRangeRowCount(sheet, sourceNamedRange);

            for (int i = 1; i <= rowCount; i++)
            {
                string sourceAddress = GetCellAddress(sheet, sourceNamedRange, i, 1);

                string formula = GetSimpleReferenceFormula(sourceAddress);

                SetNamedRangeFormula(sheet, destinationNamedRange, formula, i, 1);
            }
        }

        public static void LinkTwoNamedRangeCellsWrapper(
            _Worksheet sheet,
            string sourceRange,
            string targetRange,
            int sourceCol,
            int targetCol,
            bool addRound = false,
            bool roundForImpurity = false,
            int sourceRow = -1,
            int targetRow = -1)
        {
            LinkTwoNamedRangeCells(sheet, sourceRange, targetRange, sourceRow, sourceCol, targetRow, targetCol, addRound, roundForImpurity);
        }


        /// <summary>
        /// For Solution Stability Result Section - Links Area/Assay/LabelClaim to Summary Tables
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="namedRangeSource"></param>
        /// <param name="namedRangeTarget"></param>
        /// <param name="sourceRow"></param>
        /// <param name="sourceCol"></param>
        /// <param name="targetRow"></param>
        /// <param name="targetCol"></param>
        /// <param name="addRound">Adds Round formula - ONLY FOR ASSAY SUMMARY TABLE</param>
        /// <param name="roundForImpurity">Adds Round Formula - ONLY FOR IMPURITY SUMMARY TABLE</param>
        public static void LinkTwoNamedRangeCells(
            _Worksheet sheet,                  // Excel worksheet object
            string namedRangeSource,          // Source named range
            string namedRangeTarget,          // Target named range
            int sourceRow,                    // Row index in source range (-1 means all rows)
            int sourceCol,                    // Column index in source range
            int targetRow,                    // Row index in target range (-1 means all rows)
            int targetCol,                    // Column index in target range
            bool addRound,                    // Whether to round values
            bool roundForImpurity             // Whether rounding is for impurity values
        )
        {
            if (String.IsNullOrEmpty(namedRangeTarget) || String.IsNullOrEmpty(namedRangeSource) || sheet == null)
                return;
            Range sourceR = null;
            Range targetR = null;
            Range roundingRange = null;
            try
            {
                 sourceR = GetNamedRange(sheet, namedRangeSource);
                 targetR = GetNamedRange(sheet, namedRangeTarget);

                if (sourceR == null || targetR == null) return;

                int totalRowInSource = sourceR.Rows.Count;
                int totalColInSource = sourceR.Columns.Count;

                int totalRowInTarget = targetR.Rows.Count;
                int totalColInTarget = targetR.Columns.Count;

                if (sourceRow > 0 && sourceCol > 0)
                {
                    // link one cell
                    string fa = "=" + ((Range)sourceR.Cells[sourceRow, sourceCol]).Address.Replace("$", "");
                    if (targetRow > 0 && targetCol > 0)
                    {
                        // one cell to one cell
                        ((Range)targetR.Cells[targetRow, targetCol]).Formula = fa;
                    }
                    else if (targetRow > 0)
                    {
                        // copy one cell from source to how row of target
                        for (int c = 0; c < totalColInTarget; c++)
                        {
                            ((Range)targetR.Cells[targetRow, c]).Formula = fa;
                        }
                    }
                    else if (targetCol > 0)
                    {
                        // copy one cell from source to how row of target
                        for (int c = 0; c < totalRowInTarget; c++)
                        {
                            ((Range)targetR.Cells[c, targetCol]).Formula = fa;
                        }
                    }

                }
                else if (sourceRow > 0)  // Link By Target Row
                {
                    string fa = "";
                    string address = "";

                    for (int c = 0; c < totalColInSource; c++)
                    {
                        address = ((Range)sourceR.Cells[sourceRow, c]).Address.Replace("$", "");
                        fa = "=IF(" + address + "<>\"\"," + address + ", \"\")";
                        if (targetRow > 0)
                        {
                            ((Range)targetR.Cells[targetRow, c]).Formula = fa;
                        }
                        else if (targetCol > 0)
                        {
                            ((Range)targetR.Cells[c, targetCol]).Formula = fa;
                        }
                    }

                }
                else if (sourceCol > 0)  // Link By Target Col
                {
                    string fa = "";
                    string address = "";
                    for (int c = 1; c <= totalRowInSource; c++)
                    {
                        // Get the address of the cell at row 'c' and column 'sourceCol' in the source range,
                        // and remove the dollar signs ($) to make it a relative reference.
                        address = ((Range)sourceR.Cells[c, sourceCol]).Address.Replace("$", "");

                        // Build an Excel formula string that checks if the cell is not empty.
                        // If it's not empty, return the cell's value; otherwise, return an empty string.
                        fa = "=IF(" + address + "<>\"\", " + address + ", \"\")";

                        if (targetRow > 0)
                        {
                            ((Range)targetR.Cells[targetRow, c]).Formula = fa;
                        }
                        else if (targetCol > 0)
                        {
                            //Results Table now has a rounding part of the formula
                            if (addRound)
                            {
                                string roundAddress = ((Range)targetR.Cells[c, targetCol + 2]).Address.Replace("$", "");
                                fa = "=IF(" +"ISNUMBER("+ address +")"+ "," +"FIXED( " + address + ", "+ roundAddress + " ,TRUE)" + ", " + "IF(" + address + "<>\"\"," + address + ", \"\")" + ")";
                            }

                            //Adds rounding for Impurity Summary Table - Needs to target the same column
                            if (roundForImpurity)
                            {
                                //Here - For Example if namedRange = Impurity101 it gets cut as 01 and not 10 as it should be
                                //string rangeNumber = namedRangeSource.Substring(namedRangeSource.Length - 2);
                                //rangeNumber with this Regex leaves only the Number as a String.
                                string rangeNumber = Regex.Match(namedRangeSource, @"\d+").Value;

                                if (rangeNumber.Length > 1)
                                {
                                    rangeNumber = rangeNumber.Substring(0, rangeNumber.Length - 1);
                                }

                                roundingRange = GetNamedRange(sheet, "ImpuritySummaryImpurityColumn" + rangeNumber);

                                if (roundingRange != null)
                                {
                                    string roundAddress = ((Range)roundingRange.Cells[1, 1]).Address.Replace("$", "");
                                    fa = "=IF(" + "ISNUMBER(" + address + ")" + "," + "FIXED( " + address + ", " + roundAddress + ")" + ", " + "IF(" + address + "<>\"\"," + address + ", \"\")" + ")";
                                }
                                else
                                {
                                    rangeNumber = namedRangeSource.Substring(namedRangeSource.Length - 1);

                                    roundingRange = GetNamedRange(sheet, "ImpuritySummaryImpurityColumn" + rangeNumber);
                                    if (roundingRange != null)
                                    {
                                        string roundAddress = ((Range)roundingRange.Cells[1, 1]).Address.Replace("$", "");
                                        fa = "=IF(" + "ISNUMBER(" + address + ")" + "," + "FIXED( " + address + ", " + roundAddress + ")" + ", " + "IF(" + address + "<>\"\"," + address + ", \"\")" + ")";
                                    }
                                }
                            }

                            ((Range)targetR.Cells[c, targetCol]).Formula = fa;
                        }
                    }
                }
            }
            finally
            {
                try
                {
                    ReleaseComObject(sourceR);
                    ReleaseComObject(targetR);
                    ReleaseComObject(roundingRange);
                }
                catch (Exception ex) { }
            }
        }

        /// <summary>
        /// Get the Formulas for a namedRange & replicates them on all the named range. (Used in Linearity new formatting)
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="namedRangeName"></param>
        /// <param name="rowIndex">From where the formula needs to be pulled - Ideal with the first row that has the formula</param>
        /// <returns></returns>
        public static void RefreshFormulasforNamedRange(_Worksheet sheet, string namedRangeName, int rowIndex)
        {
            if (sheet == null) return;

            Name namedRange = null;
            Range range = null;
            Range cell = null;
            object objNamedRange;

            // Get the range by name
            try
            {
                objNamedRange = sheet.Names.Item(namedRangeName, Type.Missing, Type.Missing);
            }
            catch
            {
                objNamedRange = null;
                Logger.LogMessage(String.Format("Error executing WorksheetUtilities.UpdateSystemSuitabilityFormulas. \r\nRange {0} doesn't exist in the sheet.", namedRangeName), Level.Error);
            }
            if (objNamedRange == null || !(objNamedRange is Name)) goto CleanUp_InsertRowsForNamedRange;

            namedRange = objNamedRange as Name;
            range = namedRange.RefersToRange;



            if (range != null)
            {

                String formula = "";

                var r = range.Count - rowIndex;

                cell = range.Cells[rowIndex, 1];

                for (var i = 0; i < r; i++)
                {
                    if(formula == "")
                    {
                        formula = cell.Formula;
                    }

                    if(cell.Formula == "" && formula != "")
                    {
                        cell.Formula = formula;
                    }

                    if (i == 0)
                    {
                        cell = range.Cells[2, 1];
                    }
                    else
                    {
                        cell = range.Cells[2 + i, 1];
                    }

                }
            }

        // The com objects must always be released
        CleanUp_InsertRowsForNamedRange:
            {
                ReleaseComObject(objNamedRange);
                ReleaseComObject(range);
                ReleaseComObject(namedRange);
                ReleaseComObject(cell);

                // ReSharper disable RedundantAssignment
                sheet = null;
                // ReSharper restore RedundantAssignment
            }
        }

        /// <summary>
        /// Update Formulas by number of rows & tables - For Validation Report
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="namedRangeName"></param>
        public static void UpdateSensitivityFormulas(_Worksheet sheet, int numImpurities,int offset)
        {
            if (sheet == null) return;

            Name namedRange = null;
            Range range = null;
            object objNamedRange = null;
            Range cell = null;

            object valobjNamedRange = null;
            Name valnamedRange = null;
            Range valrange = null;
            Range valcell = null;

            for(var i = 1; i <= numImpurities; i++)
            {
                var namedRangeNum = i + 1;
                // Get the range by name
                try
                {
                    objNamedRange = sheet.Names.Item("ImpurityResults" + namedRangeNum, Type.Missing, Type.Missing);
                }
                catch
                {
                    objNamedRange = null;
                    Logger.LogMessage(String.Format("Error executing WorksheetUtilities.UpdateSystemSuitabilityFormulas. \r\nRange {0} doesn't exist in the sheet." + "ImpurityResults" + namedRangeNum), Level.Error);
                }
                if (objNamedRange == null || !(objNamedRange is Name)) goto CleanUp_InsertRowsForNamedRange;

                namedRange = objNamedRange as Name;
                range = namedRange.RefersToRange;

                if (range != null)
                {

                    String formula = "";
                    String address = "";
                    String partialFormula = "";

                    //PeakName Formula
                    cell = range.Cells[2, 2];

                    address = cell.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

                    valobjNamedRange = sheet.Names.Item("VRPeakName", Type.Missing, Type.Missing);
                    valnamedRange = valobjNamedRange as Name;
                    valrange = valnamedRange.RefersToRange;

                    valcell = valrange.Cells[1, 1];

                    formula = valcell.Formula;

                    partialFormula = formula.Split(')')[0];
                    partialFormula = partialFormula + "," + "\",\"" + "," + address + ")";


                    valcell.Formula = partialFormula;

                    //STD Dev Formula
                    cell = range.Cells[12 + offset, 2];

                    address = cell.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

                    valobjNamedRange = sheet.Names.Item("VRSTDDEV", Type.Missing, Type.Missing);
                    valnamedRange = valobjNamedRange as Name;
                    valrange = valnamedRange.RefersToRange;

                    valcell = valrange.Cells[1, 1];

                    formula = valcell.Formula;

                    //partialFormula = formula.Split(')')[0];
                    partialFormula = formula.Substring(0,formula.Length - 1);
                    partialFormula = partialFormula + "," + "\",\"" + "," + "TEXT("+ address + ",\"0.00\")" + ")";


                    valcell.Formula = partialFormula;

                    //RSD Formula
                    cell = range.Cells[14 + offset, 2];

                    address = cell.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

                    valobjNamedRange = sheet.Names.Item("VRRSD", Type.Missing, Type.Missing);
                    valnamedRange = valobjNamedRange as Name;
                    valrange = valnamedRange.RefersToRange;

                    valcell = valrange.Cells[1, 1];

                    formula = valcell.Formula;

                    partialFormula = formula.Substring(0, formula.Length - 1);
                    //partialFormula = partialFormula + "," + "\",\"" + "," + "TEXT(" + address + ",\"0.00\")" + ")";
                    partialFormula = partialFormula + "," + "\",\"" + "," + address + ")";


                    valcell.Formula = partialFormula;

                    //Mean Formula
                    cell = range.Cells[11 + offset, 2];

                    address = cell.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

                    valobjNamedRange = sheet.Names.Item("VRMean", Type.Missing, Type.Missing);
                    valnamedRange = valobjNamedRange as Name;
                    valrange = valnamedRange.RefersToRange;

                    valcell = valrange.Cells[1, 1];

                    formula = valcell.Formula;

                    partialFormula = formula.Substring(0, formula.Length - 1);
                    partialFormula = partialFormula + "," + "\",\"" + "," + "TEXT(" + address + ",0)" + ")";


                    valcell.Formula = partialFormula;
                }

                objNamedRange = sheet.Names.Item("SignalToNoiseResults" + namedRangeNum, Type.Missing, Type.Missing);

                if (objNamedRange == null || !(objNamedRange is Name)) goto CleanUp_InsertRowsForNamedRange;

                namedRange = objNamedRange as Name;
                range = namedRange.RefersToRange;

                if (range != null)
                {

                    String formula = "";
                    String address = "";
                    String partialFormula = "";

                    //Mean Formula
                    /*cell = range.Cells[9, 2];

                    address = cell.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

                    valobjNamedRange = sheet.Names.Item("VRMean", Type.Missing, Type.Missing);
                    valnamedRange = valobjNamedRange as Name;
                    valrange = valnamedRange.RefersToRange;

                    valcell = valrange.Cells[1, 1];

                    formula = valcell.Formula;

                    partialFormula = formula.Substring(0, formula.Length - 1);
                    partialFormula = partialFormula + "," + "\",\"" + "," + "TEXT(" + address + ",0)" + ")";


                    valcell.Formula = partialFormula;*/

                    //Min - Max Formula
                    cell = range.Cells[10 + offset, 2];

                    address = cell.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

                    cell = range.Cells[11 + offset, 2];

                    String addressMax = cell.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

                    valobjNamedRange = sheet.Names.Item("VRMINMAX", Type.Missing, Type.Missing);
                    valnamedRange = valobjNamedRange as Name;
                    valrange = valnamedRange.RefersToRange;

                    valcell = valrange.Cells[1, 1];

                    formula = valcell.Formula;

                    partialFormula = formula.Substring(0, formula.Length - 1);
                    partialFormula = partialFormula + "," + "\",\"" + "," + "TEXT("+ address +",0)"+"," + "\"-\"" + "," + "TEXT("+ addressMax + ",0)" + ")";

                    valcell.Formula = partialFormula;

                    // Hide row
                    //valrange.Hidden = true;

                }

            }            

        // The com objects must always be released
        CleanUp_InsertRowsForNamedRange:
            {
                ReleaseComObject(objNamedRange);
                ReleaseComObject(range);
                ReleaseComObject(namedRange);
                ReleaseComObject(cell);
                ReleaseComObject(valobjNamedRange);
                ReleaseComObject(valcell);
                ReleaseComObject(valnamedRange);
                ReleaseComObject(valrange);

                // ReSharper disable RedundantAssignment
                sheet = null;
                // ReSharper restore RedundantAssignment
            }
        }

        public static void ReleaseComObject(object obj)
        {
            if (obj != null && Marshal.IsComObject(obj))
            {
                while (Marshal.ReleaseComObject(obj) > 0) { }
            }
        }
    }
}
