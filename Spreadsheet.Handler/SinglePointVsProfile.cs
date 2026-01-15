using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using log4net.Core;
using Microsoft.Office.Interop.Excel;

namespace Spreadsheet.Handler
{
    public class SinglePointVsProfile
    {
        private static Application _app;

        private const string TempDirectoryName = "ABD_TempFiles";

        private const int defaultCompSet = 1;
        private const int defaultTimepoints = 5;

        public static string UpdateSinglePointVsProfileSheet(string sourcePath, int compSet, Dictionary<string, string> acceptanceCriteria)
        {
            string returnPath = "";
            try
            {
                returnPath = UpdateSinglePointProfileSheet(sourcePath, compSet, acceptanceCriteria);
            }
            catch (Exception ex)
            {
                Logger.LogMessage("An error occurred in the call to SinglePointVsProfile.UpdateSinglePointVsProfileSheet. Message and stack trace are:\r\n" + ex.Message + "\r\n" + ex.StackTrace, Level.Error);

                try
                {
                    if (_app.Workbooks.Count > 0)
                    {
                        try
                        {
                            _app.Workbooks[0].Save();
                            returnPath = _app.Workbooks[0].FullName;
                        }
                        catch
                        {
                            Logger.LogMessage("An error occurred in the call to SinglePointVsProfile.UpdateSinglePointVsProfileSheet. Failed to save current workbook changes and to get path.", Level.Error);
                        }

                        _app.Workbooks.Close();
                    }
                    _app = null;
                }
                catch
                {
                    Logger.LogMessage("An error occurred in the call to SinglePointVsProfile.UpdateSinglePointVsProfileSheet. Application failed to close workbooks. Message and stack trace are:\r\n" + ex.Message + "\r\n" + ex.StackTrace, Level.Error);
                }
                finally
                {
                    WorksheetUtilities.ReleaseExcelApp();
                }
            }
            return returnPath;
        }

        private static string UpdateSinglePointProfileSheet(string sourcePath, int compSet, Dictionary<string, string> acceptanceCriteria = null)
        {
            if (!File.Exists(sourcePath))
            {
                Logger.LogMessage("Error in call to SinglePointVsProfile.UpdateSinglePointProfileSheet. Invalid source file path specified.", Level.Error);
                return "";
            }

            // Generate an random temp path to save new workbook
            string savePath = WorksheetUtilities.CopyWorkbook(sourcePath, TempDirectoryName, "Single Point vs. Profile.xlsx");
            if (String.IsNullOrEmpty(savePath)) return "";

            // Try to open the file
            _app = WorksheetUtilities.GetExcelApp();
            _app.Workbooks.Open(savePath, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            Workbook book = _app.Workbooks[1];
            Worksheet sheet = book.Worksheets[1] as Worksheet;

            if (sheet != null)
            {
                bool wasProtected = WorksheetUtilities.SetSheetProtection(sheet, null, false);

                // Update the Acceptance Criteria section if values are provided
                if (acceptanceCriteria != null && acceptanceCriteria.Count > 0)
                {
                    // Set the Validation Type in cell J2 if provided
                    if (acceptanceCriteria.ContainsKey("cmbProtocolType"))
                    {
                        sheet.Cells[4, 6] = acceptanceCriteria["cmbProtocolType"];
                    }
                    if (acceptanceCriteria.ContainsKey("cmbProductType"))
                    {
                        sheet.Cells[4, 9] = acceptanceCriteria["cmbProductType"];
                    }

                    if (acceptanceCriteria.ContainsKey("cmbDiff") &&
                    acceptanceCriteria.ContainsKey("txtdiff"))
                    {
                        sheet.Cells[6, 3] = acceptanceCriteria["cmbDiff"];
                        sheet.Cells[6, 4] = acceptanceCriteria["txtdiff"];
                    }

                    if (acceptanceCriteria.ContainsKey("txtTP"))
                    {
                        sheet.Cells[8, 3] = acceptanceCriteria["txtTP"];
                    }

                    if (acceptanceCriteria.ContainsKey("cmbQTP") &&
                        acceptanceCriteria.ContainsKey("txtQTP"))
                    {
                        sheet.Cells[8, 7] = acceptanceCriteria["cmbQTP"];
                        sheet.Cells[8, 8] = acceptanceCriteria["txtQTP"];
                    }
                }

                if (compSet > defaultCompSet)
                {
                    int rowsToInsert = compSet - defaultCompSet;
                    AppendRowsCopyOnlyFormulasAndRenumber(sheet, "Sample_Table", rowsToInsert, 1, 1, 2);

                    for (int i = 2; i <= compSet; i++)
                    {
                        WorksheetUtilities.InsertRowsIntoNamedRange(20, sheet, "SampleInfoData", false, XlDirection.xlUp, XlPasteType.xlPasteAll);
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "Comparison_Set_" + (i - 1), "Comparison_Set_" + i, 20, 1, XlPasteType.xlPasteAll);
                        WorksheetUtilities.SetNamedRangeValue(sheet, "Comparison_Set_" + i, ("Set-" + i).Trim(), 1, 1);
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "Strength_" + (i - 1), "Strength_" + i, 20, 1, XlPasteType.xlPasteAll);

                    }
                    LinkStrengthNamedRanges(sheet, compSet);
                }

                try
                {
                    _app.Goto(sheet.Cells[1, 1], true);
                }
                catch
                {
                    Logger.LogMessage("Scroll of sheet failed in SinglePointVsProfile.UpdateSinglePointProfileSheet", Level.Error);
                }

                if (wasProtected) WorksheetUtilities.SetSheetProtection(sheet, null, true);

                WorksheetUtilities.ReleaseComObject(sheet);
            }

            _app.Workbooks[1].Save();

            WorksheetUtilities.ReleaseComObject(book);
            _app.Workbooks.Close();

            //while (WorksheetUtilities.ReleaseComObject(_app) >= 0) { }
            _app = null;
            WorksheetUtilities.ReleaseExcelApp();

            // Return the path
            return savePath;

        }

        private static void LinkStrengthNamedRanges(Worksheet sheet, int compSet)
        {

            for (int i = 2; i <= compSet; i++)
            {
                string namedRangeName = $"Strength_{i}";
                Range batchNumRange = null;
                try
                {
                    batchNumRange = sheet.Range[namedRangeName];
                    if (batchNumRange != null)
                    {
                        int sampleRow = 12 + i;
                        batchNumRange.Formula = "=D" + sampleRow;
                    }
                }
                finally
                {
                    WorksheetUtilities.ReleaseComObject(batchNumRange);
                }
            }
        }

        private static void AppendRowsCopyOnlyFormulasAndRenumber(
            Worksheet sheet,
            string namedRange,
            int rowsToAdd,
            int headerRows = 1,
            int templateRowIndex = 1,
            int sampleColIndex = 1)
        {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (string.IsNullOrWhiteSpace(namedRange)) throw new ArgumentNullException(nameof(namedRange));
            if (rowsToAdd <= 0) return;

            Name nm = null;
            Range rng = null;

            try
            {
                nm = sheet.Names.Item(namedRange, Type.Missing, Type.Missing);
                rng = nm.RefersToRange;
                if (rng == null) return;

                int totalCols = rng.Columns.Count;
                int totalRows = rng.Rows.Count;
                if (totalRows <= headerRows)
                    throw new InvalidOperationException("Named range has no data rows below the header.");

                int firstDataRowOffset = headerRows + templateRowIndex;
                if (firstDataRowOffset > totalRows)
                    throw new ArgumentOutOfRangeException(nameof(templateRowIndex), "Template row is outside the range.");

                Range templateRow = rng.Rows[firstDataRowOffset, Type.Missing] as Range;

                string[] templateFormulas = new string[totalCols];
                for (int c = 1; c <= totalCols; c++)
                {
                    Range cell = templateRow.Cells[1, c] as Range;
                    templateFormulas[c - 1] = (cell != null && (bool)cell.HasFormula) ? (string)cell.FormulaR1C1 : null;
                }


                int startInsertRow = rng.Row + totalRows;
                int leftCol = rng.Column;

                for (int i = 0; i < rowsToAdd; i++)
                {
                    Range targetEntireRow = sheet.Rows[startInsertRow + i, Type.Missing] as Range;
                    targetEntireRow.Insert(XlInsertShiftDirection.xlShiftDown, Type.Missing);
                }

                Range startCell = sheet.Cells[rng.Row, leftCol] as Range;
                Range endCell = sheet.Cells[rng.Row + totalRows + rowsToAdd - 1, leftCol + totalCols - 1] as Range;
                string refersToLocal = "='" + sheet.Name + "'!" +
                                       startCell.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing) + ":" +
                                       endCell.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

                sheet.Names.Add(namedRange, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, refersToLocal, Type.Missing, Type.Missing, Type.Missing);


                nm = sheet.Names.Item(namedRange, Type.Missing, Type.Missing);
                rng = nm.RefersToRange;
                totalRows = rng.Rows.Count;

                int firstNewDataRowWithinRange = headerRows + (totalRows - headerRows - rowsToAdd) + 1;
                for (int r = 0; r < rowsToAdd; r++)
                {
                    int rowWithinRange = firstNewDataRowWithinRange + r;

                    templateRow.Copy(Type.Missing);
                    Range destRow = rng.Rows[rowWithinRange, Type.Missing] as Range;
                    destRow.PasteSpecial(XlPasteType.xlPasteFormats,
                                         XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);


                    for (int c = 1; c <= totalCols; c++)
                    {
                        if (!string.IsNullOrEmpty(templateFormulas[c - 1]))
                        {
                            Range destCell = destRow.Cells[1, c] as Range;
                            destCell.FormulaR1C1 = templateFormulas[c - 1];
                        }
                    }
                }


                int dataRowCount = totalRows - headerRows;
                for (int i = 1; i <= dataRowCount; i++)
                {
                    Range sampleCell = rng.Cells[headerRows + i, sampleColIndex] as Range;
                    sampleCell.Value2 = i;
                }
            }
            finally
            {

            }
        }
    }
}
