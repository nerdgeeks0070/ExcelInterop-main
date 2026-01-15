using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using log4net.Core;
using Microsoft.Office.Interop.Excel;

namespace Spreadsheet.Handler
{
    public static class DissoManuvsAuto
    {
        private static Application _app;

        private const string TempDirectoryName = "ABD_TempFiles";

        private const int defaultCompSet = 5;
        private const int defaultTimepoints = 5;


        /// <summary>
        /// Method to be called via Scripts - GenerateExperiment
        /// </summary>
        /// <param name="sourcePath"></param>
        /// <returns></returns>
        public static string UpdateDissoManuvsAutoSheet(string sourcePath, int compSet, int timepoints)
        {
            return UpdateDissoManuvsAutoSheet(sourcePath, compSet, timepoints, null);
        }

        /// <summary>
        /// Method to be called via Scripts - GenerateExperiment with acceptance criteria
        /// </summary>
        /// <param name="sourcePath">Path to the source Excel file</param>
        /// <param name="compSet">Number of comparison sets</param>
        /// <param name="timepoints">Number of timepoints</param>
        /// <param name="acceptanceCriteria">Dictionary containing acceptance criteria values</param>
        /// <returns>Path to the updated Excel file</returns>
        public static string UpdateDissoManuvsAutoSheet(string sourcePath, int compSet, int timepoints,
            Dictionary<string, string> acceptanceCriteria)
        {
            string returnPath = "";
            try
            {
                returnPath = UpdateDissoManuvsAutoSheet4(sourcePath, compSet + 4, timepoints, acceptanceCriteria);
            }
            catch (Exception ex)
            {
                Logger.LogMessage("An error occurred in the call to DissoManuvsAuto.DissoManuvsAuto. Message and stack trace are:\r\n" + ex.Message + "\r\n" + ex.StackTrace, Level.Error);

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
                            Logger.LogMessage("An error occurred in the call to DissoManuvsAuto.UpdateDissoManuvsAutoSheet. Failed to save current workbook changes and to get path.", Level.Error);
                        }

                        _app.Workbooks.Close();
                    }
                    _app = null;
                }
                catch
                {
                    Logger.LogMessage("An error occurred in the call to DissoManuvsAuto.UpdateDissoManuvsAutoSheet. Application failed to close workbooks. Message and stack trace are:\r\n" + ex.Message + "\r\n" + ex.StackTrace, Level.Error);
                }
                finally
                {
                    WorksheetUtilities.ReleaseExcelApp();
                }
            }
            return returnPath;
        }

        /// <summary>
        /// Updates the Acceptance Criteria section in the worksheet with the provided values
        /// </summary>
        /// <param name="worksheet">The worksheet to update</param>
        /// <param name="acceptanceCriteria">Dictionary containing the acceptance criteria values</param>
        private static void UpdateAcceptanceCriteriaSection(Worksheet worksheet, Dictionary<string, string> acceptanceCriteria)
        {
            try
            {
                // Find the "Acceptance Criteria" section
                Range findRange = worksheet.Cells.Find("Acceptance Criteria", Type.Missing,
                    XlFindLookIn.xlValues, XlLookAt.xlWhole, XlSearchOrder.xlByRows,
                    XlSearchDirection.xlNext, false, false, Type.Missing);

                UpdateAcceptanceCriteriaSectionImpl(worksheet, acceptanceCriteria, findRange);
            }
            catch (Exception ex)
            {
                Logger.LogMessage("Error updating Acceptance Criteria section: " + ex.Message, Level.Error);
            }
        }

        //-------------------------
        //-----PRIVATE METHODS-----
        //-------------------------

        /// <summary>
        /// Method with the logic for calling / updating Excel spreadsheet for DissoManuvsAuto.
        /// </summary>
        /// <param name="sourcePath"></param>
        /// <param name=""></param>
        /// <returns></returns>
        private static string UpdateDissoManuvsAutoSheet2(string sourcePath, int compSet, int timepoints, Dictionary<string, string> acceptanceCriteria = null)
        {
            if (!File.Exists(sourcePath))
            {
                Logger.LogMessage("Error in call to DissoManuvsAuto.UpdateDissoManuvsAutoSheet. Invalid source file path specified.", Level.Error);
                return "";
            }

            // Generate an random temp path to save new workbook
            string savePath = WorksheetUtilities.CopyWorkbook(sourcePath, TempDirectoryName, "DissoManuvsAuto Results.xls");
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
                    if (acceptanceCriteria.ContainsKey("ValidationType"))
                    {
                        sheet.Cells[2, 10] = acceptanceCriteria["ValidationType"];
                    }
                    if (acceptanceCriteria.ContainsKey("CBProduct"))
                    {
                        sheet.Cells[2, 13] = acceptanceCriteria["CBProduct"];
                    }
                    UpdateAcceptanceCriteriaSection(sheet, acceptanceCriteria);
                }

                if (compSet < defaultCompSet)
                {
                    int setToDelete = defaultCompSet - compSet;

                    for (var x = 0; x < setToDelete; x++)
                    {
                        var setNumber = defaultCompSet - x;

                        //Delete Extra Comparison sets + Sampling
                        WorksheetUtilities.DeleteNamedRangeRows(sheet, "ComparisonSet" + setNumber);
                        WorksheetUtilities.DeleteNamedRangeRows(sheet, "SamplingSummary" + setNumber);
                        WorksheetUtilities.DeleteNamedRange(sheet, "SetManualDisso" + setNumber);
                        WorksheetUtilities.DeleteNamedRange(sheet, "SetAutoDisso" + setNumber);
                    }
                }

                //Handle Timepoints second as they need to point to the sets.

                if (timepoints > defaultTimepoints)
                {

                    int toInsert = timepoints - defaultTimepoints;


                    for (var i = 0; i < compSet; i++)
                    {
                        var setToPoint = i + 1;
                        WorksheetUtilities.InsertRowsIntoNamedRange(toInsert, sheet, "TimepointsTable" + setToPoint + "1", true, XlDirection.xlDown, XlPasteType.xlPasteFormulas);


                        //WorksheetUtilities.SetNamedRangeValues(sheet, "Timepoints"+setToPoint+"1", vals);

                        //As Timepoints2 & 3 just copy the values via formula of Timepoints1 just add new rows.
                        WorksheetUtilities.InsertRowsIntoNamedRange(toInsert, sheet, "TimepointsTable" + setToPoint + "2", true, XlDirection.xlDown, XlPasteType.xlPasteFormulas);
                        WorksheetUtilities.InsertRowsIntoNamedRange(toInsert, sheet, "TimepointsTable" + setToPoint + "3", true, XlDirection.xlDown, XlPasteType.xlPasteFormulas);

                    }

                    // Update the named range definitions to include the newly inserted rows
                    for (var i = 0; i < compSet; i++)
                    {
                        var setToPoint = i + 1;
                        WorksheetUtilities.ResizeNamedRange(sheet, "TimepointsTable" + setToPoint + "1", toInsert, 0);
                        WorksheetUtilities.ResizeNamedRange(sheet, "TimepointsTable" + setToPoint + "2", toInsert, 0);
                        WorksheetUtilities.ResizeNamedRange(sheet, "TimepointsTable" + setToPoint + "3", toInsert, 0);
                    }


                }

                if (timepoints < defaultTimepoints && timepoints != 0)
                {

                    var rowsToDelete = Math.Abs(timepoints - defaultTimepoints);

                    //As 04/02 - Default values should be removed
                    //Set default values for the Timepoints1 namedrange.                   
                    //List<string> vals = new List<string>();

                    //var initValue = "10";
                    //for (int x = 0; x < timepoints; x++)
                    //{
                    //    if (x == 0)
                    //    {
                    //        vals.Add(initValue);
                    //    }
                    //    else
                    //    {
                    //        var currentValue = Int32.Parse(initValue) + 10;
                    //        initValue = currentValue.ToString();
                    //        vals.Add(initValue);
                    //    }
                    //}

                    for (var i = 0; i < compSet; i++)
                    {
                        var setToPoint = i + 1;

                        WorksheetUtilities.DeleteRowsFromNamedRange(rowsToDelete, sheet, "TimepointsTable" + setToPoint + "1", XlDirection.xlDown);
                        WorksheetUtilities.DeleteRowsFromNamedRange(rowsToDelete, sheet, "TimepointsTable" + setToPoint + "2", XlDirection.xlDown);
                        WorksheetUtilities.DeleteRowsFromNamedRange(rowsToDelete, sheet, "TimepointsTable" + setToPoint + "3", XlDirection.xlDown);

                        //WorksheetUtilities.SetNamedRangeValues(sheet, "Timepoints"+setToPoint+"1", vals);

                    }

                    // No need to set specific values - the structure is already correct

                }

                if (compSet > defaultCompSet)
                {
                    for (int i = 6; i <= compSet; i++)
                    {
                        WorksheetUtilities.InsertRowsIntoNamedRange(timepoints * 2 + 16 + 2, sheet, "SampleInfoData", false, XlDirection.xlUp, XlPasteType.xlPasteAll);
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ComparisonSet" + (i - 1), "ComparisonSet" + i, timepoints * 2 + 16 + 2, 1, XlPasteType.xlPasteAll);
                        WorksheetUtilities.SetNamedRangeValue(sheet, "ComparisonSet" + i, ("Set-" + i).Trim(), 1, 1);
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "SetManualDisso" + (i - 1), "SetManualDisso" + i, timepoints * 2 + 16 + 2, 1, XlPasteType.xlPasteAll);

                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "SetAutoDisso" + (i - 1), "SetAutoDisso" + i, timepoints * 2 + 16 + 2, 1, XlPasteType.xlPasteAll);
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "Timepoints" + (i - 1) + "1", "Timepoints" + i + 1, timepoints * 2 + 16 + 2, 1, XlPasteType.xlPasteAll);
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ManualMeans" + (i - 1), "ManualMeans" + i, timepoints * 2 + 16 + 2, 1, XlPasteType.xlPasteAll);
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "AutoMeans" + (i - 1), "AutoMeans" + i, timepoints * 2 + 16 + 2, 1, XlPasteType.xlPasteAll);
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "BatchNum" + (i - 1), "BatchNum" + i, timepoints * 2 + 16 + 2, 1, XlPasteType.xlPasteAll);
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "Strength" + (i - 1), "Strength" + i, timepoints * 2 + 16 + 2, 1, XlPasteType.xlPasteAll);

                        WorksheetUtilities.InsertRowsIntoNamedRange(timepoints + 1, sheet, "SummaryData", false, XlDirection.xlUp, XlPasteType.xlPasteAll);
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "SamplingSummary" + (i - 1), "SamplingSummary" + i, timepoints + 1 + 1, 1, XlPasteType.xlPasteAll);
                        SetSamplingSummaryFormula(sheet, i);

                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "SummaryTimepoints" + (i - 1), "SummaryTimepoints" + i, timepoints + 1 + 1, 1, XlPasteType.xlPasteAll);
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "SummaryManualMeans" + (i - 1), "SummaryManualMeans" + i, timepoints + 1 + 1, 1, XlPasteType.xlPasteAll);
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "SummaryAutoMeans" + (i - 1), "SummaryAutoMeans" + i, timepoints + 1 + 1, 1, XlPasteType.xlPasteAll);


                        LinkNamedRanges(sheet, "Timepoints" + i + 1, "SummaryTimepoints" + i);
                        LinkNamedRanges(sheet, "ManualMeans" + i, "SummaryManualMeans" + i);
                        LinkNamedRanges(sheet, "AutoMeans" + i, "SummaryAutoMeans" + i);
                    }
                }

                try
                {
                    _app.Goto(sheet.Cells[1, 1], true);
                }
                catch
                {
                    Logger.LogMessage("Scroll of sheet failed in DissoManuvsAuto.DissoManuvsAuto!", Level.Error);
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

        private static string UpdateDissoManuvsAutoSheet4(string sourcePath, int compSet, int timepoints, Dictionary<string, string> acceptanceCriteria = null)
        {
            if (!File.Exists(sourcePath))
            {
                Logger.LogMessage("Error in call to DissoManuvsAuto.UpdateDissoManuvsAutoSheet. Invalid source file path specified.", Level.Error);
                return "";
            }

            // Generate an random temp path to save new workbook
            string savePath = WorksheetUtilities.CopyWorkbook(sourcePath, TempDirectoryName, "DissoManuvsAuto Results.xls");
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
                    if (acceptanceCriteria.ContainsKey("ValidationType"))
                    {
                        sheet.Cells[2, 10] = acceptanceCriteria["ValidationType"];
                    }
                    if (acceptanceCriteria.ContainsKey("CBProduct"))
                    {
                        sheet.Cells[2, 13] = acceptanceCriteria["CBProduct"];
                    }
                    UpdateAcceptanceCriteriaSection(sheet, acceptanceCriteria);
                }

                if (compSet < defaultCompSet)
                {
                    int setToDelete = defaultCompSet - compSet;

                    for (var x = 0; x < setToDelete; x++)
                    {
                        var setNumber = defaultCompSet - x;

                        //Delete Extra Comparison sets + Sampling
                        WorksheetUtilities.DeleteNamedRangeRows(sheet, "ComparisonSet" + setNumber);
                        WorksheetUtilities.DeleteNamedRangeRows(sheet, "SamplingSummary" + setNumber);
                        WorksheetUtilities.DeleteNamedRange(sheet, "SetManualDisso" + setNumber);
                        WorksheetUtilities.DeleteNamedRange(sheet, "SetAutoDisso" + setNumber);
                    }
                }

                //Handle Timepoints second as they need to point to the sets.

                if (timepoints > defaultTimepoints)
                {

                    int toInsert = timepoints - defaultTimepoints;


                    for (var i = 4; i < compSet; i++)
                    {
                        var setToPoint = i + 1;
                        WorksheetUtilities.InsertRowsIntoNamedRange(toInsert, sheet, "TimepointsTable" + setToPoint + "1", true, XlDirection.xlDown, XlPasteType.xlPasteFormulas);


                        //WorksheetUtilities.SetNamedRangeValues(sheet, "Timepoints"+setToPoint+"1", vals);

                        //As Timepoints2 & 3 just copy the values via formula of Timepoints1 just add new rows.
                        WorksheetUtilities.InsertRowsIntoNamedRange(toInsert, sheet, "TimepointsTable" + setToPoint + "2", true, XlDirection.xlDown, XlPasteType.xlPasteFormulas);
                        WorksheetUtilities.InsertRowsIntoNamedRange(toInsert, sheet, "TimepointsTable" + setToPoint + "3", true, XlDirection.xlDown, XlPasteType.xlPasteFormulas);

                    }

                    // Update the named range definitions to include the newly inserted rows
                    for (var i = 4; i < compSet; i++)
                    {
                        var setToPoint = i + 1;
                        WorksheetUtilities.ResizeNamedRange(sheet, "TimepointsTable" + setToPoint + "1", toInsert, 0);
                        WorksheetUtilities.ResizeNamedRange(sheet, "TimepointsTable" + setToPoint + "2", toInsert, 0);
                        WorksheetUtilities.ResizeNamedRange(sheet, "TimepointsTable" + setToPoint + "3", toInsert, 0);
                    }


                }

                if (timepoints < defaultTimepoints && timepoints != 0)
                {

                    var rowsToDelete = Math.Abs(timepoints - defaultTimepoints);

                    //As 04/02 - Default values should be removed
                    //Set default values for the Timepoints1 namedrange.                   
                    //List<string> vals = new List<string>();

                    //var initValue = "10";
                    //for (int x = 0; x < timepoints; x++)
                    //{
                    //    if (x == 0)
                    //    {
                    //        vals.Add(initValue);
                    //    }
                    //    else
                    //    {
                    //        var currentValue = Int32.Parse(initValue) + 10;
                    //        initValue = currentValue.ToString();
                    //        vals.Add(initValue);
                    //    }
                    //}

                    for (var i = 5; i < compSet; i++)
                    {
                        var setToPoint = i + 1;

                        WorksheetUtilities.DeleteRowsFromNamedRange(rowsToDelete, sheet, "TimepointsTable" + setToPoint + "1", XlDirection.xlDown);
                        WorksheetUtilities.DeleteRowsFromNamedRange(rowsToDelete, sheet, "TimepointsTable" + setToPoint + "2", XlDirection.xlDown);
                        WorksheetUtilities.DeleteRowsFromNamedRange(rowsToDelete, sheet, "TimepointsTable" + setToPoint + "3", XlDirection.xlDown);

                        //WorksheetUtilities.SetNamedRangeValues(sheet, "Timepoints"+setToPoint+"1", vals);

                    }

                    // No need to set specific values - the structure is already correct

                }

                if (compSet > defaultCompSet)
                {
                    int rowsToInsert = compSet - 5;
                    AppendRowsCopyOnlyFormulasAndRenumber(sheet, "Sample_Table", rowsToInsert, 1, 1, 2);

                    int headerColIndex = 1;
                    for (int i = 6; i <= compSet; i++)
                    {
                        int setNum = i - 4;
                        WorksheetUtilities.InsertRowsIntoNamedRange(timepoints * 2 + 16 + 2, sheet, "SampleInfoData", false, XlDirection.xlUp, XlPasteType.xlPasteAll);
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ComparisonSet" + (i - 1), "ComparisonSet" + i, timepoints * 2 + 16 + 2, 1, XlPasteType.xlPasteAll);
                        WorksheetUtilities.SetNamedRangeValue(sheet, "ComparisonSet" + i, ("Set-" + setNum).Trim(), 1, 1);
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "SetManualDisso" + (i - 1), "SetManualDisso" + i, timepoints * 2 + 16 + 2, 1, XlPasteType.xlPasteAll);

                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "SetAutoDisso" + (i - 1), "SetAutoDisso" + i, timepoints * 2 + 16 + 2, 1, XlPasteType.xlPasteAll);
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "Timepoints" + (i - 1) + "1", "Timepoints" + i + 1, timepoints * 2 + 16 + 2, 1, XlPasteType.xlPasteAll);
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ManualMeans" + (i - 1), "ManualMeans" + i, timepoints * 2 + 16 + 2, 1, XlPasteType.xlPasteAll);
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "AutoMeans" + (i - 1), "AutoMeans" + i, timepoints * 2 + 16 + 2, 1, XlPasteType.xlPasteAll);
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "BatchNum" + (i - 1), "BatchNum" + i, timepoints * 2 + 16 + 2, 1, XlPasteType.xlPasteAll);
                        //WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "Strength" + (i - 1), "Strength" + i, timepoints * 2 + 16 + 2, 1, XlPasteType.xlPasteAll);

                        WorksheetUtilities.InsertRowsIntoNamedRange(timepoints + 1, sheet, "SummaryData", false, XlDirection.xlUp, XlPasteType.xlPasteAll);



                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "SamplingSummary" + (i - 1), "SamplingSummary" + i, timepoints + 1 , 1, XlPasteType.xlPasteAll);
                        SetSamplingSummaryFormula(sheet, i);

                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "SummaryTimepoints" + (i - 1), "SummaryTimepoints" + i, timepoints + 1 , 1, XlPasteType.xlPasteAll);
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "SummaryManualMeans" + (i - 1), "SummaryManualMeans" + i, timepoints + 1, 1, XlPasteType.xlPasteAll);
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "SummaryAutoMeans" + (i - 1), "SummaryAutoMeans" + i, timepoints + 1, 1, XlPasteType.xlPasteAll);


                        LinkNamedRanges(sheet, "Timepoints" + i + 1, "SummaryTimepoints" + i);
                        LinkNamedRanges(sheet, "ManualMeans" + i, "SummaryManualMeans" + i);
                        LinkNamedRanges(sheet, "AutoMeans" + i, "SummaryAutoMeans" + i);

                        if (((setNum -1) % 4) == 0)
                        {
                            int newHeaderColIndex = headerColIndex + 1;
                            InsertSummaryHeaderBeforeSamplingSummary(sheet, "SamplingSummary" + i, "Sampling_Summary_Header" + headerColIndex, "Sampling_Summary_Header" + newHeaderColIndex);
                            headerColIndex++;
                        }
                    }

                    LinkBatchNumNamedRanges(sheet, compSet - 4);
                }

                try
                {
                    _app.Goto(sheet.Cells[1, 1], true);
                }
                catch
                {
                    Logger.LogMessage("Scroll of sheet failed in DissoManuvsAuto.DissoManuvsAuto!", Level.Error);
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
        
        /// <summary>
        /// Implementation of the UpdateAcceptanceCriteriaSection method
        /// </summary>
        private static void UpdateAcceptanceCriteriaSectionImpl(Worksheet worksheet, Dictionary<string, string> acceptanceCriteria, Range findRange)
        {
            if (findRange != null)
            {
                int startRow = findRange.Row;

                // Add +1 to startRow to adjust for the difference in row numbering
                startRow = startRow + 1;

                // Update the values in the Acceptance Criteria section
                // First row: Recoveries
                if (acceptanceCriteria.ContainsKey("RecoveriesOperator1") &&
                    acceptanceCriteria.ContainsKey("RecoveriesValue1"))
                {
                    worksheet.Cells[startRow + 1, 2] = acceptanceCriteria["RecoveriesOperator1"];
                    worksheet.Cells[startRow + 1, 3] = acceptanceCriteria["RecoveriesValue1"];
                }

                if (acceptanceCriteria.ContainsKey("RecoveriesOperator2") &&
                    acceptanceCriteria.ContainsKey("RecoveriesValue2"))
                {
                    worksheet.Cells[startRow + 1, 6] = acceptanceCriteria["RecoveriesOperator2"];
                    worksheet.Cells[startRow + 1, 7] = acceptanceCriteria["RecoveriesValue2"];
                }

                // Second row: Dissolved
                if (acceptanceCriteria.ContainsKey("DissolvedOperator1") &&
                    acceptanceCriteria.ContainsKey("DissolvedValue1"))
                {
                    worksheet.Cells[startRow + 2, 2] = acceptanceCriteria["DissolvedOperator1"];
                    worksheet.Cells[startRow + 2, 3] = acceptanceCriteria["DissolvedValue1"];
                }

                if (acceptanceCriteria.ContainsKey("DissolvedOperator2") &&
                    acceptanceCriteria.ContainsKey("DissolvedValue2"))
                {
                    worksheet.Cells[startRow + 2, 6] = acceptanceCriteria["DissolvedOperator2"];
                    worksheet.Cells[startRow + 2, 7] = acceptanceCriteria["DissolvedValue2"];
                }

                // Clean up
                if (findRange != null)
                {
                    WorksheetUtilities.ReleaseComObject(findRange);
                }
            }
        }

        public static void AppendRowsCopyOnlyFormulasAndRenumber(
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

        private static void InsertSummaryHeaderBeforeSamplingSummary(
            Worksheet sheet,
            string samplingSummaryName,
            string previousHeaderName,
            string newHeaderName
        )
        {

            var ssName = sheet.Names.Item(samplingSummaryName, Type.Missing, Type.Missing) as Name;
            var ssRange = ssName?.RefersToRange
                          ?? throw new InvalidOperationException($"Named range '{samplingSummaryName}' not found.");

            var prevHdr = sheet.Names.Item(previousHeaderName, Type.Missing, Type.Missing) as Name;
            var prevRange = prevHdr?.RefersToRange
                            ?? throw new InvalidOperationException($"Previous header '{previousHeaderName}' not found.");

            int insertRow = ssRange.Row;
            var rowToInsert = sheet.Rows[insertRow, Type.Missing] as Range;
            rowToInsert.Insert(XlInsertShiftDirection.xlShiftDown, Type.Missing);

            int leftCol = prevRange.Column;
            int widthCols = prevRange.Columns.Count;
            var start = sheet.Cells[insertRow, leftCol] as Range;
            var end = sheet.Cells[insertRow, leftCol + widthCols - 1] as Range;
            var newRange = sheet.Range[start, end];

            prevRange.Copy(Type.Missing);
            newRange.PasteSpecial(XlPasteType.xlPasteAll,
                XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

            string refersToLocal = "='" + sheet.Name + "'!" +
                newRange.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

            sheet.Names.Add(newHeaderName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, refersToLocal, Type.Missing, Type.Missing, Type.Missing);
        }

        private static void CreateTimepointsTablesForCompSet(Worksheet sheet, int compSetIndex, int timepoints, int startRow)
        {
            string[] tableNames = { "TimepointsTable11", "TimepointsTable12", "TimepointsTable13" };
            int offset = 0;

            foreach (string tableName in tableNames)
            {
                Name sourceName = sheet.Names.Item(tableName, Type.Missing, Type.Missing) as Name;
                if (sourceName == null) continue;

                Range sourceRange = sourceName.RefersToRange;
                int rowCount = timepoints;
                int colCount = sourceRange.Columns.Count;
                int tableStartRow = startRow + offset;

                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, tableName, $"TimepointsTable{compSetIndex}{tableName.Substring(tableName.Length - 2)}", rowCount, tableStartRow, XlPasteType.xlPasteAll);

                Name newTableName = sheet.Names.Item($"TimepointsTable{compSetIndex}{tableName.Substring(tableName.Length - 2)}", Type.Missing, Type.Missing) as Name;
                Range newTableRange = newTableName.RefersToRange;
                int currentRows = newTableRange.Rows.Count;
                int rowsToAdjust = timepoints - currentRows;

                if (rowsToAdjust > 0)
                {
                    WorksheetUtilities.InsertRowsIntoNamedRange(rowsToAdjust, sheet, $"TimepointsTable{compSetIndex}{tableName.Substring(tableName.Length - 2)}", true, XlDirection.xlDown, XlPasteType.xlPasteFormulas);
                }
                else if (rowsToAdjust < 0)
                {
                    Range rangeToDelete = sheet.Range[sheet.Cells[newTableRange.Row + currentRows + rowsToAdjust, newTableRange.Column], sheet.Cells[newTableRange.Row + currentRows - 1, newTableRange.Column + colCount - 1]];
                    rangeToDelete.Delete(XlDeleteShiftDirection.xlShiftUp);
                    WorksheetUtilities.ReleaseComObject(rangeToDelete);
                }

                Range startCell = sheet.Cells[tableStartRow, sourceRange.Column];
                Range endCell = sheet.Cells[tableStartRow + timepoints - 1, sourceRange.Column + colCount - 1];
                string newAddress = $"'{sheet.Name}'!{startCell.get_Address(false, false, XlReferenceStyle.xlA1)}:{endCell.get_Address(false, false, XlReferenceStyle.xlA1)}";
                sheet.Names.Add($"TimepointsTable{compSetIndex}{tableName.Substring(tableName.Length - 2)}", newAddress);

                WorksheetUtilities.ReleaseComObject(sourceRange);
                WorksheetUtilities.ReleaseComObject(newTableRange);
                WorksheetUtilities.ReleaseComObject(startCell);
                WorksheetUtilities.ReleaseComObject(endCell);

                offset += rowCount + 1;
            }
        }

        private static void UpdateFormulasForCompSet(Worksheet sheet, string compSetName, int compSetIndex)
        {
            Name compSetNameObj = sheet.Names.Item(compSetName, Type.Missing, Type.Missing) as Name;
            if (compSetNameObj == null) return;

            Range compSetRange = compSetNameObj.RefersToRange;
            for (int row = 1; row <= compSetRange.Rows.Count; row++)
            {
                for (int col = 1; col <= compSetRange.Columns.Count; col++)
                {
                    Range cell = compSetRange.Cells[row, col] as Range;
                    if (cell != null && cell.HasFormula)
                    {
                        string formula = cell.Formula;
                        formula = formula.Replace("TimepointsTable11", $"TimepointsTable{compSetIndex}1")
                                        .Replace("TimepointsTable12", $"TimepointsTable{compSetIndex}2")
                                        .Replace("TimepointsTable13", $"TimepointsTable{compSetIndex}3");
                        cell.Formula = formula;
                        WorksheetUtilities.ReleaseComObject(cell);
                    }
                }
            }
            WorksheetUtilities.ReleaseComObject(compSetRange);
            WorksheetUtilities.ReleaseComObject(compSetNameObj);
        }
        private static void LinkNamedRanges(Worksheet worksheet, string sourceRangeName, string destinationRangeName)
        {
            Range sourceRange = worksheet.Range[sourceRangeName];
            Range destinationRange = worksheet.Range[destinationRangeName];

            int rowCount = sourceRange.Rows.Count;

            for (int i = 1; i <= rowCount; i++)
            {
                Range srcCell = sourceRange.Cells[i, 1] as Range;
                Range destCell = destinationRange.Cells[i, 1] as Range;

                if (srcCell != null && destCell != null)
                {
                    string srcAddress = srcCell.get_Address(false, false, XlReferenceStyle.xlA1);
                    destCell.Formula = $"=IF({srcAddress}=\"\",\"\",{srcAddress})";

                    WorksheetUtilities.ReleaseComObject(srcCell);
                    WorksheetUtilities.ReleaseComObject(destCell);
                }
            }

            WorksheetUtilities.ReleaseComObject(sourceRange);
            WorksheetUtilities.ReleaseComObject(destinationRange);
        }

        private static void SetSamplingSummaryFormula(Worksheet worksheet, int i)
        {
            string samplingSummaryRangeName = $"SamplingSummary{i}";
            //string strengthName = $"Strength{i}";
            string batchNumName = $"BatchNum{i}";

            Range samplingSummaryRange = worksheet.Range[samplingSummaryRangeName];
            Range firstCell = samplingSummaryRange.Cells[1, 2] as Range;

            if (firstCell != null)
            {
                //firstCell.Formula = $"=CONCATENATE({strengthName},\" (Batch#\", {batchNumName}, \")\")";
                firstCell.Formula = $"=BatchNum{i}";
                WorksheetUtilities.ReleaseComObject(firstCell);
            }

            WorksheetUtilities.ReleaseComObject(samplingSummaryRange);
        }

        private static void LinkBatchNumNamedRanges(Worksheet sheet, int compSet)
        {

            for (int i = 1; i <= compSet; i++)
            {
                int batchNum = i + 4;
                string namedRangeName = $"BatchNum{batchNum}";
                Range batchNumRange = null;
                try
                {
                    batchNumRange = sheet.Range[namedRangeName];
                    if (batchNumRange != null)
                    {
                        int sampleRow = 10 + i;
                        batchNumRange.Formula = "=I" + sampleRow;
                    }
                }
                finally
                {
                    WorksheetUtilities.ReleaseComObject(batchNumRange);
                }
            }
        }
    }//End of Class
}
