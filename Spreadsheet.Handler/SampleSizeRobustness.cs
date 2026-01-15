using log4net.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace Spreadsheet.Handler
{
    public static class SampleSizeRobustness
    {
        private static Application _app;
        private const int DefaultNumRowsSampleInfo = 2; 
        private const int DefaultNumRows100pct = 2;
        private const int DefaultNumRowsExcluding100pct = 2;

        private const string TempDirectoryName = "ABD_TempFiles";

        /// <summary>
        /// Method to be called via Scripts - GenerateExperiment
        /// </summary>
        /// <returns></returns>
        public static string UpdateSampleSizeRobustnessSheet(string sourcePath,
            string strcmbProtocolType,
            string strcmbProductType,
            int numSamples,
            int numReps100pct,
            int numLevels,
            int numRepsExcluding100pct,
            string wcOperator1,
            decimal wcValue1,
            string wcOperator3,
            decimal wcValue2,
            string strcmbAbsoluteRelative1,
            string wcOperator2,
            decimal wcValue3,
            string wcOperator4,
            decimal wcValue4)
        {
            string returnPath = "";
            try
            {
                returnPath = UpdateSampleSizeRobustnessSheet2(
                                sourcePath,
                                strcmbProtocolType,
                                strcmbProductType,
                                numSamples,
                                numReps100pct,
                                numLevels,
                                numRepsExcluding100pct,
                                wcOperator1,
                                wcValue1,
                                wcOperator3,
                                wcValue2,
                                strcmbAbsoluteRelative1,
                                wcOperator2,
                                wcValue3,
                                wcOperator4,
                                wcValue4);
            }
            catch (Exception ex)
            {
                Logger.LogMessage("An error occurred in the call to SampleSizeRobustness.UpdateSampleSizeRobustnessSheet. Message and stack trace are:\r\n" + ex.Message + "\r\n" + ex.StackTrace, Level.Error);

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
                            Logger.LogMessage("An error occurred in the call to SampleSizeRobustness.UpdateSampleSizeRobustnessSheet. Failed to save current workbook changes and to get path.", Level.Error);
                        }

                        _app.Workbooks.Close();
                    }
                    _app = null;
                }
                catch
                {
                    Logger.LogMessage("An error occurred in the call to SampleSizeRobustness.UpdateSampleSizeRobustnessSheet. Application failed to close workbooks. Message and stack trace are:\r\n" + ex.Message + "\r\n" + ex.StackTrace, Level.Error);
                }
                finally
                {
                    WorksheetUtilities.ReleaseExcelApp();
                }
            }
            return returnPath;
        }

        /// <summary>
        /// Method with the logic for calling / updating Excel spreadsheet for SampleSizeRobustness.
        /// </summary>
        private static string UpdateSampleSizeRobustnessSheet2(string sourcePath, string strcmbProtocolType,
            string strcmbProductType, int numSamples, int numReps100pct, int numLevels, int numRepsExcluding100pct, string wcOperator1,
            decimal wcValue1,
            string wcOperator3,
            decimal wcValue2,
            string strcmbAbsoluteRelative1,
            string wcOperator2,
            decimal wcValue3,
            string wcOperator4,
            decimal wcValue4)
        {
            if (!File.Exists(sourcePath))
            {
                Logger.LogMessage("Error in call to SampleSizeRobustness.UpdateSampleSizeRobustnessSheet. Invalid source file path specified.", Level.Error);
                return "";
            }

            // Generate an random temp path to save new workbook
            string savePath = WorksheetUtilities.CopyWorkbook(sourcePath, TempDirectoryName, "Sample Size Robustness Results.xls");
            if (String.IsNullOrEmpty(savePath)) return "";

            // Try to open the file
            _app = WorksheetUtilities.GetExcelApp();
            _app.Workbooks.Open(savePath, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            Workbook book = _app.Workbooks[1];
            Worksheet sheet = book.Worksheets[1] as Worksheet;

            if (sheet != null)
            {
                bool wasProtected = WorksheetUtilities.SetSheetProtection(sheet, null, false);
                
                WorksheetUtilities.SetMetadataValues(sheet, strcmbProtocolType, strcmbProductType);

                string sampleInfoNamedRange = WorksheetUtilities.ProcessSampleInfo(sheet, strcmbProductType, numSamples);

                UpdateWorksheet(
                    sheet,
                    sampleInfoNamedRange,
                    numSamples,
                    numReps100pct,
                    numLevels,
                    numRepsExcluding100pct,
                    wcOperator1,
                    wcValue1,
                    wcOperator3,
                    wcValue2,
                    strcmbAbsoluteRelative1,
                    wcOperator2,
                    wcValue3,
                    wcOperator4,
                    wcValue4);

                // sheet.Rows.AutoFit();

                WorksheetUtilities.PostProcessSheet(sheet);
            }

            _app.Workbooks[1].Save();

            WorksheetUtilities.ReleaseComObject(book);
            _app.Workbooks.Close();

            _app = null;
            WorksheetUtilities.ReleaseExcelApp();

            // Return the path
            return savePath;
        }

        private static void UpdateWorksheet(Worksheet sheet, string sampleInfoNamedRange, int numSamples, int numReps100pct, int numLevels, int numRepsExcluding100pct, string wcOperator1,
            decimal wcValue1,
            string wcOperator3,
            decimal wcValue2,
            string strcmbAbsoluteRelative1,
            string wcOperator2,
            decimal wcValue3,
            string wcOperator4,
            decimal wcValue4)
        {
            SetAcceptanceCriteriaValues(sheet, numRepsExcluding100pct, wcOperator1,
                    wcValue1,
                    wcOperator3,
                    wcValue2,
                    strcmbAbsoluteRelative1,
                    wcOperator2,
                    wcValue3,
                    wcOperator4,
                    wcValue4);

            // vertical copies
            // Handle 100% Water Level
            SetupRawAndSummaryTables(sheet, numReps100pct, DefaultNumRows100pct, "RawPreps", "RawPercent1", "SummaryPreps", "SummaryPercent1");

            // Handle Excluding 100% Water Levels
            SetupRawAndSummaryTables(sheet, numRepsExcluding100pct, DefaultNumRowsExcluding100pct, "RawPrepsLevel1", "RawPercentLevel1", "SummaryPrepsLevel1", "SummaryPercentLevel1");

            // Create copies of multiple levels for Excluding 100% Water Levels
            // each level has 6 fixed rows and preps rows.
            for (int i = 2; i <= numLevels; i++)
            {
                int rawFixedRows = 6;
                WorksheetUtilities.InsertRowsIntoNamedRange(rawFixedRows /*fixed rows*/ + numRepsExcluding100pct + 1 /* default */, sheet, "RawLevelsData", false, XlDirection.xlUp, XlPasteType.xlPasteAll);
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"RawLevel{i - 1}", $"RawLevel{i}", rawFixedRows /*fixed rows*/ + numRepsExcluding100pct + 1 /*empty row*/ + 1 /* default */, 1, XlPasteType.xlPasteAll);

                int summaryFixedRows = 3;
                WorksheetUtilities.InsertRowsIntoNamedRange(summaryFixedRows /*fixed rows*/ + numRepsExcluding100pct + 1 /* default */, sheet, "SummaryLevelsData", false, XlDirection.xlUp, XlPasteType.xlPasteAll);
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"SummaryLevel{i - 1}", $"SummaryLevel{i}", summaryFixedRows /*fixed rows*/ + numRepsExcluding100pct + 1 /*empty row*/ + 1 /* default */, 1, XlPasteType.xlPasteAll);

                WorksheetUtilities.ResizeNamedRange(sheet, "SampleCol1", summaryFixedRows /*fixed rows*/ + numRepsExcluding100pct + 1, 0);

                LinkExcluding100pctLevelFormulas(sheet, $"RawLevel{i}", $"SummaryLevel{i}", numRepsExcluding100pct);
            }

            // horizontal copies
            GenerateSampleColCopies(sheet, numSamples);

            // Link Formulas
            // Link Sample Info rows last column (summary) to Raw table headers of 100% Water Level
            if (numSamples > 1)
            {
                WorksheetUtilities.ResizeNamedRange(sheet, "RawHeaders", 0, numSamples - 1);
            }

            LinkRawHeadersToSampleInfoLastColumn(sheet, sampleInfoNamedRange, "RawHeaders");
        }

        private static void SetAcceptanceCriteriaValues(Worksheet sheet, int numRepsExcluding100pct, string wcOperator1,
            decimal wcValue1,
            string wcOperator3,
            decimal wcValue2,
            string strcmbAbsoluteRelative1,
            string wcOperator2,
            decimal wcValue3,
            string wcOperator4,
            decimal wcValue4)
        {
            // line 1
            WorksheetUtilities.SetNamedRangeValue(sheet, "AcceptanceCriteriaRange", wcOperator1.ToString(), 1, 1);
            WorksheetUtilities.SetNamedRangeValue(sheet, "AcceptanceCriteriaRange", wcValue1.ToString(), 1, 2);
            WorksheetUtilities.SetNamedRangeValue(sheet, "AcceptanceCriteriaRange", wcOperator3.ToString(), 1, 3);
            WorksheetUtilities.SetNamedRangeValue(sheet, "AcceptanceCriteriaRange", wcValue2.ToString(), 1, 4);
            WorksheetUtilities.SetNamedRangeValue(sheet, "AcceptanceCriteriaRange", strcmbAbsoluteRelative1.ToString(), 1, 5);
            // line 2
            WorksheetUtilities.SetNamedRangeValue(sheet, "AcceptanceCriteriaRange", wcOperator2.ToString(), 2, 1);
            WorksheetUtilities.SetNamedRangeValue(sheet, "AcceptanceCriteriaRange", wcValue3.ToString(), 2, 2);
            WorksheetUtilities.SetNamedRangeValue(sheet, "AcceptanceCriteriaRange", wcOperator4.ToString(), 2, 3);
            WorksheetUtilities.SetNamedRangeValue(sheet, "AcceptanceCriteriaRange", wcValue4.ToString(), 2, 4);

            WorksheetUtilities.SetNamedRangeValue(sheet, "NumRepsExcluding100pct", numRepsExcluding100pct.ToString(), 1, 1);
        }

        private static void LinkRawHeadersToSampleInfoLastColumn(Worksheet sheet, string sampleInfoNamedRange, string rawHeadersNamedRange)
        {
            // Retrieve the named ranges
            Range sampleInfoRange = sheet.Range[sampleInfoNamedRange];
            Range rawHeadersRange = sheet.Range[rawHeadersNamedRange];

            // Get dimensions
            int sampleInfoRows = sampleInfoRange.Rows.Count;
            int sampleInfoCols = sampleInfoRange.Columns.Count;
            int rawHeadersCols = rawHeadersRange.Columns.Count;

            // Ensure dimensions match
            if (sampleInfoRows == rawHeadersCols)
            {
                for (int i = 1; i <= rawHeadersCols; i++)
                {
                    // Get cell in RawHeaders (1 row, i-th column)
                    Range rawHeaderCell = rawHeadersRange.Cells[1, i] as Range;

                    // Get cell in last column of sampleInfoRange (i-th row, last column)
                    Range sampleInfoCell = sampleInfoRange.Cells[i, sampleInfoCols] as Range;

                    // Get address of sampleInfoCell
                    string sourceAddress = sampleInfoCell.Address[false, false, XlReferenceStyle.xlA1];

                    // Set formula in RawHeaders cell
                    rawHeaderCell.Formula = $"=IF({sourceAddress}=\"\",\"\",{sourceAddress})";

                    // Release COM objects
                    WorksheetUtilities.ReleaseComObject(rawHeaderCell);
                    WorksheetUtilities.ReleaseComObject(sampleInfoCell);
                }
            }

            // Release COM objects
            WorksheetUtilities.ReleaseComObject(sampleInfoRange);
            WorksheetUtilities.ReleaseComObject(rawHeadersRange);
        }

        private static void GenerateSampleColCopies(_Worksheet sheet, int numSamples)
        {
            string prevName = "SampleCol1";

            for (int i = 2; i <= numSamples; i++)
            {
                string newName = $"SampleCol{i}";

                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, prevName, newName, 1, 2, XlPasteType.xlPasteAll);

                prevName = newName; // Update for next iteration
            }
        }

        private static void SetupRawAndSummaryTables(Worksheet sheet, int numReps, int defaultNumRows, string rawTablePrepsNamedRange, string rawTableRowsNamedRange, string summaryTablePrepsNamedrange, string summaryTableRowsNamedRange)
        {
            if (numReps > defaultNumRows)
            {
                WorksheetUtilities.InsertRowsIntoNamedRange(numReps - defaultNumRows, sheet, rawTableRowsNamedRange, true, XlDirection.xlDown, XlPasteType.xlPasteAll);
                WorksheetUtilities.InsertRowsIntoNamedRange(numReps - defaultNumRows, sheet, summaryTableRowsNamedRange, true, XlDirection.xlDown, XlPasteType.xlPasteAll);
            }
            else if (numReps < defaultNumRows)
            {
                WorksheetUtilities.DeleteRowsFromNamedRange(1, sheet, rawTableRowsNamedRange, XlDirection.xlDown);
                WorksheetUtilities.DeleteRowsFromNamedRange(1, sheet, summaryTableRowsNamedRange, XlDirection.xlDown);
            }

            // set row numbers
            List<string> list = new List<string>(0);
            for (int i = 1; i <= numReps; i++)
            {
                list.Add(i.ToString());
            }

            WorksheetUtilities.SetNamedRangeValues(sheet, rawTablePrepsNamedRange, list);
            WorksheetUtilities.SetNamedRangeValues(sheet, summaryTablePrepsNamedrange, list);
        }

        private static void LinkExcluding100pctLevelFormulas(Worksheet sheet, string rawLevelNamedRange, string summaryLevelNamedRange, int n)
        {
            // Retrieve the named ranges
            Range rawLevel = sheet.Range[rawLevelNamedRange];
            Range summaryLevel = sheet.Range[summaryLevelNamedRange];

            string rawCell1 = rawLevel.Cells[1, 1].Address[false, false, XlReferenceStyle.xlA1];
            string rawCell2 = rawLevel.Cells[1, 2].Address[false, false, XlReferenceStyle.xlA1];

            summaryLevel.Cells[1, 1].Formula = $"=CONCATENATE({rawCell2}, {rawCell1})";

            int summaryRows = summaryLevel.Rows.Count;

            // Fill formulas for DecimalPercent
            FillFixedFormulas(rawLevel, summaryLevel, 2, n + 1, "DecimalPercent");

            // Fill formulas for DecimalStats
            FillFixedFormulas(rawLevel, summaryLevel, n + 2, summaryRows, "DecimalStats");
        }

        private static void FillFixedFormulas(Range rawLevel, Range summaryLevel, int startRow, int endRow, string fixedRef)
        {
            for (int i = startRow; i <= endRow; i++)
            {
                string rawCell = rawLevel.Cells[i + 1, 2].Address[false, false, XlReferenceStyle.xlA1]; // Second column

                Range summaryCell = summaryLevel.Cells[i, 2];

                summaryCell.Formula = $"=IF({rawCell}=\"\",\" \",FIXED({rawCell}, {fixedRef}))";
            }
        }
    }//End of Class
}