using log4net.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Runtime.InteropServices;

namespace Spreadsheet.Handler
{
    public static class Identification
    {
        private static Application _app;

        private const int DefaultNumWorkingStandards = 2;
        private const int DefaultRTRTableCols = 3;
        private const int DefaultRTRSummaryCols = 1;

        private const string TempDirectoryName = "ABD_TempFiles";

        /// <summary>
        /// Method to be called via Scripts - GenerateExperiment
        /// </summary>
        /// <param name="sourcePath"></param>
        /// <returns></returns>
        public static string UpdateIdentificationSheet(
            string sourcePath,
            int numSamples,
            bool hasRetentionLevel,
            int numInjections,
            string rtr1Operator,
            decimal rtr1Value,
            string rtr2Operator,
            decimal rtr2Value,
            bool hasUVLevel,
            int numLambdaMax,
            decimal lambdaMaxValue,
            string cmbValidation,
            string cmbProduct)
        {
            string returnPath = "";
            try
            {
                returnPath = UpdateIdentificationSheet2(sourcePath, numSamples, hasRetentionLevel, numInjections, rtr1Operator, rtr1Value, rtr2Operator, rtr2Value, hasUVLevel, numLambdaMax, lambdaMaxValue, cmbValidation, cmbProduct);
            }
            catch (Exception ex)
            {
                Logger.LogMessage("An error occurred in the call to Identification.UpdateIdentificationSheet. Message and stack trace are:\r\n" + ex.Message + "\r\n" + ex.StackTrace, Level.Error);

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
                            Logger.LogMessage("An error occurred in the call to Identification.UpdateIdentificationSheet. Failed to save current workbook changes and to get path.", Level.Error);
                        }

                        _app.Workbooks.Close();
                    }
                    _app = null;
                }
                catch
                {
                    Logger.LogMessage("An error occurred in the call to Identification.UpdateIdentificationSheet. Application failed to close workbooks. Message and stack trace are:\r\n" + ex.Message + "\r\n" + ex.StackTrace, Level.Error);
                }
                finally
                {
                    WorksheetUtilities.ReleaseExcelApp();
                }
            }
            return returnPath;
        }

        /// <summary>
        /// Method with the logic for calling / updating Excel spreadsheet for Identification.
        /// </summary>
        /// <param name="sourcePath"></param>
        /// <param name=""></param>
        /// <returns></returns>
        private static string UpdateIdentificationSheet2(
            string sourcePath,
            int numSamples,
            bool hasRetentionLevel,
            int numInjections,
            string rtr1Operator,
            decimal rtr1Value,
            string rtr2Operator,
            decimal rtr2Value,
            bool hasUVLevel,
            int numLambdaMax,
            decimal lambdaMaxValue,
            string strcmbProtocolType,
            string strcmbProductType)
        {
            if (!File.Exists(sourcePath))
            {
                Logger.LogMessage("Error in call to Identification.UpdateIdentificationSheet. Invalid source file path specified.", Level.Error);
                return "";
            }

            // Generate an random temp path to save new workbook
            string savePath = WorksheetUtilities.CopyWorkbook(sourcePath, TempDirectoryName, "Identification Results.xls");
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

                if (hasRetentionLevel)
                {
                    ProcessRTR(sheet, numSamples, numInjections, rtr1Operator, rtr1Value, rtr2Operator, rtr2Value);
                }
                else
                {
                    WorksheetUtilities.DeleteRowsFromNamedRange(33, sheet, "RetentionLevel", XlDirection.xlDown);
                }

                if (hasUVLevel)
                {
                    ProcessUVSpectrum(sheet, numSamples, numLambdaMax, lambdaMaxValue);
                }
                else
                {
                    WorksheetUtilities.DeleteRowsFromNamedRange(30, sheet, "UVLevel", XlDirection.xlDown);
                }

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

        private static void ProcessRTR(Worksheet sheet, int numSamples, int numInjections, string rtr1Operator, decimal rtr1Value, string rtr2Operator, decimal rtr2Value)
        {
            WorksheetUtilities.SetNamedRangeValue(sheet, "RTR1Operator", rtr1Operator, 1, 1);
            WorksheetUtilities.SetNamedRangeValue(sheet, "RTR1Value", rtr1Value.ToString(), 1, 1);
            WorksheetUtilities.SetNamedRangeValue(sheet, "RTR2Operator", rtr2Operator, 1, 1);
            WorksheetUtilities.SetNamedRangeValue(sheet, "RTR2Value", rtr2Value.ToString(), 1, 1);

            int numRowsToInsert = numInjections - DefaultNumWorkingStandards;
            WorksheetUtilities.InsertRowsIntoNamedRange(numRowsToInsert, sheet, "RTRWorkingStandard", true, XlDirection.xlDown, XlPasteType.xlPasteAll);

            // replicate horizontally based on Number of Samples
            for (int i = 2; i <= numSamples; i++)
            {
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "RTRTable" + (i - 1), "RTRTable" + i, 1, DefaultRTRTableCols + 1, XlPasteType.xlPasteAll);
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "RTWStandardAvg" + (i - 1), "RTWStandardAvg" + i, 1, DefaultRTRTableCols + 1, XlPasteType.xlPasteAll);
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "RTWSampleAvg" + (i - 1), "RTWSampleAvg" + i, 1, DefaultRTRTableCols + 1, XlPasteType.xlPasteAll);
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "RTRAvg" + (i - 1), "RTRAvg" + i, 1, DefaultRTRTableCols + 1, XlPasteType.xlPasteAll);

                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "RTRSummary" + (i - 1), "RTRSummary" + i, 1, DefaultRTRSummaryCols + 1, XlPasteType.xlPasteAll);

                // update the formulas
                LinkRTRSummaryCells(sheet, "RTRSummary" + i, "RTRTable" + i, "RTWStandardAvg" + i, "RTWSampleAvg" + i, "RTRAvg" + i);
            }
        }

        private static void ProcessUVSpectrum(Worksheet sheet, int numSamples, int numLambdaMax, decimal lambdaMaxValue)
        {
            WorksheetUtilities.SetNamedRangeValue(sheet, "UVValue1", lambdaMaxValue.ToString(), 1, 1);
            WorksheetUtilities.SetNamedRangeValue(sheet, "UVValue2", lambdaMaxValue.ToString(), 1, 1);

            const int UVTableBaseCols = 3, UVTableColSpacing = 1;
            GenerateUVTables(sheet, UVTableBaseCols, UVTableColSpacing, numSamples, numLambdaMax, "UVTable", "UVLamdaMaxCol");

            // process UV Summary
            const int UVSummaryBaseCols = 1, UVSummaryColSpacing = 1;
            GenerateUVTables(sheet, UVSummaryBaseCols, UVSummaryColSpacing, numSamples, numLambdaMax, "UVSummary", "UVSummaryCol");

            // update formulas
            LinkUVSummaryCells(sheet, numSamples, numLambdaMax, "UVLamdaMaxCol", "UVSummaryCol");

            for (int i = 2; i <= numSamples; i++) WorksheetUtilities.LinkFirstCell(sheet, "UVTable" + i, "UVSummary" + i);
        }

        private static void GenerateUVTables(Worksheet sheet, int UVTableBaseCols, int UVTableColSpacing, int numSamples, int numLambdaMax, string uvTableBaseName, string uvLambdaMaxColBaseName)
        {
            for (int sampleIdx = 1; sampleIdx <= numSamples; sampleIdx++)
            {
                if (sampleIdx > 1)
                {
                    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"{uvLambdaMaxColBaseName}{sampleIdx - 1}a", $"{uvLambdaMaxColBaseName}{sampleIdx}a", 1, (UVTableBaseCols - 1) + numLambdaMax + 1, XlPasteType.xlPasteAll);
                }

                for (int lambdaIdx = 2; lambdaIdx <= numLambdaMax; lambdaIdx++)
                {
                    char prevChar = (char)('a' + lambdaIdx - 2);
                    char currChar = (char)('a' + lambdaIdx - 1);

                    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"{uvLambdaMaxColBaseName}{sampleIdx}{prevChar}", $"{uvLambdaMaxColBaseName}{sampleIdx}{currChar}", 1, UVTableColSpacing + 1, XlPasteType.xlPasteAll);
                    WorksheetUtilities.SetNamedRangeValue(sheet, $"{uvLambdaMaxColBaseName}{sampleIdx}{currChar}", lambdaIdx + " Lambda Max (nm)", 1, 1);

                    WorksheetUtilities.ResizeNamedRange(sheet, $"{uvTableBaseName}{sampleIdx}", 0, 1);
                }

                if (sampleIdx > 1)
                {
                    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"{uvTableBaseName}{sampleIdx - 1}", $"{uvTableBaseName}{sampleIdx}", 1, (UVTableBaseCols - 1) + numLambdaMax + 1, XlPasteType.xlPasteAll);
                }
            }
        }

        private static void LinkUVSummaryCells(Worksheet sheet, int numSamples, int numLambdaMax, string sourcePrefix, string destinationPrefix)
        {
            for (int sampleIdx = 1; sampleIdx <= numSamples; sampleIdx++)
            {
                for (int lambdaIdx = 0; lambdaIdx < numLambdaMax; lambdaIdx++)
                {
                    char lambdaChar = (char)('a' + lambdaIdx);
                    string srcName = $"{sourcePrefix}{sampleIdx}{lambdaChar}";
                    string destName = $"{destinationPrefix}{sampleIdx}{lambdaChar}";

                    Range sourceRange = sheet.Range[srcName];
                    Range destinationRange = sheet.Range[destName];
                    int rowCount = destinationRange.Rows.Count;

                    for (int row = 1; row <= rowCount; row++)
                    {
                        Range srcCell = sourceRange.Cells[row, 1] as Range;
                        Range destCell = destinationRange.Cells[row, 1] as Range;
                        string srcAddress = srcCell.Address[false, false, XlReferenceStyle.xlA1];
                        destCell.Formula = $"=IF({srcAddress}=\"\", \"\", {srcAddress})";

                        WorksheetUtilities.ReleaseComObject(srcCell);
                        WorksheetUtilities.ReleaseComObject(destCell);
                    }

                    WorksheetUtilities.ReleaseComObject(sourceRange);
                    WorksheetUtilities.ReleaseComObject(destinationRange);
                }
            }
        }

        private static void LinkRTRSummaryCells(Worksheet sheet, string summaryNamedRange, string rtrTable, string workingStandardNamedRange, string workingSampleNamedRange, string rtrAvgNamedRange)
        {
            // Helper to assign the IF-wrapped formula from a source cell to a target cell
            void LinkFormula(Range summaryRange, int row, int col, string sourceRangeName)
            {
                Range sourceRange = sheet.Range[sourceRangeName];
                Range sourceCell = sourceRange.Cells[1, 1] as Range;
                string sourceAddress = sourceCell.Address[false, false, XlReferenceStyle.xlA1];

                Range destinationCell = sheet.Range[summaryNamedRange].Cells[row, col] as Range;
                destinationCell.Formula = $"=IF({sourceAddress}=\"\", \"\", {sourceAddress})";

                WorksheetUtilities.ReleaseComObject(sourceCell);
                WorksheetUtilities.ReleaseComObject(sourceRange);
                WorksheetUtilities.ReleaseComObject(destinationCell);
            }

            // Get summary range once
            Range destRange = sheet.Range[summaryNamedRange];

            // Link specific summary cells to their respective sources
            LinkFormula(destRange, 1, 1, rtrTable);
            LinkFormula(destRange, 3, 1, workingStandardNamedRange);
            LinkFormula(destRange, 4, 1, workingSampleNamedRange);
            LinkFormula(destRange, 5, 1, rtrAvgNamedRange);

            WorksheetUtilities.ReleaseComObject(destRange);
        }
    }//End of Class
}
