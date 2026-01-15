using log4net.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace Spreadsheet.Handler
{
    public class SampleRepeatabilityDP
    {
        private static Application _app;

        private const int DefaultNumReps = 6;
        private const int MinNumReps = 2;
        private const int DefaultNumDataTables = 1;
        private const string DefaultApiDosageUnits = "mg/ml";

        // This is the raw data table column count + 1 (as a column spacer)
        private const int RawDataColOffset = 7;

        private const string TempDirectoryName = "ABD_TempFiles";

        public static string UpdateSampleRepeatabilityProductSheet(
            string sourcePath,
            int numRepsGeneral, string strcmbProtocolType, string strcmbProductType, string strcmbTestType,
            int numSamplesAssay, string strcmbRSD, decimal valRSD1,
            int numSamplesCU,
            int numSamplesImp, string strcmbNoS, int numPeaksImp, string strcmbOperator1Imp, decimal valAC1Imp, decimal valAC2Imp,
            string strAutoOperator1Imp, string strAutoOperator2Imp, decimal valacceptancecriteria1Imp, decimal valacceptancecriteria3Imp,
            string strcmbOperator2Imp, decimal valAC3Imp, decimal valAC4Imp, decimal valAC5Imp, string strcmbOperator4Imp, string strcmbOperator5Imp,
            int numSamplesWC, string strcmbWaterContent1WC, decimal valAC6WC, string strcmbWaterContent3WC, decimal valAC7WC,
            string strcmbWaterContent2WC, decimal valAC8WC, string strcmbWaterContent4WC, decimal valAC9WC,
            int numSamplesDisso, int numRepsDisso, string strcmbOperator1Disso, decimal valAC1Disso,
            string strcmbOperator2Disso, decimal valAC2Disso, string strcmbOperator3Disso,
            string strcmbOperator4Disso, decimal valAC3Disso, string strcmbOperator5Disso,
            decimal valAC4Disso, string strcmbOperator6Disso)
        {
            string returnPath = "";

            try
            {
                returnPath = UpdateSampleRepeatabilityProductSheet2(
                    sourcePath,
                    numRepsGeneral, strcmbProtocolType, strcmbProductType, strcmbTestType,
                    numSamplesAssay, strcmbRSD, valRSD1,
                    numSamplesCU,
                    numSamplesImp, strcmbNoS, numPeaksImp, strcmbOperator1Imp, valAC1Imp, valAC2Imp,
                    strAutoOperator1Imp, strAutoOperator2Imp, valacceptancecriteria1Imp, valacceptancecriteria3Imp,
                    strcmbOperator2Imp, valAC3Imp, valAC4Imp, valAC5Imp, strcmbOperator4Imp, strcmbOperator5Imp,
                    numSamplesWC, strcmbWaterContent1WC, valAC6WC, strcmbWaterContent3WC, valAC7WC,
                    strcmbWaterContent2WC, valAC8WC, strcmbWaterContent4WC, valAC9WC,
                    numSamplesDisso, numRepsDisso, strcmbOperator1Disso, valAC1Disso,
                    strcmbOperator2Disso, valAC2Disso, strcmbOperator3Disso,
                    strcmbOperator4Disso, valAC3Disso, strcmbOperator5Disso,
                    valAC4Disso, strcmbOperator6Disso
                );
            }
            catch (Exception ex)
            {
                Logger.LogMessage("An error occurred in the call to SampleRepeatability.UpdateSampleRepeatabilityProductSheet. " +
                    "Message and stack trace are:\r\n" + ex.Message + "\r\n" + ex.StackTrace, Level.Error);

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
                            Logger.LogMessage("Failed to save current workbook changes and to get path.", Level.Error);
                        }

                        _app.Workbooks.Close();
                    }

                    _app = null;
                }
                catch
                {
                    Logger.LogMessage("Application failed to close workbooks. Message and stack trace are:\r\n"
                        + ex.Message + "\r\n" + ex.StackTrace, Level.Error);
                }
                finally
                {
                    WorksheetUtilities.ReleaseExcelApp();
                }
            }

            return returnPath;
        }


        private static string UpdateSampleRepeatabilityProductSheet2(
            string sourcePath,
            int numRepsGeneral, string strcmbProtocolType, string strcmbProductType, string strcmbTestType,
            int numSamplesAssay, string strcmbRSD, decimal valRSD1,
            int numSamplesCU,
            int numSamplesImp, string strcmbNoS, int numPeaksImp, string strcmbOperator1Imp, decimal valAC1Imp, decimal valAC2Imp,
            string strAutoOperator1Imp, string strAutoOperator2Imp, decimal valacceptancecriteria1Imp, decimal valacceptancecriteria3Imp,
            string strcmbOperator2Imp, decimal valAC3Imp, decimal valAC4Imp, decimal valAC5Imp, string strcmbOperator4Imp, string strcmbOperator5Imp,
            int numSamplesWC, string strcmbWaterContent1WC, decimal valAC6WC, string strcmbWaterContent3WC, decimal valAC7WC,
            string strcmbWaterContent2WC, decimal valAC8WC, string strcmbWaterContent4WC, decimal valAC9WC,
            int numSamplesDisso, int numRepsDisso, string strcmbOperator1Disso, decimal valAC1Disso,
            string strcmbOperator2Disso, decimal valAC2Disso, string strcmbOperator3Disso,
            string strcmbOperator4Disso, decimal valAC3Disso, string strcmbOperator5Disso,
            decimal valAC4Disso, string strcmbOperator6Disso, bool chkAssay = true, bool chkImpurity = true, 
            bool chkContentUniformity = true, bool chkWaterContent = true, bool chkDissolution  =true)
        {
            //if ((chkAssay || chkContentUniformity || chkDissolution))
            //{
            //    Logger.LogMessage("Error in call to SampleRepeatability.UpdateSampleRepeatabilityProductSheet2. Dosage strengths parameter is empty!", Level.Error);
            //    return "";
            //}

            if (!File.Exists(sourcePath))
            {
                Logger.LogMessage("Error in call to SampleRepeatability.UpdateSampleRepeatabilityProductSheet2. Invalid source file path specified.", Level.Error);
                return "";
            }

            // Generate an random temp path to save new workbook
            string savePath = WorksheetUtilities.CopyWorkbook(sourcePath, TempDirectoryName, "Sample Repeatability Product Results.xls");
            if (String.IsNullOrEmpty(savePath)) return "";

            // Try to open the file
            _app = WorksheetUtilities.GetExcelApp();
            _app.Workbooks.Open(savePath, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            Workbook book = _app.Workbooks[1];
            Worksheet sheet = book.Worksheets[1] as Worksheet;

            if (sheet != null)
            {
                bool wasProtected = WorksheetUtilities.SetSheetProtection(sheet, null, false);

                WorksheetUtilities.SetMetadataValues(sheet, strcmbProtocolType, strcmbProductType, strcmbTestType);

                if (chkAssay) HandleAssayProduct(sheet, numRepsGeneral, numSamplesAssay);

                if (chkImpurity) HandleImpurityProduct(sheet, numRepsGeneral, numPeaksImp, numSamplesImp);

                if (chkContentUniformity) HandleContentUniformity(sheet, numSamplesCU);

                if (chkWaterContent) HandleWaterContent();

                if (chkDissolution) HandleDissolution(sheet, numRepsDisso, numSamplesDisso);

                //Remove Assay name range if Assay checkbox is false
                if (chkAssay != true)
                {
                    WorksheetUtilities.DeleteNamedRangeRows(sheet, "Assay");
                    WorksheetUtilities.DeleteNamedRange(sheet, "Assay");
                }
                //Remove Impurity name range if Assay checkbox is false
                if (chkImpurity != true)
                {
                    WorksheetUtilities.DeleteNamedRangeRows(sheet, "Impurity");
                    WorksheetUtilities.DeleteNamedRange(sheet, "Impurity");
                }
                //Remove Content Uniformity name range if Assay checkbox is false
                if (chkContentUniformity != true)
                {
                    WorksheetUtilities.DeleteNamedRangeRows(sheet, "CU");
                    WorksheetUtilities.DeleteNamedRange(sheet, "CU");
                }

                if (chkContentUniformity != true)
                {
                    // delete named ranges of Water Content.
                }

                //Remove Dissolution name range if Assay checkbox is false
                if (chkDissolution != true)
                {
                    WorksheetUtilities.DeleteNamedRangeRows(sheet, "Dissolution");
                    WorksheetUtilities.DeleteNamedRange(sheet, "Dissolution");
                }

                WorksheetUtilities.PostProcessSheet(sheet);
            }
            _app.Workbooks[1].Save();

            WorksheetUtilities.ReleaseComObject(book);
            _app.Workbooks.Close();

            //WorksheetUtilities.ReleaseComObject(_app);
            _app = null;
            WorksheetUtilities.ReleaseExcelApp();

            // Return the path
            return savePath;
        }

        private static void HandleWaterContent()
        {
        }

        private static void HandleAssayProduct(Worksheet sheet, int numReps, int numSamples)
        {
            // Expand or contract the tables based on the number of replicates
            if (numReps > DefaultNumReps)
            {
                int numRowsToInsert = numReps - DefaultNumReps;
                WorksheetUtilities.InsertRowsIntoNamedRange(numRowsToInsert, sheet, "RawDataValues1", true, XlDirection.xlDown, XlPasteType.xlPasteFormulas);
                WorksheetUtilities.InsertRowsIntoNamedRange(numRowsToInsert, sheet, "ValidationTableValues1", true, XlDirection.xlDown, XlPasteType.xlPasteFormulas);
            }
            else if (numReps < DefaultNumReps)
            {
                int numRowsToRemove = DefaultNumReps - numReps;
                // There needs to be at least 2 rows in order to not corrupt the sheet's formulas
                if (DefaultNumReps - numRowsToRemove < MinNumReps) numRowsToRemove = DefaultNumReps - MinNumReps;
                WorksheetUtilities.DeleteRowsFromNamedRange(numRowsToRemove, sheet, "RawDataValues1", XlDirection.xlDown);
                WorksheetUtilities.DeleteRowsFromNamedRange(numRowsToRemove, sheet, "ValidationTableValues1", XlDirection.xlDown);
            }

            // Re-number the preps in the rawdata and validation results tables
            if (numReps < MinNumReps) numReps = MinNumReps;
            List<string> prepNumbers = new List<string>(0);
            for (int i = 1; i <= numReps; i++) prepNumbers.Add(i.ToString());
            WorksheetUtilities.SetNamedRangeValues(sheet, "PrepNumsRawData1", prepNumbers);
            WorksheetUtilities.SetNamedRangeValues(sheet, "PrepNumsValidationResults", prepNumbers);

            // Handle the dosage strengths and insert additionals if needed
            if (numSamples > DefaultNumDataTables)
            {
                int numDataTables = numSamples - DefaultNumDataTables;
                for (int i = 1; i <= numDataTables; i++)
                {
                    // Copy the named ranges as needed for each data table
                    int namedRangeNum = i + 1;
                    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "RawDataTable1", "RawDataTable" + namedRangeNum, 1, ((RawDataColOffset - 1) * i) + (i + 1), XlPasteType.xlPasteAll);
                    WorksheetUtilities.SetNamedRangeValue(sheet, "RawDataTable" + namedRangeNum, "Raw Data Table " + namedRangeNum, 1, 2);
                    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "DosageStrengthRawData1", "DosageStrengthRawData" + namedRangeNum, 1, ((RawDataColOffset - 1) * i) + (i + 1), XlPasteType.xlPasteAll);
                    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ValidDosageStrength1", "ValidDosageStrength" + namedRangeNum, 1, i + 1, XlPasteType.xlPasteAll);
                    WorksheetUtilities.ResizeNamedRange(sheet, "ValidationTable", 0, 1);

                    // Expand named ranges after the first iteration
                    if (i > 1)
                    {
                        WorksheetUtilities.ResizeNamedRange(sheet, "DosageStrengthTable", 0, 1);
                        WorksheetUtilities.ResizeNamedRange(sheet, "PrepsTable", 0, 1);
                    }
                }

                // Expand the results table
                WorksheetUtilities.InsertRowsIntoSingleRowedNamedRange(numDataTables, sheet, "ResultsTable", false, XlPasteType.xlPasteFormulas);
                UpdateResultTableFormulas(sheet, numDataTables);
                UpdateResultsPercentRsdValsRange(sheet);
                UpdateDosageStrengthTableFormulas(sheet, numSamples, numReps);
            }

            // Replace Comprehensive Validation Results column headers with dosage strength + units
            //for (int i = 1; i <= dosageStrengths.Count; i++)
            //{
            //    WorksheetUtilities.SetNamedRangeValue(sheet, "RawDataTable" + i, dosageStrengths[i - 1], 2, 2);
            //    WorksheetUtilities.SetNamedRangeValue(sheet, "ValidDosageStrength" + i, dosageStrengths[i - 1], 1, 1);
            //}

            WorksheetUtilities.DeleteNamedRangeRows(sheet, "Upper95Confidence");
            WorksheetUtilities.DeleteNamedRange(sheet, "Upper95Confidence");
            WorksheetUtilities.DeleteNamedRangeRows(sheet, "FullResultTable");
            WorksheetUtilities.DeleteNamedRange(sheet, "FullResultTable");
        }

        private static void HandleImpurityProduct(Worksheet sheet, int numReps, int numImp, int numSamples)
        {
            if (numImp <= 0)
            {
                WorksheetUtilities.DeleteNamedRangeRows(sheet, "ImpurityAndImpuritySummary");
                WorksheetUtilities.DeleteNamedRange(sheet, "ImpurityAndImpuritySummary");
                return;
            }

            // Expand or contract the tables based on the number of replicates
            int DefaultNumImps = 1;
            int numRowsToInsert = numReps - DefaultNumReps;
            if (numReps > DefaultNumReps)
            {
                WorksheetUtilities.InsertRowsIntoNamedRange(numRowsToInsert, sheet, "PrepNumsImpurityRawData1", true, XlDirection.xlDown, XlPasteType.xlPasteFormulas);
            }
            else if (numReps < DefaultNumReps)
            {
                int numRowsToRemove = DefaultNumReps - numReps;
                // There needs to be at least 2 rows in order to not corrupt the sheet's formulas
                if (DefaultNumReps - numRowsToRemove < MinNumReps) numRowsToRemove = DefaultNumReps - MinNumReps;
                WorksheetUtilities.DeleteRowsFromNamedRange(numRowsToRemove, sheet, "PrepNumsImpurityRawData1", XlDirection.xlDown);
            }

            // Re-number the preps in the rawdata and validation results tables
            if (numReps < MinNumReps) numReps = MinNumReps;
            List<string> prepNumbers = new List<string>(0);
            for (int i = 1; i <= numReps; i++) prepNumbers.Add(i.ToString());
            WorksheetUtilities.SetNamedRangeValues(sheet, "PrepNumsImpurityRawData1", prepNumbers);

            //Repeat the table with respect to number of Impurities
            int numDataTables = numImp - DefaultNumDataTables;
            for (int i = 1; i <= numDataTables; i++)
            {
                // Insert additional columns into the Data range
                int colOffset = 2;

                // Copy the named ranges as needed for each data table
                int namedRangeNum = i + 1;
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ImpurityRawDataTable1", "ImpurityRawDataTable" + namedRangeNum, 1, (colOffset * i) + 1, XlPasteType.xlPasteAll);
                //WorksheetUtilities.SetNamedRangeValue(sheet, "ImpurityRawDataTable" + namedRangeNum, "Impurity " + namedRangeNum, 1, 1);
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ImpurityResults1", "ImpurityResults" + namedRangeNum, 1, (colOffset * i) + 1, XlPasteType.xlPasteAll);
            }

            //Repeat the table with respect to number of Sample
            if (numSamples > 1)
            {
                int rowOffset = 11 + numReps;
                for (int i = 2; i <= numSamples; i++)
                {
                    WorksheetUtilities.InsertRowsAfterForNamedRange(sheet, "ImpurityFullRawDataTable" + (i - 1).ToString(), rowOffset, "ImpurityFullRawDataTable" + (i - 1).ToString(), true, XlPasteType.xlPasteAll, "ImpurityFullRawDataTable" + i.ToString());

                    //Repeat the table with respect to number of Impurities
                    numDataTables = numImp; // - DefaultNumDataTables;
                    for (int j = 1; j <= numDataTables; j++)
                    {
                        // Insert additional columns into the Data range
                        int colOffset = 2;

                        // Copy the named ranges as needed for each data table

                        int namedRangeNum = (numImp * (i - 1)) + j;

                        //int newRowOffset = -(rowOffset - 1);
                        int newRowOffset = rowOffset * i - (rowOffset + i - 2);

                        //int rowPos = newRowOffset + 11 + numReps;
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ImpurityRawDataTable1", "ImpurityRawDataTable" + namedRangeNum, newRowOffset, (colOffset * j) - 1, XlPasteType.xlPasteAll);
                        //WorksheetUtilities.SetNamedRangeValue(sheet, "ImpurityRawDataTable" + namedRangeNum, "Impurity " + namedRangeNum, 1, 1);
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ImpurityResults1", "ImpurityResults" + namedRangeNum, newRowOffset, (colOffset * j) - 1, XlPasteType.xlPasteAll);
                    }
                }
            }

            // Expand or contract the tables based on the number of Impurities Summary
            numRowsToInsert = numImp - DefaultNumImps;
            int rowOffset1 = 2;
            for (int i = DefaultNumImps + 1; i <= numImp; i++)
            {
                WorksheetUtilities.InsertRowsAfterForNamedRange(sheet, "ImpuritySummary" + (i - 1).ToString(), rowOffset1, "ImpuritySummary" + (i - 1).ToString(), true, XlPasteType.xlPasteAll, "ImpuritySummary" + i.ToString());
            }

            for (int i = 2; i <= numSamples; i++)
            {
                WorksheetUtilities.CopyNamedRangeToNewLocation(sheet, "ImpuritySummaryTitle" + (i - 1).ToString(), "ImpuritySummaryTitle" + i.ToString(), numImp + 1, 0, XlPasteType.xlPasteAll);

                //Copy the Impurity Summary with respect to number of Impuities
                //int intSequence = (i * numImp) - 2;
                int intSequence = (i * numImp) - numImp;

                WorksheetUtilities.CopyNamedRangeToNewLocation(sheet, "ImpuritySummary" + intSequence.ToString(), "ImpuritySummary" + (intSequence + 1).ToString(), 2, 0, XlPasteType.xlPasteAll);

                for (int j = 2; j <= numImp; j++)
                {
                    intSequence += 1;
                    WorksheetUtilities.CopyNamedRangeToNewLocation(sheet, "ImpuritySummary" + intSequence.ToString(), "ImpuritySummary" + (intSequence + 1).ToString(), 1, 0, XlPasteType.xlPasteAll);
                }
            }

            //UpdateImpuritySummaryTableFormulas(sheet, (numImp * numSamples), "");
            UpdateFormulasInNamedRange(sheet, "ImpuritySummary", (numImp * numSamples), "ImpurityResults1");
        }

        private static void HandleContentUniformity(Worksheet sheet, int numSamples)
        {
            // Handle the dosage strengths and insert additionals if needed
            if (numSamples > DefaultNumDataTables)
            {
                int numDataTables = numSamples - DefaultNumDataTables;
                for (int i = 1; i <= numDataTables; i++)
                {
                    // Insert additional columns into the Data range
                    //WorksheetUtilities.InsertColumnsIntoNamedRange(RawDataColOffset, sheet, "Data", XlDirection.xlToRight);
                    //}

                    // Copy the named ranges as needed for each data table
                    int namedRangeNum = i + 1;
                    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "CURawDataTable1", "CURawDataTable" + namedRangeNum, 1, ((RawDataColOffset) * i) + 1, XlPasteType.xlPasteAll);
                }

                // Expand the results table
                for (int i = 2; i <= numSamples; i++)
                {
                    WorksheetUtilities.InsertRowsAfterForNamedRange(sheet, "CUSummary" + (i - 1).ToString(), 2, "CUSummary" + (i - 1).ToString(), true, XlPasteType.xlPasteAll, "CUSummary" + i.ToString());
                }

                UpdateFormulasInCUNamedRange(sheet, "CUSummary", numSamples, "CURawDataTable");
            }
        }

        private static void HandleDissolution(Worksheet sheet, int numTimePoints, int numSamples)
        {
            // Expand or contract the tables based on the number of Time Points
            int DefaultNumTimePoints = 5;
            int numRowsToInsert = numTimePoints - DefaultNumTimePoints;

            if (numTimePoints > DefaultNumTimePoints)
            {
                // Changed insert direction as first row has specific formatting
                WorksheetUtilities.InsertRowsIntoNamedRange(numRowsToInsert, sheet, "DissoTimePointTable1", true, XlDirection.xlUp, XlPasteType.xlPasteFormulasAndNumberFormats);
                UpdateDissolutionFormulas(sheet, "DissoTable1Formulas", "DissoDecimalsRow", 1, 1);
            }
            else if (numTimePoints < DefaultNumTimePoints)
            {
                int numRowsToRemove = DefaultNumTimePoints - numTimePoints;
                // There needs to be at least 2 rows in order to not corrupt the sheet's formulas
                WorksheetUtilities.DeleteRowsFromNamedRange(numRowsToRemove, sheet, "DissoTimePointTable1", XlDirection.xlUp);
            }

            // Repeat the table with respect to number of Samples
            if (numSamples > 1)
            {
                int rowOffset = 11 + numTimePoints;
                for (int i = 2; i <= numSamples; i++)
                {
                    WorksheetUtilities.InsertRowsAfterForNamedRange(sheet, "DissoRawDataTable" + (i - 1).ToString(), rowOffset, "DissoRawDataTable" + (i - 1).ToString(), true, XlPasteType.xlPasteAll, "DissoRawDataTable" + i.ToString());
                }
            }

            if (numSamples > 1)
            {
                int rowOffset = 2;
                for (int i = 2; i <= numSamples; i++)
                {
                    WorksheetUtilities.InsertRowsAfterForNamedRange(sheet, "DissoSummary" + (i - 1).ToString(), rowOffset, "DissoSummary" + (i - 1).ToString(), true, XlPasteType.xlPasteAllExceptBorders, "DissoSummary" + i.ToString());
                }

                UpdateFormulasInNamedRange(sheet, "DissoSummary", numSamples, "DissoRawDataTable1", "DissoRawDataTable");
            }
        }

        /// <summary>
        /// Updates the formulas within the ResultsTable range to use the correct cell addresses.
        /// Specifically, the dosage strength, rsd, and upper 95 are updated/corrected.
        /// </summary>
        /// <param name="sheet">The sheet</param>
        /// <param name="numDataTables">The number of data tables.</param>
        private static void UpdateResultTableFormulas(_Worksheet sheet, int numDataTables)
        {
            if (sheet == null || numDataTables <= 0) return;

            const int SRC_TABLE_DOSAGE_ROW_INDEX = 2;
            const int SRC_TABLE_DOSAGE_COL_INDEX = 2;
            const int SRC_TABLE_STATS_COL_INDEX = 4;

            const int DEST_TABLE_DOSAGE_COL_INDEX = 1;
            const int DEST_TABLE_RSD_COL_INDEX = 2;
            const int DEST_TABLE_UPPER95_COL_INDEX = 3;

            Name resultTableName = null;
            Range resultsTableRange = null;
            object objRawDataNamedRange = null;
            Name rawDataTableName = null;
            Range rawDataTableRange = null;
            Range srcCell = null;
            Range destCell = null;
            Range resultsTableRow = null;

            // Get the range by name
            object objResultsTableNamedRange = sheet.Names.Item("ResultsTable", Type.Missing, Type.Missing);
            if (!(objResultsTableNamedRange is Name)) goto Cleanup_UpdateResultsTableFormulas;

            resultTableName = objResultsTableNamedRange as Name;
            resultsTableRange = resultTableName.RefersToRange;

            for (int i = 2; i <= numDataTables + 1; i++)
            {
                objRawDataNamedRange = sheet.Names.Item("RawDataTable" + i, Type.Missing, Type.Missing);
                if (!(objRawDataNamedRange is Name)) continue;

                resultsTableRow = resultsTableRange.Rows[i, Type.Missing] as Range;

                rawDataTableName = objRawDataNamedRange as Name;
                rawDataTableRange = rawDataTableName.RefersToRange;

                // Set the first cell of the current row - dosage strength
                srcCell = rawDataTableRange.Cells[SRC_TABLE_DOSAGE_ROW_INDEX, SRC_TABLE_DOSAGE_COL_INDEX] as Range;
                if (srcCell != null && resultsTableRow != null)
                {
                    string cellAddress = srcCell.get_Address(Type.Missing, Type.Missing, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                    cellAddress = cellAddress.Replace("$", "");

                    destCell = resultsTableRow.Cells[1, DEST_TABLE_DOSAGE_COL_INDEX] as Range;
                    if (destCell != null)
                    {
                        destCell.Value2 = "=" + cellAddress;
                        WorksheetUtilities.ReleaseComObject(destCell);
                    }
                    WorksheetUtilities.ReleaseComObject(srcCell);
                }

                // Set the second cell of the current row - %rsd
                srcCell = rawDataTableRange.Cells[rawDataTableRange.Rows.Count - 1, SRC_TABLE_STATS_COL_INDEX] as Range;
                if (srcCell != null && resultsTableRow != null)
                {
                    string cellAddress = srcCell.get_Address(Type.Missing, Type.Missing, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                    cellAddress = cellAddress.Replace("$", "");

                    destCell = resultsTableRow.Cells[1, DEST_TABLE_RSD_COL_INDEX] as Range;
                    if (destCell != null)
                    {
                        destCell.Value2 = String.Format("=IF({0}=\"\",\" \",{0})", cellAddress);
                        WorksheetUtilities.ReleaseComObject(destCell);
                    }

                    WorksheetUtilities.ReleaseComObject(srcCell);
                }

                // Set the third cell of the current row - upper 95%
                srcCell = rawDataTableRange.Cells[rawDataTableRange.Rows.Count, SRC_TABLE_STATS_COL_INDEX] as Range;
                if (srcCell != null && resultsTableRow != null)
                {
                    string srcCellAddress = srcCell.get_Address(Type.Missing, Type.Missing, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                    srcCellAddress = srcCellAddress.Replace("$", "");

                    destCell = resultsTableRow.Cells[1, DEST_TABLE_UPPER95_COL_INDEX] as Range;
                    if (destCell != null)
                    {
                        //string destCellAddress = destCell.get_Address(Type.Missing, Type.Missing, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

                        destCell.Value2 = String.Format("=IF({0}=\"\",\" \",{0})", srcCellAddress);
                        WorksheetUtilities.ReleaseComObject(destCell);
                    }

                    WorksheetUtilities.ReleaseComObject(srcCell);
                }

                // Clean up
                try
                {
                    WorksheetUtilities.ReleaseComObject(resultsTableRow);
                    WorksheetUtilities.ReleaseComObject(rawDataTableRange);
                    WorksheetUtilities.ReleaseComObject(rawDataTableName);
                    WorksheetUtilities.ReleaseComObject(objRawDataNamedRange);
                }
                catch
                {
                    continue;
                }
            }

        Cleanup_UpdateResultsTableFormulas:
            {
                try
                {
                    WorksheetUtilities.ReleaseComObject(objResultsTableNamedRange);
                    WorksheetUtilities.ReleaseComObject(resultTableName);
                    WorksheetUtilities.ReleaseComObject(resultsTableRange);
                    WorksheetUtilities.ReleaseComObject(objRawDataNamedRange);
                    WorksheetUtilities.ReleaseComObject(rawDataTableName);
                    WorksheetUtilities.ReleaseComObject(rawDataTableRange);
                    WorksheetUtilities.ReleaseComObject(srcCell);
                    WorksheetUtilities.ReleaseComObject(destCell);
                    WorksheetUtilities.ReleaseComObject(resultsTableRow);
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
        /// Expands the ResultsPercentRsdVals named range based on the current number of rows in the ResultsTable range.
        /// </summary>
        /// <param name="sheet">The sheet</param>
        private static void UpdateResultsPercentRsdValsRange(_Worksheet sheet)
        {
            if (sheet == null) return;

            const int COL_INDEX = 2;

            Name resultTableName = null;
            Range resultsTableRange = null;
            Range startCell = null;
            Range endCell = null;

            // Get the range by name
            object objResultsTableNamedRange = sheet.Names.Item("ResultsTable", Type.Missing, Type.Missing);
            if (!(objResultsTableNamedRange is Name)) goto Cleanup_UpdateResultsTableFormulas;

            resultTableName = objResultsTableNamedRange as Name;
            resultsTableRange = resultTableName.RefersToRange;

            // Update the named range to point to the new address
            startCell = resultsTableRange.Cells[1, COL_INDEX] as Range;
            endCell = resultsTableRange.Cells[resultsTableRange.Rows.Count, COL_INDEX] as Range;
            if (startCell != null && endCell != null)
            {
                //string startCellAddress = startCell.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                //string endCellAddress = endCell.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

                string refersToLocal = "='" + sheet.Name + "'!" +
                                 startCell.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing) + ":" +
                                 endCell.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

                // Update the range
                sheet.Names.Add("ResultsPercentRsdVals", Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, refersToLocal, Type.Missing, Type.Missing, Type.Missing);
                //namedRange.RefersToLocal = refersToLocal;
            }

        Cleanup_UpdateResultsTableFormulas:
            {
                try
                {
                    WorksheetUtilities.ReleaseComObject(objResultsTableNamedRange);
                    WorksheetUtilities.ReleaseComObject(resultTableName);
                    WorksheetUtilities.ReleaseComObject(resultsTableRange);
                    WorksheetUtilities.ReleaseComObject(startCell);
                    WorksheetUtilities.ReleaseComObject(endCell);
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
        /// Updates the ValidDosageStrengthN ranges to point to the correct raw data cell addresses.
        /// </summary>
        /// <param name="sheet">The sheet.</param>
        /// <param name="numDoses">The number of doses.</param>
        /// <param name="numReps">The number of replicates.</param>
        private static void UpdateDosageStrengthTableFormulas(_Worksheet sheet, int numDoses, int numReps)
        {
            if (sheet == null || numDoses <= 0 || numReps <= 0) return;

            const int RAW_DATA_PREPS_ROW = 6;
            const int RAW_DATA_PREPS_COL = 4;
            const int DOSAGE_PREPS_ROW = 3;

            object objRawDataNamedRange = null;
            Name rawDataTableName = null;
            Range rawDataTableRange = null;
            object objDosageStrengthNamedRange = null;
            Name dosageStrengthName = null;
            Range dosageStrengthRange = null;
            Range srcCell = null;
            Range destCell = null;

            for (int i = 2; i <= numDoses; i++)
            {
                // Get raw data table range by name
                objRawDataNamedRange = sheet.Names.Item("RawDataTable" + i, Type.Missing, Type.Missing);
                if (!(objRawDataNamedRange is Name)) goto Cleanup_UpdateResultsTableFormulas;
                rawDataTableName = objRawDataNamedRange as Name;
                rawDataTableRange = rawDataTableName.RefersToRange;

                // Get dosage strength range by name
                objDosageStrengthNamedRange = sheet.Names.Item("ValidDosageStrength" + i, Type.Missing, Type.Missing);
                if (!(objDosageStrengthNamedRange is Name)) continue;
                dosageStrengthName = objDosageStrengthNamedRange as Name;
                dosageStrengthRange = dosageStrengthName.RefersToRange;

                //for (int j = 0; j < numReps; j++)
                for (int j = 0; j < dosageStrengthRange.Rows.Count - 2; j++)
                {
                    srcCell = rawDataTableRange.Cells[j + RAW_DATA_PREPS_ROW, RAW_DATA_PREPS_COL] as Range;
                    if (srcCell != null)
                    {
                        string srcCellAddress = srcCell.get_Address(Type.Missing, Type.Missing, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                        srcCellAddress = srcCellAddress.Replace("$", "");

                        destCell = dosageStrengthRange.Cells[j + DOSAGE_PREPS_ROW, 1] as Range;
                        if (destCell != null)
                        {
                            //string destCellAddress = destCell.get_Address(Type.Missing, Type.Missing, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

                            destCell.Value2 = String.Format("=IF({0}=\"\",\" \",{0})", srcCellAddress);
                            WorksheetUtilities.ReleaseComObject(destCell);
                        }

                        WorksheetUtilities.ReleaseComObject(srcCell);
                    }
                }

                // Clean up
                try
                {
                    WorksheetUtilities.ReleaseComObject(objRawDataNamedRange);
                    WorksheetUtilities.ReleaseComObject(rawDataTableRange);
                    WorksheetUtilities.ReleaseComObject(rawDataTableName);
                    WorksheetUtilities.ReleaseComObject(objDosageStrengthNamedRange);
                    WorksheetUtilities.ReleaseComObject(dosageStrengthRange);
                    WorksheetUtilities.ReleaseComObject(dosageStrengthName);
                }
                catch
                {
                    continue;
                }
            }

        Cleanup_UpdateResultsTableFormulas:
            {
                try
                {
                    WorksheetUtilities.ReleaseComObject(objRawDataNamedRange);
                    WorksheetUtilities.ReleaseComObject(rawDataTableName);
                    WorksheetUtilities.ReleaseComObject(rawDataTableRange);
                    WorksheetUtilities.ReleaseComObject(objDosageStrengthNamedRange);
                    WorksheetUtilities.ReleaseComObject(dosageStrengthName);
                    WorksheetUtilities.ReleaseComObject(dosageStrengthRange);
                    WorksheetUtilities.ReleaseComObject(srcCell);
                    WorksheetUtilities.ReleaseComObject(destCell);
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

        private static void UpdateFormulasInNamedRange(_Worksheet sheet, string NamedRange, int NumberOfNamedRange, string sourceString)
        {
            Name sourceTableName = null;
            Range sourceTableRange = null;
            Range destCell = null;

            for (int i = 2; i <= NumberOfNamedRange; i++)
            {
                // Get the range by name
                object objSourceTableNamedRange = sheet.Names.Item(NamedRange + i.ToString(), Type.Missing, Type.Missing);
                if (!(objSourceTableNamedRange is Name)) goto Cleanup_UpdateValidationResultsTableFormulas;
                sourceTableName = objSourceTableNamedRange as Name;
                sourceTableRange = sourceTableName.RefersToRange;
                for (int iRow = 1; iRow <= sourceTableRange.Rows.Count; iRow++)
                {
                    for (int iCol = 1; iCol <= sourceTableRange.Columns.Count; iCol++)
                    {
                        destCell = sourceTableRange.Cells[iRow, iCol] as Range;
                        if (destCell != null && destCell.HasFormula)
                        {
                            destCell.Formula = destCell.Formula.Replace("ImpurityRawDataTable1", "ImpurityRawDataTable" + i.ToString());
                            destCell.Formula = destCell.Formula.Replace("ImpurityResults1", "ImpurityResults" + i.ToString());
                        }
                    }
                }
            }

        Cleanup_UpdateValidationResultsTableFormulas:
            {
                try
                {
                    WorksheetUtilities.ReleaseComObject(destCell);
                    WorksheetUtilities.ReleaseComObject(sourceTableRange);
                    WorksheetUtilities.ReleaseComObject(sourceTableName);
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

        private static void UpdateFormulasInCUNamedRange(_Worksheet sheet, string NamedRange, int NumberOfNamedRange, string sourceString)
        {
            Name sourceTableName = null;
            Range sourceTableRange = null;
            Range destCell = null;

            for (int i = 2; i <= NumberOfNamedRange; i++)
            {
                // Get the range by name
                object objSourceTableNamedRange = sheet.Names.Item(NamedRange + i.ToString(), Type.Missing, Type.Missing);
                if (!(objSourceTableNamedRange is Name)) goto Cleanup_UpdateValidationResultsTableFormulas;
                sourceTableName = objSourceTableNamedRange as Name;
                sourceTableRange = sourceTableName.RefersToRange;
                for (int iRow = 1; iRow <= sourceTableRange.Rows.Count; iRow++)
                {
                    for (int iCol = 1; iCol <= sourceTableRange.Columns.Count; iCol++)
                    {
                        destCell = sourceTableRange.Cells[iRow, iCol] as Range;
                        if (destCell != null && destCell.HasFormula)
                        {
                            destCell.Formula = destCell.Formula.Replace("CURawDataTable1", "CURawDataTable" + i.ToString());
                        }
                    }
                }
            }

        Cleanup_UpdateValidationResultsTableFormulas:
            {
                try
                {
                    WorksheetUtilities.ReleaseComObject(destCell);
                    WorksheetUtilities.ReleaseComObject(sourceTableRange);
                    WorksheetUtilities.ReleaseComObject(sourceTableName);
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

        private static void UpdateFormulasInNamedRange(_Worksheet sheet, string NamedRange, int NumberOfNamedRange, string sourceString, string destString)
        {
            Name sourceTableName = null;
            Range sourceTableRange = null;
            Range destCell = null;

            for (int i = 2; i <= NumberOfNamedRange; i++)
            {
                // Get the range by name
                object objSourceTableNamedRange = sheet.Names.Item(NamedRange + i.ToString(), Type.Missing, Type.Missing);
                if (!(objSourceTableNamedRange is Name)) goto Cleanup_UpdateValidationResultsTableFormulas;
                sourceTableName = objSourceTableNamedRange as Name;
                sourceTableRange = sourceTableName.RefersToRange;
                for (int iRow = 1; iRow <= sourceTableRange.Rows.Count; iRow++)
                {
                    for (int iCol = 1; iCol <= sourceTableRange.Columns.Count; iCol++)
                    {
                        destCell = sourceTableRange.Cells[iRow, iCol] as Range;
                        if (destCell != null && destCell.HasFormula)
                        {
                            destCell.Formula = destCell.Formula.Replace(sourceString, destString + i.ToString());
                        }
                    }
                }
            }

        Cleanup_UpdateValidationResultsTableFormulas:
            {
                try
                {
                    WorksheetUtilities.ReleaseComObject(destCell);
                    WorksheetUtilities.ReleaseComObject(sourceTableRange);
                    WorksheetUtilities.ReleaseComObject(sourceTableName);
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

        //HERE - Disso Tables Update
        private static void UpdateDissolutionFormulas(_Worksheet sheet, string namedRangeBaseName, string referenceRange, int baseRow, int baseCol)
        {
            object objNamedRange = null;
            Name name = null;
            Range range = null;
            object objRefNamedRange = null;
            Name refName = null;
            Range refRange = null;

            objNamedRange = sheet.Names.Item(namedRangeBaseName, Type.Missing, Type.Missing);
            if (objNamedRange == null) return;
            objRefNamedRange = sheet.Names.Item(referenceRange, Type.Missing, Type.Missing);
            if (objRefNamedRange == null) return;

            name = objNamedRange as Name;
            range = name.RefersToRange;

            refName = objRefNamedRange as Name;
            refRange = refName.RefersToRange;

            var numRows = range.Rows.Count;
            var numColumns = range.Columns.Count;
            Range refCell;
            string refAddress;

            for (int y = 0; y < numColumns; y++)
            {
                if (y + 1 < 7)
                {
                    refCell = (Range)refRange.Cells[1, y + 1];
                    refAddress = ((Range)refCell.Cells[1, 1]).get_AddressLocal(false, false, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                }
                else
                {
                    refCell = (Range)refRange.Cells[1, 6];
                    refAddress = ((Range)refCell.Cells[1, 1]).get_AddressLocal(false, false, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                }

                for (int x = 0; x < numRows; x++)
                {
                    if (x == 0)
                    {
                        //No change required for the first cell
                    }
                    else
                    {
                        Range cell = range.Cells[x + 1, y + 1] as Range;
                        if (cell != null)
                        {
                            if (cell.HasFormula)
                            {
                                string cellFormula = cell.Formula;
                                //Check if the formula of the cell has the reference column
                                if (!cellFormula.Contains(refAddress))
                                {
                                    if (y + 1 == 4)
                                    {
                                        //Column 4 doesn't use the formula
                                    }
                                    else
                                    {
                                        //Get every cell Address in the formula
                                        MatchCollection matches = Regex.Matches(cellFormula, @"(?<![A-Za-z0-9_])([A-Za-z]{1,3}\$?[0-9]+)");
                                        var toReplace = "";
                                        char toCompare = refAddress[0];

                                        // Add the cell addresses to the list
                                        foreach (Match match in matches)
                                        {
                                            //cellAddresses.Add(match.Value);
                                            if (match.Value.Contains(toCompare.ToString()))
                                            {
                                                if (match.Value != refAddress)
                                                {
                                                    toReplace = match.Value;
                                                }
                                            }
                                        }
                                        if (toReplace != "")
                                        {
                                            cell.Formula = cellFormula.Replace(toReplace, refAddress);
                                        }
                                    }
                                }
                            }

                            WorksheetUtilities.ReleaseComObject(cell);
                        }
                    }
                }

                WorksheetUtilities.ReleaseComObject(refCell);
            }

            // Clean up (if needed)
            WorksheetUtilities.ReleaseComObject(objNamedRange);
            WorksheetUtilities.ReleaseComObject(range);
            WorksheetUtilities.ReleaseComObject(name);
            WorksheetUtilities.ReleaseComObject(objRefNamedRange);
            WorksheetUtilities.ReleaseComObject(refRange);
            WorksheetUtilities.ReleaseComObject(refName);
            // ReSharper disable RedundantAssignment
            sheet = null;
            // ReSharper restore RedundantAssignment
        }

        //END Disso Formulas
    }
}