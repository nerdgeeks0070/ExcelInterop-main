using log4net.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Security.Cryptography;

namespace Spreadsheet.Handler
{
    public static class Robustness
    {
        private static Application _app;
        private const int DefaultNumRows = 2;

        private const string TempDirectoryName = "ABD_TempFiles";

        public static string UpdateRobustnessSheet(
            string sourcePath,
            // General Fields
            string strcmbProtocolType, string strcmbProductType, string strcmbTestType,
            // Assay Level Parameters
            int numNumSamples, int numNosArea, int numNosAssay, int numNosLC, string strCmbDiff, decimal valTxtDiff,
            // Impurity Level Parameters
            int numNoP, int numNoS, int numNumConditionIV, string strCmbNoS,
            string strCmbOperator1, decimal valAC1, decimal valAC2, decimal valAC3, string strCmbAbsoluteRelative1,
            string strTxtoperator1, decimal valacceptancecriteria1, string strCmbOperator2, decimal valAC4, decimal valAC5,
            string strTxtoperator2, decimal valacceptancecriteria3,
            // Water Content Parameters
            int numNumSamplesWC, int numNumConditionWC,
            string strCmbWaterContent1, decimal valAC6, string strCmbWaterContent3, decimal valAC7,
            string strCmbWaterContent2, decimal valAC8, string strCmbWaterContent4, decimal valAC9,
            // Dissolution Parameters
            int numNoSDisso, int numNumConditionDisso, int numNumTimepointsDisso,
            // Dissolution Criteria (Recoveries)
            string strCmbCB1Recoveries, decimal valRecoveriesTB1AccCriteria, string strCmbCB2Recoveries, decimal valRecoveriesTB2AccCriteria, decimal valRecoveriesCBAcceptanceCriteria,
            // Dissolution Criteria (Dissolved)
            string strCmbCB1Dissolved, decimal valDissolvedTB1AccCriteria, string strCmbCB2Dissolved, decimal valDissolvedTB2AccCriteria, decimal valDissolvedCBAcceptanceCriteria
        )
        {
            string returnPath = "";
            try
            {
                returnPath = UpdateRobustnessSheet2(
                                sourcePath,
                                // General Fields
                                strcmbProtocolType, strcmbProductType, strcmbTestType,
                                // Assay Level Parameters
                                numNumSamples, numNosArea, numNosAssay, numNosLC, strCmbDiff, valTxtDiff,
                                // Impurity Level Parameters
                                numNoP, numNoS, numNumConditionIV, strCmbNoS,
                                strCmbOperator1, valAC1, valAC2, valAC3, strCmbAbsoluteRelative1,
                                strTxtoperator1, valacceptancecriteria1, strCmbOperator2, valAC4, valAC5,
                                strTxtoperator2, valacceptancecriteria3,
                                // Water Content Parameters
                                numNumSamplesWC, numNumConditionWC,
                                strCmbWaterContent1, valAC6, strCmbWaterContent3, valAC7,
                                strCmbWaterContent2, valAC8, strCmbWaterContent4, valAC9,
                                // Dissolution Parameters
                                numNoSDisso, numNumConditionDisso, numNumTimepointsDisso,
                                // Dissolution Criteria (Recoveries)
                                strCmbCB1Recoveries, valRecoveriesTB1AccCriteria, strCmbCB2Recoveries, valRecoveriesTB2AccCriteria, valRecoveriesCBAcceptanceCriteria,
                                // Dissolution Criteria (Dissolved)
                                strCmbCB1Dissolved, valDissolvedTB1AccCriteria, strCmbCB2Dissolved, valDissolvedTB2AccCriteria, valDissolvedCBAcceptanceCriteria
                            );
            }
            catch (Exception ex)
            {
                Logger.LogMessage("An error occurred in the call to Robustness.UpdateRobustnessSheet. Message and stack trace are:\r\n" + ex.Message + "\r\n" + ex.StackTrace, Level.Error);

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
                            Logger.LogMessage("An error occurred in the call to Robustness.UpdateRobustnessSheet. Failed to save current workbook changes and to get path.", Level.Error);
                        }

                        _app.Workbooks.Close();
                    }
                    _app = null;
                }
                catch
                {
                    Logger.LogMessage("An error occurred in the call to Robustness.UpdateRobustnessSheet. Application failed to close workbooks. Message and stack trace are:\r\n" + ex.Message + "\r\n" + ex.StackTrace, Level.Error);
                }
                finally
                {
                    WorksheetUtilities.ReleaseExcelApp();
                }
            }

            return returnPath;
        }

        private static string UpdateRobustnessSheet2(
            string sourcePath,
            // General Fields
            string strcmbProtocolType, string strcmbProductType, string strcmbTestType,
            // Assay Level Parameters
            int numNumSamples, int numNosArea, int numNosAssay, int numNosLC, string strCmbDiff, decimal valTxtDiff,
            // Impurity Level Parameters
            int numNoP, int numNoS, int numNumConditionIV, string strCmbNoS,
            string strCmbOperator1, decimal valAC1, decimal valAC2, decimal valAC3, string strCmbAbsoluteRelative1,
            string strTxtoperator1, decimal valacceptancecriteria1, string strCmbOperator2, decimal valAC4, decimal valAC5,
            string strTxtoperator2, decimal valacceptancecriteria3,
            // Water Content Parameters
            int numNumSamplesWC, int numNumConditionWC,
            string strCmbWaterContent1, decimal valAC6, string strCmbWaterContent3, decimal valAC7,
            string strCmbWaterContent2, decimal valAC8, string strCmbWaterContent4, decimal valAC9,
            // Dissolution Parameters
            int numNoSDisso, int numNumConditionDisso, int numNumTimepointsDisso,
            // Dissolution Criteria (Recoveries)
            string strCmbCB1Recoveries, decimal valRecoveriesTB1AccCriteria, string strCmbCB2Recoveries, decimal valRecoveriesTB2AccCriteria, decimal valRecoveriesCBAcceptanceCriteria,
            // Dissolution Criteria (Dissolved)
            string strCmbCB1Dissolved, decimal valDissolvedTB1AccCriteria, string strCmbCB2Dissolved, decimal valDissolvedTB2AccCriteria, decimal valDissolvedCBAcceptanceCriteria
        )
        {
            if (!File.Exists(sourcePath))
            {
                Logger.LogMessage("Error in call to Robustness.UpdateRobustnessSheet. Invalid source file path specified.", Level.Error);
                return "";
            }

            string savePath = WorksheetUtilities.CopyWorkbook(sourcePath, TempDirectoryName, "Robustness Results.xls");
            if (string.IsNullOrEmpty(savePath)) return "";

            _app = WorksheetUtilities.GetExcelApp();
            _app.Workbooks.Open(savePath, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            Workbook book = _app.Workbooks[1];

            Worksheet sheetAssay = book.Worksheets["Assay level"] as Worksheet;
            Worksheet sheetImpurity = book.Worksheets["Impurity"] as Worksheet;
            Worksheet sheetWater = book.Worksheets["Water Content"] as Worksheet;
            Worksheet sheetDissolution = book.Worksheets["Dissolution"] as Worksheet;

            switch (strcmbTestType)
            {
                case "Assay Level":
                    WorksheetUtilities.DeleteSheet(sheetImpurity);
                    WorksheetUtilities.DeleteSheet(sheetWater);
                    WorksheetUtilities.DeleteSheet(sheetDissolution);

                    UpdateAssaySheet(sheetAssay, strcmbProtocolType, strcmbProductType, strcmbTestType,
                        numNumSamples, numNosArea, numNosAssay, numNosLC, strCmbDiff, valTxtDiff);

                    WorksheetUtilities.ReplicateSheetAndDeleteOriginal(book, sheetAssay, numNumSamples);
                    break;

                case "Impurities":
                    WorksheetUtilities.DeleteSheet(sheetAssay);
                    WorksheetUtilities.DeleteSheet(sheetWater);
                    WorksheetUtilities.DeleteSheet(sheetDissolution);

                    UpdateImpuritySheet(sheetImpurity, strcmbProtocolType, strcmbProductType, strcmbTestType,
                        numNoP, numNoS /* number of Solutions or Samples */, numNumConditionIV, strCmbNoS,
                        strCmbOperator1, valAC1, valAC2, valAC3, strCmbAbsoluteRelative1,
                        strTxtoperator1, valacceptancecriteria1, strCmbOperator2, valAC4, valAC5,
                        strTxtoperator2, valacceptancecriteria3);

                    WorksheetUtilities.ReplicateSheetAndDeleteOriginal(book, sheetImpurity, numNoS);
                    break;

                case "Water Content":
                    WorksheetUtilities.DeleteSheet(sheetAssay);
                    WorksheetUtilities.DeleteSheet(sheetImpurity);
                    WorksheetUtilities.DeleteSheet(sheetDissolution);

                    UpdateWaterContentSheet(sheetWater, strcmbProtocolType, strcmbProductType, strcmbTestType,
                        numNumSamplesWC, numNumConditionWC,
                        strCmbWaterContent1, valAC6, strCmbWaterContent3, valAC7,
                        strCmbWaterContent2, valAC8, strCmbWaterContent4, valAC9);
                    //WorksheetUtilities.ReplicateSheetAndDeleteOriginal(book, sheetWater, numNumSamplesWC);
                    break;

                case "Dissolution":
                    WorksheetUtilities.DeleteSheet(sheetAssay);
                    WorksheetUtilities.DeleteSheet(sheetImpurity);
                    WorksheetUtilities.DeleteSheet(sheetWater);

                    UpdateDissolutionSheet(sheetDissolution, strcmbProtocolType, strcmbProductType, strcmbTestType,
                        numNoSDisso, numNumConditionDisso, numNumTimepointsDisso,
                        strCmbCB1Recoveries, valRecoveriesTB1AccCriteria, strCmbCB2Recoveries, valRecoveriesTB2AccCriteria, valRecoveriesCBAcceptanceCriteria,
                        strCmbCB1Dissolved, valDissolvedTB1AccCriteria, strCmbCB2Dissolved, valDissolvedTB2AccCriteria, valDissolvedCBAcceptanceCriteria);

                    WorksheetUtilities.ReplicateSheetAndDeleteOriginal(book, sheetDissolution, numNoSDisso);
                    break;

                case "AssayLevel_Impurities":
                    WorksheetUtilities.DeleteSheet(sheetWater);
                    WorksheetUtilities.DeleteSheet(sheetDissolution);

                    UpdateAssaySheet(sheetAssay, strcmbProtocolType, strcmbProductType, strcmbTestType,
                        numNumSamples, numNosArea, numNosAssay, numNosLC, strCmbDiff, valTxtDiff);

                    WorksheetUtilities.ReplicateSheetAndDeleteOriginal(book, sheetAssay, numNumSamples);

                    UpdateImpuritySheet(sheetImpurity, strcmbProtocolType, strcmbProductType, strcmbTestType,
                        numNoP, numNoS, numNumConditionIV, strCmbNoS,
                        strCmbOperator1, valAC1, valAC2, valAC3, strCmbAbsoluteRelative1,
                        strTxtoperator1, valacceptancecriteria1, strCmbOperator2, valAC4, valAC5,
                        strTxtoperator2, valacceptancecriteria3);

                    WorksheetUtilities.ReplicateSheetAndDeleteOriginal(book, sheetImpurity, numNoS);
                    break;

                default:
                    Logger.LogMessage($"Unknown test type: {strcmbTestType}", Level.Error);
                    break;
            }

            book.Save();
            WorksheetUtilities.ReleaseComObject(book);
            _app.Workbooks.Close();
            _app = null;
            WorksheetUtilities.ReleaseExcelApp();

            return savePath;
        }

        private static void UpdateAssaySheet(
            Worksheet sheet, string strcmbProtocolType, string strcmbProductType, string strcmbTestType,
            int numSamples, int numNosArea, int numNosAssay, int numNosLC, string strCmbDiff, decimal valTxtDiff)
        {
            if (sheet == null)
                return;

            bool wasProtected = WorksheetUtilities.SetSheetProtection(sheet, null, false);

            WorksheetUtilities.SetMetadataValues(sheet, strcmbProtocolType, strcmbProductType, strcmbTestType);

            WorksheetUtilities.SetNamedRangeValue(sheet, "AssayLevelOperator", strCmbDiff, 1, 1);
            WorksheetUtilities.SetNamedRangeValue(sheet, "AssayLevelValue", valTxtDiff.ToString(), 1, 1);

            // we pass numSamples as 1, as we want to have separate sheet for each sample.
            string sampleInfoNamedRange = WorksheetUtilities.ProcessSampleInfo(sheet, strcmbProductType, 1 /*numSamples*/);

            // --- Handle each section
            ProcessAssaySection(sheet, "Area", numNosArea, sampleInfoNamedRange);
            ProcessAssaySection(sheet, "Assay", numNosAssay, sampleInfoNamedRange);
            ProcessAssaySection(sheet, "Claim", numNosLC, sampleInfoNamedRange);

            // remaining code here...

            WorksheetUtilities.DeleteInvalidNamedRanges(sheet);
            WorksheetUtilities.ScrollToTopLeft(sheet);
            if (wasProtected) WorksheetUtilities.SetSheetProtection(sheet, null, true);
        }

        /// <summary>
        /// Handles a named range section: deletes if count=0, otherwise adjusts rows and replicates sets as needed.
        /// </summary>
        private static void ProcessAssaySection(Worksheet sheet, string sectionName, int count, string sampleInfoNamedRange)
        {
            string dataRangeName = $"{sectionName}Data";
            string rowRangeName = $"{sectionName}DataRows1";
            string baseSetName = $"{sectionName}Set1";
            string sampleSummaryRef = $"{sectionName}SampleSummaryRef";

            if (count == 0)
            {
                WorksheetUtilities.DeleteNamedRangeRows(sheet, dataRangeName);
                return; // skip further processing
            }

            SetSampleSummaryRef(sheet, sampleInfoNamedRange, sampleSummaryRef);

            if (count > DefaultNumRows)
                WorksheetUtilities.InsertRowsIntoNamedRange(count - DefaultNumRows, sheet, rowRangeName, true, XlDirection.xlDown, XlPasteType.xlPasteAll);
            else if (count < DefaultNumRows)
                WorksheetUtilities.DeleteRowsFromNamedRange(1, sheet, rowRangeName, XlDirection.xlDown);

            // update: we want to have separate sheet for each, so no duplication in same sheet is needed.
            // DuplicateSetRanges(sheet, baseSetName, numSamples);
        }


        private static void DuplicateSetRanges(Worksheet sheet, string baseSetName, int numSamples)
        {
            if (numSamples <= 1)
                return;

            Name baseNamedRange = sheet.Names.Item(baseSetName);
            Range lastRange = baseNamedRange.RefersToRange;
            int rows = lastRange.Rows.Count;
            int cols = lastRange.Columns.Count;

            for (int i = 2; i <= numSamples; i++)
            {
                // Calculate where to insert next block
                Range insertPoint = lastRange.get_Offset(rows, 0);

                // Copy and insert, shifting rows down
                lastRange.Copy();
                insertPoint.Insert(XlInsertShiftDirection.xlShiftDown);

                // Recalculate the new top-left cell after insertion
                Range newTopLeft = lastRange.get_Offset(rows, 0);
                Range newRange = newTopLeft.get_Resize(rows, cols);

                // Create new named range
                string newRangeName = baseSetName.Replace("1", i.ToString());
                sheet.Names.Add(Name: newRangeName, RefersTo: $"={newRange.Address}");

                // Update lastRange to the newly inserted block
                lastRange = newRange;
            }
        }

        private static void UpdateImpuritySheet(
            Worksheet sheet, string strcmbProtocolType, string strcmbProductType, string strcmbTestType,
            int numNoP, int numNoS, int numNumConditionIV, string strCmbNoS,
            string strCmbOperator1, decimal valAC1, decimal valAC2, decimal valAC3, string strCmbAbsoluteRelative1,
            string strTxtoperator1, decimal valacceptancecriteria1, string strCmbOperator2, decimal valAC4, decimal valAC5,
            string strTxtoperator2, decimal valacceptancecriteria3)
        {
            if (sheet == null)
                return;

            bool wasProtected = WorksheetUtilities.SetSheetProtection(sheet, null, false);

            WorksheetUtilities.SetMetadataValues(sheet, strcmbProtocolType, strcmbProductType, strcmbTestType);

            WorksheetUtilities.SetNamedRangeValue(sheet, "ImpurityRangeOperator1", strCmbOperator1, 1, 1);
            WorksheetUtilities.SetNamedRangeValue(sheet, "ImpurityRangeOperator2", strCmbOperator2, 1, 1);
            WorksheetUtilities.SetNamedRangeValue(sheet, "ImpurityRangeValue1", valAC1.ToString(), 1, 1);
            WorksheetUtilities.SetNamedRangeValue(sheet, "ImpurityRangeValue2", valAC3.ToString(), 1, 1);

            WorksheetUtilities.SetNamedRangeValue(sheet, "ImpurityDiffValue1", valAC2.ToString(), 1, 1);
            WorksheetUtilities.SetNamedRangeValue(sheet, "ImpurityDiffValue2", valAC4.ToString(), 1, 1);
            WorksheetUtilities.SetNamedRangeValue(sheet, "ImpurityDiffValue3", valAC5.ToString(), 1, 1);
            WorksheetUtilities.SetNamedRangeValue(sheet, "ImpurityDiffRelAbs", strCmbAbsoluteRelative1, 1, 1);

            WorksheetUtilities.SetNamedRangeValue(sheet, "ImpurityQuantitationType", strCmbNoS, 1, 1);

            string sampleInfoNamedRange = WorksheetUtilities.ProcessSampleInfo(sheet, strcmbProductType, 1 /*numSamples*/);

            ProcessImpurityTables(sheet, numNoP, numNumConditionIV, sampleInfoNamedRange);

            // remaining code here...

            WorksheetUtilities.DeleteInvalidNamedRanges(sheet);
            WorksheetUtilities.ScrollToTopLeft(sheet);
            if (wasProtected) WorksheetUtilities.SetSheetProtection(sheet, null, true);
        }

        private static void ProcessImpurityTables(Worksheet sheet, int numNoP, int numNumConditionIV, string sampleInfoNamedRange)
        {
            SetSampleSummaryRef(sheet, sampleInfoNamedRange, "ImpuritySampleSummaryRef");

            string prevRangeName = "Impurity1";
            string prevSummaryRangeName = "ImpuritySummaryColumn1";
            int colOffsetRawTable = 3 + 1;
            int colOffsetSummaryTable = 1 + 1;

            for (int i = 2; i <= numNoP; i++)
            {
                string impurityName = $"Impurity{i}";
                string summaryName = $"ImpuritySummaryColumn{i}";

                // Copy ranges from previous
                // raw table ranges
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, prevRangeName, impurityName, 1, colOffsetRawTable, XlPasteType.xlPasteAll);

                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"ImpurityDifferenceConditions1{i - 1}", $"ImpurityDifferenceConditions1{i}", 1, colOffsetRawTable, XlPasteType.xlPasteAll);
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"ImpurityInitialEmpower1{i - 1}", $"ImpurityInitialEmpower1{i}", 1, colOffsetRawTable, XlPasteType.xlPasteAll);
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"ImpurityDifference1{i - 1}", $"ImpurityDifference1{i}", 1, colOffsetRawTable, XlPasteType.xlPasteAll);

                // summary table ranges
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, prevSummaryRangeName, summaryName, 1, colOffsetSummaryTable, XlPasteType.xlPasteAll);

                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"Decimals{i - 1}", $"Decimals{i}", 1, colOffsetSummaryTable, XlPasteType.xlPasteAll);

                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"ImpuritySummaryConditions1{i - 1}", $"ImpuritySummaryConditions1{i}", 1, colOffsetSummaryTable, XlPasteType.xlPasteAll);
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"ImpuritySummaryInitialPct1{i - 1}", $"ImpuritySummaryInitialPct1{i}", 1, colOffsetSummaryTable, XlPasteType.xlPasteAll);
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"ImpuritySummaryDifference1{i - 1}", $"ImpuritySummaryDifference1{i}", 1, colOffsetSummaryTable, XlPasteType.xlPasteAll);

                // Set cell (1,1) of Impurity{i} to "Impurity {i}" for Impurity Name
                WorksheetUtilities.SetNamedRangeValue(sheet, impurityName, $"Impurity {i}", 1, 1);

                // Set cell (3,1) of ImpuritySummaryColumn{i} to formula referencing Impurity{i} cell (1,1)
                string sourceAddress = WorksheetUtilities.GetCellAddress(sheet, $"Impurity{i}", 1, 1);
                string formula = WorksheetUtilities.GetSimpleReferenceFormula(sourceAddress);
                WorksheetUtilities.SetNamedRangeFormula(sheet, summaryName, formula, 3, 1);

                // Link each cell in ImpuritySummaryConditions1{i} to corresponding cell in ImpurityDifferenceConditions1{i}
                string decimalsUpperCellAbsoluteAddress = WorksheetUtilities.GetCellAddress(sheet, $"Decimals{i}", 1, 1, true, true);
                string decimalsLowerCellAbsoluteAddress = WorksheetUtilities.GetCellAddress(sheet, $"Decimals{i}", 2, 1, true, true);

                // set ImpurityInitialEmpower1 linking
                string sourceInitialEmpowerAddress = WorksheetUtilities.GetCellAddress(sheet, $"ImpurityInitialEmpower1{i}", 1, 1);
                string linkInitialEmpowerConditionsFormula = WorksheetUtilities.GetDifferenceConditionsLinkingFormula(sourceInitialEmpowerAddress, decimalsUpperCellAbsoluteAddress);
                WorksheetUtilities.SetNamedRangeFormula(sheet, $"ImpuritySummaryInitialPct1{i}", linkInitialEmpowerConditionsFormula, 1, 1);

                // set ImpurityDifference1 linking
                string sourceInitialDiffAddress = WorksheetUtilities.GetCellAddress(sheet, $"ImpurityDifference1{i}", 1, 1);
                string linkInitialDiffConditionsFormula = WorksheetUtilities.GetDifferenceConditionsLinkingFormula(sourceInitialDiffAddress, decimalsLowerCellAbsoluteAddress);
                WorksheetUtilities.SetNamedRangeFormula(sheet, $"ImpuritySummaryDifference1{i}", linkInitialDiffConditionsFormula, 1, 1);

                Range diffRange = sheet.Range[$"ImpurityDifferenceConditions1{i}"];

                int rowCount = diffRange.Rows.Count;
                for (int r = 1; r <= rowCount; r++)
                {
                    // set difference condition linking
                    string srcCellAddress = WorksheetUtilities.GetCellAddress(sheet, $"ImpurityDifferenceConditions1{i}", r, 1);
                    string linkDiffConditionsFormula = WorksheetUtilities.GetDifferenceConditionsLinkingFormula(srcCellAddress, decimalsLowerCellAbsoluteAddress);
                    WorksheetUtilities.SetNamedRangeFormula(sheet, $"ImpuritySummaryConditions1{i}", linkDiffConditionsFormula, r, 1);
                }

                // Update previous range names
                prevRangeName = impurityName;
                prevSummaryRangeName = summaryName;
            }

            // Adjust RawConditionRows and SummaryConditionRows
            if (numNumConditionIV > DefaultNumRows)
            {
                int rowsToInsert = numNumConditionIV - DefaultNumRows;
                WorksheetUtilities.InsertRowsIntoNamedRange(rowsToInsert, sheet, "RawConditionRows", true, XlDirection.xlDown, XlPasteType.xlPasteAll);
                WorksheetUtilities.InsertRowsIntoNamedRange(rowsToInsert, sheet, "SummaryConditionRows", true, XlDirection.xlDown, XlPasteType.xlPasteAll);
            }
            else if (numNumConditionIV < DefaultNumRows)
            {
                WorksheetUtilities.DeleteRowsFromNamedRange(1, sheet, "RawConditionRows", XlDirection.xlDown);
                WorksheetUtilities.DeleteRowsFromNamedRange(1, sheet, "SummaryConditionRows", XlDirection.xlDown);
            }
        }

        private static void UpdateWaterContentSheet(
            Worksheet sheet, string strcmbProtocolType, string strcmbProductType, string strcmbTestType,
            int numNumSamplesWC, int numNumConditionWC,
            string strCmbWaterContent1, decimal valAC6, string strCmbWaterContent3, decimal valAC7,
            string strCmbWaterContent2, decimal valAC8, string strCmbWaterContent4, decimal valAC9)
        {
            // Implement water content update logic here
            if (sheet == null)
            {
                return;
            }

            bool wasProtected = WorksheetUtilities.SetSheetProtection(sheet, null, false);

            if (!string.IsNullOrEmpty(strCmbWaterContent1))
            {
                sheet.Cells[5, 2] = strCmbWaterContent1;
                sheet.Cells[5, 3] = valAC6;
            }
            if (!string.IsNullOrEmpty(strCmbWaterContent3))
            {
                sheet.Cells[5, 4] = strCmbWaterContent3;
                sheet.Cells[5, 5] = valAC7;
            }
            if (!string.IsNullOrEmpty(strCmbWaterContent2))
            {
                sheet.Cells[7, 2] = strCmbWaterContent3;
                sheet.Cells[7, 3] = valAC8;
            }
            if (!string.IsNullOrEmpty(strCmbWaterContent4))
            {
                sheet.Cells[7, 4] = strCmbWaterContent4;
                sheet.Cells[7, 5] = valAC9;
            }

            if (!string.IsNullOrEmpty(strcmbProtocolType))
            {
                sheet.Cells[5, 9] = strcmbProtocolType;
            }

            if (!string.IsNullOrEmpty(strcmbTestType))
            {
                sheet.Cells[7, 9] = strcmbTestType;
            }

            if (!string.IsNullOrEmpty(strcmbProductType))
            {
                sheet.Cells[6, 9] = strcmbProductType;

                if(strcmbProductType == "Drug Substance")
                {
                    WorksheetUtilities.DeleteNamedRangeRows(sheet, "Drug_Product");
                    WorksheetUtilities.DeleteNamedRangeRows(sheet, "SDD");
                }

                if (strcmbProductType == "Drug Product")
                {
                    WorksheetUtilities.DeleteNamedRangeRows(sheet, "Drug_Substance");
                    WorksheetUtilities.DeleteNamedRangeRows(sheet, "SDD");
                }

                if (strcmbProductType == "SDD")
                {
                    WorksheetUtilities.DeleteNamedRangeRows(sheet, "Drug_Product");
                    WorksheetUtilities.DeleteNamedRangeRows(sheet, "Drug_Substance");
                }
            }

            if(numNumConditionWC > 1)
            {
                var rowsToInsert = numNumConditionWC - 1;
                WorksheetUtilities.InsertRowsIntoNamedRange(rowsToInsert, sheet, "Summary_Table", false, XlDirection.xlDown, XlPasteType.xlPasteAll);
                WorksheetUtilities.ResizeNamedRange(sheet, "Summary_Table", rowsToInsert, 0);
            }

            WorksheetUtilities.DeleteInvalidNamedRanges(sheet);
            WorksheetUtilities.ScrollToTopLeft(sheet);
            if (wasProtected) WorksheetUtilities.SetSheetProtection(sheet, null, true);
        }

        private static void UpdateDissolutionSheet(
            Worksheet sheet, string strcmbProtocolType, string strcmbProductType, string strcmbTestType,
            // Dissolution Parameters
            int numNoSDisso, int numNumConditionDisso, int numNumTimepointsDisso,
            // Dissolution Criteria (Recoveries)
            string strCmbCB1Recoveries, decimal valRecoveriesTB1AccCriteria, string strCmbCB2Recoveries, decimal valRecoveriesTB2AccCriteria, decimal valRecoveriesCBAcceptanceCriteria,
            // Dissolution Criteria (Dissolved)
            string strCmbCB1Dissolved, decimal valDissolvedTB1AccCriteria, string strCmbCB2Dissolved, decimal valDissolvedTB2AccCriteria, decimal valDissolvedCBAcceptanceCriteria)
        {
            if (sheet == null)
            {
                return;
            }

            bool wasProtected = WorksheetUtilities.SetSheetProtection(sheet, null, false);
            var defaultNumTimePoints = 5;

            //WorksheetUtilities.SetMetadataValues(sheet, strcmbProtocolType, strcmbProductType, strcmbTestType);

            if (!string.IsNullOrEmpty(strCmbCB1Recoveries))
            {
                sheet.Cells[5, 2] = strCmbCB1Recoveries;
                sheet.Cells[5, 3] = valRecoveriesTB1AccCriteria;
            }
            if (!string.IsNullOrEmpty(strCmbCB2Recoveries))
            {
                sheet.Cells[5, 6] = strCmbCB2Recoveries;
                sheet.Cells[5, 7] = valRecoveriesTB2AccCriteria;
            }
            if (!string.IsNullOrEmpty(strCmbCB1Dissolved))
            {
                sheet.Cells[6, 2] = strCmbCB1Dissolved;
                sheet.Cells[6, 3] = valDissolvedTB1AccCriteria;
            }
            if (!string.IsNullOrEmpty(strCmbCB2Dissolved))
            {
                sheet.Cells[6, 6] = strCmbCB2Dissolved;
                sheet.Cells[6, 7] = valDissolvedTB2AccCriteria;
            }

            // remaining code here
            if (numNumTimepointsDisso > defaultNumTimePoints)
            {
                var rowsToInsert = numNumTimepointsDisso - defaultNumTimePoints;

                AppendRowsCopyOnlyFormulasAndRenumber(sheet, "DissolutionMethodTimePoints", rowsToInsert, 0, 1, 1);
                AppendRowsCopyOnlyFormulasAndRenumber(sheet, "DissolutionConditionTable1", rowsToInsert, 2, 1, 1);
                AppendRowsCopyOnlyFormulasAndRenumber(sheet, "DissolutionSummaryTimePoints1", rowsToInsert, 0, 1, 1);
            }
            else if (numNumTimepointsDisso < defaultNumTimePoints)
            {
                var rowsToDelete = defaultNumTimePoints - numNumTimepointsDisso;

                WorksheetUtilities.DeleteRowsFromNamedRange(rowsToDelete, sheet, "DissolutionMethodTimePoints", XlDirection.xlUp);
                WorksheetUtilities.DeleteRowsFromNamedRange(rowsToDelete, sheet, "DissolutionConditionTable1", XlDirection.xlUp);
                WorksheetUtilities.DeleteRowsFromNamedRange(rowsToDelete, sheet, "DissolutionSummaryTimePoints1", XlDirection.xlUp);
            }

            SetBorderToNamedRange(sheet, "DissolutionMethodTimePoints");
            SetBorderToNamedRange(sheet, "DissolutionConditionTable1");
            SetBorderToNamedRange(sheet, "DissolutionSummaryTimePoints1");

            if (numNumConditionDisso > 1)
            {
                var rowsToInsert = numNumConditionDisso - 1;

                AppendRowsCopyOnlyFormulasAndRenumber(sheet, "DissolutionVariablesInfo", rowsToInsert, 0, 1, 1);


                for (int i = 2; i <= numNumConditionDisso; i++)
                {
                    WorksheetUtilities.InsertRowsIntoNamedRange(3 + numNumTimepointsDisso, sheet, "DissolutionCondition", false, XlDirection.xlUp, XlPasteType.xlPasteAll);

                    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "DissolutionConditionTable" + (i - 1), "DissolutionConditionTable" + i, 4 + numNumTimepointsDisso, 1, XlPasteType.xlPasteAll);

                    WorksheetUtilities.InsertRowsIntoNamedRange(2 + numNumTimepointsDisso, sheet, "DissolutionSummary", false, XlDirection.xlUp, XlPasteType.xlPasteAll);

                    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "DissolutionSummaryTimePoints" + (i - 1), "DissolutionSummaryTimePoints" + i, 3 + numNumTimepointsDisso, 1, XlPasteType.xlPasteAll);

                    SetDissolutionConditionHeaderFormulas(sheet, i);

                    LinkNameRanges(sheet, "DissolutionSummaryTimePoints" + i, "DissolutionMethodTimePoints", "H", "I");
                    LinkNameRanges(sheet, "DissolutionSummaryTimePoints" + i, "DissolutionConditionTable" + i, "I", "I");
                    LinkNameRanges(sheet, "DissolutionSummaryTimePoints" + i, "DissolutionMethodTimePoints", "B", "B");
                }
            }

            WorksheetUtilities.DeleteInvalidNamedRanges(sheet);
            WorksheetUtilities.ScrollToTopLeft(sheet);
            if (wasProtected) WorksheetUtilities.SetSheetProtection(sheet, null, true);
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

                int originalTotalRows = totalRows;
                int originalDataRowCount = originalTotalRows - headerRows;

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

                int firstNewDataRowWithinRange = headerRows + originalDataRowCount + 1;

                for (int r = 0; r < rowsToAdd; r++)
                {
                    int rowWithinRange = firstNewDataRowWithinRange + r;

                    templateRow.Copy(Type.Missing);
                    Range destRow = rng.Rows[rowWithinRange, Type.Missing] as Range;
                    destRow.PasteSpecial(XlPasteType.xlPasteFormats,
                                         XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                    Border newRowTopBorder = destRow.Borders[XlBordersIndex.xlEdgeTop];
                    newRowTopBorder.LineStyle = XlLineStyle.xlLineStyleNone;

                    for (int c = 1; c <= totalCols; c++)
                    {
                        if (!string.IsNullOrEmpty(templateFormulas[c - 1]))
                        {
                            Range destCell = destRow.Cells[1, c] as Range;
                            destCell.FormulaR1C1 = templateFormulas[c - 1];
                        }
                    }
                }

                if (originalDataRowCount > 0)
                {
                    int originalLastRowIndex = headerRows + originalDataRowCount;
                    Range originalLastRow = rng.Rows[originalLastRowIndex, Type.Missing] as Range;
                    if (originalLastRow != null)
                    {
                        Border origBottom = originalLastRow.Borders[XlBordersIndex.xlEdgeBottom];
                        origBottom.LineStyle = XlLineStyle.xlLineStyleNone;
                    }
                }

                //Range lastRow = rng.Rows[totalRows, Type.Missing] as Range;
                //if (lastRow != null)
                //{
                //    Border templateBottom = templateRow.Borders[XlBordersIndex.xlEdgeBottom];
                //    Border lastBottom = lastRow.Borders[XlBordersIndex.xlEdgeBottom];

                //    lastBottom.LineStyle = templateBottom.LineStyle;
                //    lastBottom.Weight = templateBottom.Weight;
                //    lastBottom.Color = templateBottom.Color;

                //}

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

        private static void LinkNameRanges(
            Worksheet sheet,
            string summaryNamedRangeName,
            string srcNamedRangeName,
            string destColumn,
            string srcColumn)
        {
            Range methodRange = null;
            Range summaryRange = null;

            try
            {
                methodRange = sheet.Range[srcNamedRangeName];
                summaryRange = sheet.Range[summaryNamedRangeName];

                if (methodRange == null || summaryRange == null)
                    return;

                int methodRowCount = methodRange.Rows.Count;
                int summaryRowCount = summaryRange.Rows.Count;

                int headerRows = methodRowCount - summaryRowCount;

                if (summaryRowCount <= 0)
                    return;

                for (int i = 0; i < summaryRowCount; i++)
                {
                    int srcRow = methodRange.Row + headerRows + i;
                    int dstRow = summaryRange.Row + i;

                    Range srcCell = sheet.Cells[srcRow, srcColumn] as Range;
                    Range dstCell = sheet.Cells[dstRow, destColumn] as Range;

                    if (srcCell != null && dstCell != null)
                    {
                        string srcAddress = srcCell.get_Address(false, false, XlReferenceStyle.xlA1);
                        dstCell.Formula = $"={srcAddress}";
                    }

                    WorksheetUtilities.ReleaseComObject(srcCell);
                    WorksheetUtilities.ReleaseComObject(dstCell);
                }
            }
            finally
            {
                WorksheetUtilities.ReleaseComObject(methodRange);
                WorksheetUtilities.ReleaseComObject(summaryRange);
            }
        }

        private static void SetDissolutionConditionHeaderFormulas(Worksheet sheet, int tableIndex)
        {
            string tableName = $"DissolutionConditionTable{tableIndex}";

            Range range = WorksheetUtilities.GetNamedRange(sheet, tableName);
            if (range == null) return;

            Range firstRow = range.Rows[1];

            int baseRow = 13;

            int targetRow = baseRow + (tableIndex - 1);

            firstRow.Cells[1, 2].Formula = $"=F{targetRow}";
            firstRow.Cells[1, 3].Formula = $"=D{targetRow}";
            //firstRow.Cells[1, 4].Formula = $"=E{targetRow}";

            Range eCell = firstRow.Cells[1, 4] as Range;

            eCell.ClearContents();
            eCell.NumberFormat = "General";
            eCell.Formula = $"=E{targetRow}";

            WorksheetUtilities.ReleaseComObject(range);
            WorksheetUtilities.ReleaseComObject(firstRow);
        }
        private static void SetBorderToNamedRange(Worksheet sheet, string sampleNamedRange)
        {
            var range = sheet.Names.Item(sampleNamedRange).RefersToRange;
            range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            range.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;
        }


        private static void SetSampleSummaryRef(Worksheet sheet, string sampleInfoNamedRange, string sampleSummaryRef)
        {
            Range targetRange = sheet.Range[sampleInfoNamedRange];
            int col = targetRange.Columns.Count;
            string sourceAddress = WorksheetUtilities.GetCellAddress(sheet, sampleInfoNamedRange, 1, col);
            string formula = WorksheetUtilities.GetSimpleReferenceFormula(sourceAddress);
            WorksheetUtilities.SetNamedRangeFormula(sheet, sampleSummaryRef, formula, 1, 1);
        }
    }
}
