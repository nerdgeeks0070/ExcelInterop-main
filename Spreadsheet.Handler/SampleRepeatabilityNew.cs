using log4net.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Security.Cryptography;

namespace Spreadsheet.Handler
{
    public static class SampleRepeatabilityNew
    {
        private static Application _app;

        private const string TempDirectoryName = "ABD_TempFiles";

        public static string UpdateSampleRepeatabilitySheet(
            string sourcePath,

            // --- General ---
            int numRepsGeneral, string strcmbProtocolType, string strcmbProductType, string strcmbTestType,
            // --- Assay ---
            int numSamplesAssay, string strcmbRSD, decimal valRSD1,
            // --- Content Uniformity (CU) ---
            int numSamplesCU,
            // --- Impurity ---
            int numSamplesImp, string strcmbNoS, int numPeaksImp, string strcmbOperator1Imp, decimal valAC1Imp, decimal valAC2Imp,
            string strAutoOperator1Imp, string strAutoOperator2Imp, decimal valacceptancecriteria1Imp, decimal valacceptancecriteria3Imp,
            string strcmbOperator2Imp, decimal valAC3Imp, decimal valAC4Imp, decimal valAC5Imp, string strcmbOperator4Imp, string strcmbOperator5Imp,
            // --- Water Content ---
            int numSamplesWC, string strcmbWaterContent1WC, decimal valAC6WC, string strcmbWaterContent3WC, decimal valAC7WC,
            string strcmbWaterContent2WC, decimal valAC8WC, string strcmbWaterContent4WC, decimal valAC9WC,
            // --- Dissolution ---
            int numSamplesDisso, int numRepsDisso, string strcmbOperator1Disso, decimal valAC1Disso,
            string strcmbOperator2Disso, decimal valAC2Disso, string strcmbOperator3Disso,
            string strcmbOperator4Disso, decimal valAC3Disso, string strcmbOperator5Disso,
            decimal valAC4Disso, string strcmbOperator6Disso)
        {
            string returnPath = "";

            try
            {
                returnPath = UpdateSampleRepeatabilitySheet2(
                    sourcePath,
                    // --- General ---
                    numRepsGeneral, strcmbProtocolType, strcmbProductType, strcmbTestType,
                    // --- Assay ---
                    numSamplesAssay, strcmbRSD, valRSD1,
                    // --- Content Uniformity ---
                    numSamplesCU,
                    // --- Impurity ---
                    numSamplesImp, strcmbNoS, numPeaksImp, strcmbOperator1Imp, valAC1Imp, valAC2Imp,
                    strAutoOperator1Imp, strAutoOperator2Imp, valacceptancecriteria1Imp, valacceptancecriteria3Imp,
                    strcmbOperator2Imp, valAC3Imp, valAC4Imp, valAC5Imp, strcmbOperator4Imp, strcmbOperator5Imp,
                    // --- Water Content ---
                    numSamplesWC, strcmbWaterContent1WC, valAC6WC, strcmbWaterContent3WC, valAC7WC,
                    strcmbWaterContent2WC, valAC8WC, strcmbWaterContent4WC, valAC9WC,
                    // --- Dissolution ---
                    numSamplesDisso, numRepsDisso, strcmbOperator1Disso, valAC1Disso,
                    strcmbOperator2Disso, valAC2Disso, strcmbOperator3Disso,
                    strcmbOperator4Disso, valAC3Disso, strcmbOperator5Disso,
                    valAC4Disso, strcmbOperator6Disso
                );
            }
            catch (Exception ex)
            {
                Logger.LogMessage("An error occurred in the call to SampleRepeatabilityNew.UpdateSampleRepeatabilitySheet. " +
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

        private static string UpdateSampleRepeatabilitySheet2(
            string sourcePath,
            // --- General ---
            int numRepsGeneral, string strcmbProtocolType, string strcmbProductType, string strcmbTestType,
            // --- Assay ---
            int numSamplesAssay, string strcmbRSD, decimal valRSD1,
            // --- Content Uniformity ---
            int numSamplesCU,
            // --- Impurity ---
            int numSamplesImp, string strcmbNoS, int numPeaksImp, string strcmbOperator1Imp, decimal valAC1Imp, decimal valAC2Imp,
            string strAutoOperator1Imp, string strAutoOperator2Imp, decimal valacceptancecriteria1Imp, decimal valacceptancecriteria3Imp,
            string strcmbOperator2Imp, decimal valAC3Imp, decimal valAC4Imp, decimal valAC5Imp,
            string strcmbOperator4Imp, string strcmbOperator5Imp,
            // --- Water Content ---
            int numSamplesWC, string strcmbWaterContent1WC, decimal valAC6WC, string strcmbWaterContent3WC, decimal valAC7WC,
            string strcmbWaterContent2WC, decimal valAC8WC, string strcmbWaterContent4WC, decimal valAC9WC,
            // --- Dissolution ---
            int numSamplesDisso, int numRepsDisso, string strcmbOperator1Disso, decimal valAC1Disso,
            string strcmbOperator2Disso, decimal valAC2Disso, string strcmbOperator3Disso,
            string strcmbOperator4Disso, decimal valAC3Disso, string strcmbOperator5Disso,
            decimal valAC4Disso, string strcmbOperator6Disso
        )
        {
            if (!File.Exists(sourcePath))
            {
                Logger.LogMessage("Error in call to SampleRepeatabilityNew.UpdateSampleRepeatabilitySheet2. Invalid source file path specified.", Level.Error);
                return "";
            }

            string savePath = WorksheetUtilities.CopyWorkbook(sourcePath, TempDirectoryName, "Sample Repeatability Results.xls");
            if (string.IsNullOrEmpty(savePath)) return "";

            _app = WorksheetUtilities.GetExcelApp();
            _app.Workbooks.Open(savePath, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            Workbook book = _app.Workbooks[1];

            Worksheet sheetAssayAPI = book.Worksheets["Assay API"] as Worksheet;
            Worksheet sheetAssayDP = book.Worksheets["Assay DP"] as Worksheet;
            Worksheet sheetImpurity = book.Worksheets["Impurity"] as Worksheet;
            Worksheet sheetContentUniformity = book.Worksheets["Content Uniformity"] as Worksheet;
            Worksheet sheetDissolution = book.Worksheets["Dissolution"] as Worksheet;
            Worksheet sheetWater = book.Worksheets["Water Content"] as Worksheet;

            if (strcmbProductType == "Drug Substance")
            {
                WorksheetUtilities.DeleteSheet(sheetAssayDP);
                WorksheetUtilities.DeleteSheet(sheetContentUniformity);
                WorksheetUtilities.DeleteSheet(sheetDissolution);

                if (strcmbTestType == "Water Content")
                {
                    WorksheetUtilities.DeleteSheet(sheetAssayAPI);
                    WorksheetUtilities.DeleteSheet(sheetImpurity);

                    // call WC method
                }
                else
                {
                    WorksheetUtilities.DeleteSheet(sheetWater);

                    if (numSamplesAssay == 0)
                    {
                        WorksheetUtilities.DeleteSheet(sheetAssayAPI);
                    }
                    else
                    {
                        UpdateAssayAPI(sheetAssayAPI, strcmbProtocolType, strcmbProductType, strcmbTestType, numRepsGeneral, numSamplesAssay, strcmbRSD, valRSD1);
                    }

                    if (numSamplesImp == 0)
                    {
                        WorksheetUtilities.DeleteSheet(sheetImpurity);
                    }
                    else
                    {
                        UpdateImpurity(sheetImpurity, strcmbProtocolType, strcmbProductType, strcmbTestType, numRepsGeneral,
                            numSamplesImp, strcmbNoS, numPeaksImp,
                            strcmbOperator1Imp, valAC1Imp, valAC2Imp, strAutoOperator1Imp, strAutoOperator2Imp, valacceptancecriteria1Imp, valacceptancecriteria3Imp,
                            strcmbOperator2Imp, valAC3Imp, valAC4Imp, valAC5Imp, strcmbOperator4Imp, strcmbOperator5Imp);
                    }
                }
            }
            else // SDD and Drug Product
            {
                WorksheetUtilities.DeleteSheet(sheetAssayAPI);

                if (strcmbTestType == "Water Content")
                {
                    WorksheetUtilities.DeleteSheet(sheetAssayDP);
                    WorksheetUtilities.DeleteSheet(sheetImpurity);
                    WorksheetUtilities.DeleteSheet(sheetContentUniformity);
                    WorksheetUtilities.DeleteSheet(sheetDissolution);

                    UpdateWaterContent(sheetWater,
                        strcmbProtocolType, strcmbProductType, strcmbTestType, numRepsGeneral,
                        numSamplesWC, 
                        strcmbWaterContent1WC, valAC6WC, strcmbWaterContent3WC, valAC7WC,
                        strcmbWaterContent2WC, valAC8WC, strcmbWaterContent4WC, valAC9WC);
                }
                else if (strcmbTestType == "Dissolution")
                {
                    WorksheetUtilities.DeleteSheet(sheetAssayDP);
                    WorksheetUtilities.DeleteSheet(sheetImpurity);
                    WorksheetUtilities.DeleteSheet(sheetContentUniformity);
                    WorksheetUtilities.DeleteSheet(sheetWater);

                    UpdateDissolution(sheetDissolution, 
                        strcmbProtocolType, strcmbProductType, strcmbTestType, numRepsGeneral,
                        numSamplesDisso, numRepsDisso, 
                        strcmbOperator1Disso, valAC1Disso, strcmbOperator2Disso, valAC2Disso, strcmbOperator3Disso, strcmbOperator4Disso, valAC3Disso, 
                        strcmbOperator5Disso, valAC4Disso, strcmbOperator6Disso);
                }
                else
                {
                    WorksheetUtilities.DeleteSheet(sheetWater);
                    WorksheetUtilities.DeleteSheet(sheetDissolution);

                    if (numSamplesAssay == 0)
                    {
                        WorksheetUtilities.DeleteSheet(sheetAssayDP);
                    }
                    else
                    {
                        UpdateAssayDP(sheetAssayDP, strcmbProtocolType, strcmbProductType, strcmbTestType, numRepsGeneral, numSamplesAssay, strcmbRSD, valRSD1);
                    }

                    if (numSamplesImp == 0)
                    {
                        WorksheetUtilities.DeleteSheet(sheetImpurity);
                    }
                    else
                    {
                        UpdateImpurity(sheetImpurity, strcmbProtocolType, strcmbProductType, strcmbTestType, numRepsGeneral, 
                            numSamplesImp, strcmbNoS, numPeaksImp, 
                            strcmbOperator1Imp, valAC1Imp, valAC2Imp, strAutoOperator1Imp, strAutoOperator2Imp, valacceptancecriteria1Imp, valacceptancecriteria3Imp, 
                            strcmbOperator2Imp, valAC3Imp, valAC4Imp, valAC5Imp, strcmbOperator4Imp, strcmbOperator5Imp);
                    }

                    if (numSamplesCU == 0)
                    {
                        WorksheetUtilities.DeleteSheet(sheetContentUniformity);
                    }
                    else
                    {
                        UpdateContentUniformity(sheetContentUniformity, strcmbProtocolType, strcmbProductType, strcmbTestType, numSamplesCU);
                    }
                }
            }

            book.Save();
            WorksheetUtilities.ReleaseComObject(book);
            _app.Workbooks.Close();
            _app = null;
            WorksheetUtilities.ReleaseExcelApp();

            return savePath;
        }

        private static void UpdateAssayAPI(Worksheet sheet,
            string strcmbProtocolType, string strcmbProductType, string strcmbTestType,
            int numRepsGeneral,
            int numSamplesAssay, string strcmbRSD, decimal valRSD1
        )
        {
            if (sheet == null)
                return;

            bool wasProtected = WorksheetUtilities.SetSheetProtection(sheet, null, false);

            WorksheetUtilities.SetMetadataValues(sheet, strcmbProtocolType, strcmbProductType, strcmbTestType);

            bool isNDA = (strcmbProtocolType == "NDA");

            // Delete NDA sections if not NDA
            if (!isNDA)
            {
                WorksheetUtilities.DeleteNamedRangeRows(sheet, "NDASection1");
                WorksheetUtilities.DeleteNamedRangeRows(sheet, "NDASection2");
            }

            // set Acceptance Criteria values
            WorksheetUtilities.SetNamedRangeValue(sheet, "RSD_Operator", strcmbRSD, 1, 1);
            WorksheetUtilities.SetNamedRangeValue(sheet, "RSD_Value", valRSD1.ToString(), 1, 1);

            // Process Sample Info
            string sampleInfoNamedRange = WorksheetUtilities.ProcessSampleInfo(sheet, strcmbProductType, numSamplesAssay);

            // Adjust num rows in Raw Data table and Summary tables
            WorksheetUtilities.ResizeNamedRangeRows(sheet, "PrepNumsRawData1", numRepsGeneral);
            WorksheetUtilities.SetSequentialNumbersInNamedRange(sheet, "PrepNumsRawData1", numRepsGeneral);

            WorksheetUtilities.ResizeNamedRangeRows(sheet, "ValidationDataRows", numRepsGeneral);
            WorksheetUtilities.SetSequentialNumbersInNamedRange(sheet, "PrepNumsValidationResults", numRepsGeneral);

            // Handle replicates in Raw Data tables.
            int rawFixedCols = 4;
            for (int i = 2; i <= numSamplesAssay; i++)
            {
                // Handle Raw Data Tables, replicate horizontally
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"RawDataTable{i - 1}", $"RawDataTable{i}", 1, rawFixedCols + 1 /* gapBetweenTables */ + 1, XlPasteType.xlPasteAll);

                // copy over named ranges internal to the Raw Data Table
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"RawBatch{i - 1}", $"RawBatch{i}", 1, rawFixedCols + 1 /* gapBetweenTables */ + 1, XlPasteType.xlPasteAll);
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"RawHeader{i - 1}", $"RawHeader{i}", 1, rawFixedCols + 1 /* gapBetweenTables */ + 1, XlPasteType.xlPasteAll);
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"RawData{i - 1}", $"RawData{i}", 1, rawFixedCols + 1 /* gapBetweenTables */ + 1, XlPasteType.xlPasteAll);
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"RawDataStats{i - 1}", $"RawDataStats{i}", 1, rawFixedCols + 1 /* gapBetweenTables */ + 1, XlPasteType.xlPasteAll);

                if (isNDA)
                {
                    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"RawDataConfidence{i - 1}", $"RawDataConfidence{i}", 1, rawFixedCols + 1 /* gapBetweenTables */ + 1, XlPasteType.xlPasteAll);
                }

                // set label of the table
                WorksheetUtilities.SetNamedRangeValue(sheet, $"RawDataTable{i}", $"Raw Data Table {i}", 1, 1);

                // set column width of last columns of the table.
                WorksheetUtilities.SetColumnWidth(sheet, $"RawDataTable{i}", 1, 13);
                WorksheetUtilities.SetColumnWidth(sheet, $"RawDataTable{i}", 2, 13);
                WorksheetUtilities.SetColumnWidth(sheet, $"RawDataTable{i}", 3, 13);
                WorksheetUtilities.SetColumnWidth(sheet, $"RawDataTable{i}", 4, 13);
            }

            // Link Batch refs to Sample Info Summary cells.
            LinkSampleInfoToBatchCells(sheet, numSamplesAssay, sampleInfoNamedRange);

            // Handle replicates in Summary tables.
            // Height of each summary block (based on ValidationCol1)
            int summaryHeight = WorksheetUtilities.GetNamedRangeRowCount(sheet, "ValidationCol1");

            // Number of 4‑column blocks needed (1–4, 5–8, 9–12, etc.)
            int numBlocks = (numSamplesAssay + 3) / 4;   // ceiling division

            int startRow = WorksheetUtilities.GetNamedRangeStartRow(sheet, "ValidationCol1");
            int colIndex = WorksheetUtilities.GetNamedRangeStartColumn(sheet, "ValidationCol1");


            // ============================================================================
            // 1. VERTICAL REPLICATION
            //    e.g. Copies only the FIRST column of each block:
            //      ValidationCol1 → ValidationCol5
            //      ValidationCol5 → ValidationCol9
            //    Uses vertical offset = summaryHeight + 1 (Insert Copied Cells behavior)
            // ============================================================================

            for (int block = 1; block < numBlocks; block++)
            {
                int dst = block * 4 + 1;   // 5, 9, 13, ...

                // Move startRow to the new block
                startRow = CopyBlockDown(sheet, startRow, summaryHeight, colIndex, dst);

                // Copy summary block internal named ranges
                CopySummaryBlock(sheet, dst - 4, dst, 1 + summaryHeight + 2, 1, isNDA);

                // Link formulas to Raw Data tables
                LinkSummaryBlock(sheet, dst, isNDA);
            }

            // ============================================================================
            // 2. HORIZONTAL REPLICATION
            //    e.g. For each block, copy 3 columns horizontally:
            //      ValidationCol1 → 2 → 3 → 4
            //      ValidationCol5 → 6 → 7 → 8
            //      ValidationCol9 stops early (no 10)
            // ============================================================================

            int fixedSummaryCols = 1;

            for (int block = 0; block < numBlocks; block++)
            {
                int baseCol = block * 4 + 1;

                for (int offset = 1; offset <= 3; offset++)
                {
                    int src = baseCol + offset - 1;
                    int dst = baseCol + offset;

                    if (dst > numSamplesAssay)
                        break;

                    // Copy ValidationColX horizontally
                    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"ValidationCol{src}", $"ValidationCol{dst}", 1, fixedSummaryCols + 1, XlPasteType.xlPasteAll);

                    // Copy summary block internal named ranges
                    CopySummaryBlock(sheet, src, dst, 1, fixedSummaryCols + 1, isNDA);

                    // Link formulas to Raw Data tables
                    LinkSummaryBlock(sheet, dst, isNDA);
                }
            }

            // can remove Autofit in this sheet as column widths are being set manually
            sheet.Columns.AutoFit();

            WorksheetUtilities.DeleteInvalidNamedRanges(sheet);
            WorksheetUtilities.ScrollToTopLeft(sheet);
            if (wasProtected) WorksheetUtilities.SetSheetProtection(sheet, null, true);
        }

        private static void UpdateAssayDP(Worksheet sheet,
            string strcmbProtocolType, string strcmbProductType, string strcmbTestType,
            int numRepsGeneral,
            int numSamplesAssay, string strcmbRSD, decimal valRSD1
        )
        {
            if (sheet == null)
                return;

            bool wasProtected = WorksheetUtilities.SetSheetProtection(sheet, null, false);

            WorksheetUtilities.SetMetadataValues(sheet, strcmbProtocolType, strcmbProductType, strcmbTestType);

            bool isNDA = (strcmbProtocolType == "NDA");

            // Delete NDA sections if not NDA
            if (!isNDA)
            {
                WorksheetUtilities.DeleteNamedRangeRows(sheet, "NDASection1");
                WorksheetUtilities.DeleteNamedRangeRows(sheet, "NDASection2");
            }

            // set Acceptance Criteria values
            WorksheetUtilities.SetNamedRangeValue(sheet, "RSD_Operator", strcmbRSD, 1, 1);
            WorksheetUtilities.SetNamedRangeValue(sheet, "RSD_Value", valRSD1.ToString(), 1, 1);

            // Process Sample Info
            string sampleInfoNamedRange = WorksheetUtilities.ProcessSampleInfo(sheet, strcmbProductType, numSamplesAssay);

            // Adjust num rows in Raw Data table and Summary tables
            WorksheetUtilities.ResizeNamedRangeRows(sheet, "PrepNumsRawData1", numRepsGeneral);
            WorksheetUtilities.SetSequentialNumbersInNamedRange(sheet, "PrepNumsRawData1", numRepsGeneral);

            WorksheetUtilities.ResizeNamedRangeRows(sheet, "ValidationDataRows", numRepsGeneral);
            WorksheetUtilities.SetSequentialNumbersInNamedRange(sheet, "PrepNumsValidationResults", numRepsGeneral);

            // Handle replicates in Raw Data tables.
            int rawFixedCols = 6;
            for (int i = 2; i <= numSamplesAssay; i++)
            {
                // Handle Raw Data Tables, replicate horizontally
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"RawDataTable{i - 1}", $"RawDataTable{i}", 1, rawFixedCols + 1 /* gapBetweenTables */ + 1, XlPasteType.xlPasteAll);

                // copy over named ranges internal to the Raw Data Table
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"RawBatch{i - 1}", $"RawBatch{i}", 1, rawFixedCols + 1 /* gapBetweenTables */ + 1, XlPasteType.xlPasteAll);
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"RawHeader{i - 1}", $"RawHeader{i}", 1, rawFixedCols + 1 /* gapBetweenTables */ + 1, XlPasteType.xlPasteAll);
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"RawData{i - 1}", $"RawData{i}", 1, rawFixedCols + 1 /* gapBetweenTables */ + 1, XlPasteType.xlPasteAll);
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"RawDataStats{i - 1}", $"RawDataStats{i}", 1, rawFixedCols + 1 /* gapBetweenTables */ + 1, XlPasteType.xlPasteAll);

                if (isNDA)
                {
                    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"RawDataConfidence{i - 1}", $"RawDataConfidence{i}", 1, rawFixedCols + 1 /* gapBetweenTables */ + 1, XlPasteType.xlPasteAll);
                }

                // set label of the table
                WorksheetUtilities.SetNamedRangeValue(sheet, $"RawDataTable{i}", $"Raw Data Table {i}", 1, 1);

                // set column width of last columns of the table.
                WorksheetUtilities.SetColumnWidth(sheet, $"RawDataTable{i}", 1, 13);
                WorksheetUtilities.SetColumnWidth(sheet, $"RawDataTable{i}", 2, 13);
                WorksheetUtilities.SetColumnWidth(sheet, $"RawDataTable{i}", 3, 13);
                WorksheetUtilities.SetColumnWidth(sheet, $"RawDataTable{i}", 4, 13);
            }

            // Link Batch refs to Sample Info Summary cells.
            LinkSampleInfoToBatchCells(sheet, numSamplesAssay, sampleInfoNamedRange);

            // Handle replicates in Summary tables.
            // Height of each summary block (based on ValidationCol1)
            int summaryHeight = WorksheetUtilities.GetNamedRangeRowCount(sheet, "ValidationCol1");

            // Number of 4‑column blocks needed (1–4, 5–8, 9–12, etc.)
            int numBlocks = (numSamplesAssay + 3) / 4;   // ceiling division

            int startRow = WorksheetUtilities.GetNamedRangeStartRow(sheet, "ValidationCol1");
            int colIndex = WorksheetUtilities.GetNamedRangeStartColumn(sheet, "ValidationCol1");


            // ============================================================================
            // 1. VERTICAL REPLICATION
            //    e.g. Copies only the FIRST column of each block:
            //      ValidationCol1 → ValidationCol5
            //      ValidationCol5 → ValidationCol9
            //    Uses vertical offset = summaryHeight + 1 (Insert Copied Cells behavior)
            // ============================================================================

            for (int block = 1; block < numBlocks; block++)
            {
                int dst = block * 4 + 1;   // 5, 9, 13, ...

                // Move startRow to the new block
                startRow = CopyBlockDown(sheet, startRow, summaryHeight, colIndex, dst);

                // Copy summary block internal named ranges
                CopySummaryBlock(sheet, dst - 4, dst, 1 + summaryHeight + 2, 1, isNDA);

                // Link formulas to Raw Data tables
                LinkSummaryBlock(sheet, dst, isNDA);
            }

            // ============================================================================
            // 2. HORIZONTAL REPLICATION
            //    e.g. For each block, copy 3 columns horizontally:
            //      ValidationCol1 → 2 → 3 → 4
            //      ValidationCol5 → 6 → 7 → 8
            //      ValidationCol9 stops early (no 10)
            // ============================================================================

            int fixedSummaryCols = 1;

            for (int block = 0; block < numBlocks; block++)
            {
                int baseCol = block * 4 + 1;

                for (int offset = 1; offset <= 3; offset++)
                {
                    int src = baseCol + offset - 1;
                    int dst = baseCol + offset;

                    if (dst > numSamplesAssay)
                        break;

                    // Copy ValidationColX horizontally
                    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"ValidationCol{src}", $"ValidationCol{dst}", 1, fixedSummaryCols + 1, XlPasteType.xlPasteAll);

                    // Copy summary block internal named ranges
                    CopySummaryBlock(sheet, src, dst, 1, fixedSummaryCols + 1, isNDA);

                    // Link formulas to Raw Data tables
                    LinkSummaryBlock(sheet, dst, isNDA);
                }
            }

            // can remove Autofit in this sheet as column widths are being set manually
            // sheet.Columns.AutoFit();

            WorksheetUtilities.DeleteInvalidNamedRanges(sheet);
            WorksheetUtilities.ScrollToTopLeft(sheet);
            if (wasProtected) WorksheetUtilities.SetSheetProtection(sheet, null, true);
        }

        private static void UpdateImpurity(Worksheet sheet,
            string strcmbProtocolType, string strcmbProductType, string strcmbTestType,
            int numRepsGeneral,
            int numSamplesImp, string strcmbNoS, int numPeaksImp,
            string strcmbOperator1Imp, decimal valAC1Imp, decimal valAC2Imp,
            string strAutoOperator1Imp, string strAutoOperator2Imp,
            decimal valacceptancecriteria1Imp, decimal valacceptancecriteria3Imp,
            string strcmbOperator2Imp, decimal valAC3Imp, decimal valAC4Imp, decimal valAC5Imp,
            string strcmbOperator4Imp, string strcmbOperator5Imp
        )
        {
            if (sheet == null)
                return;

            bool wasProtected = WorksheetUtilities.SetSheetProtection(sheet, null, false);

            WorksheetUtilities.SetMetadataValues(sheet, strcmbProtocolType, strcmbProductType, strcmbTestType);

            bool isNDA = (strcmbProtocolType == "NDA");
            bool isPAV = (strcmbProtocolType == "PAV");

            // Delete NDA sections if not NDA
            if (!isNDA)
            {
                WorksheetUtilities.DeleteNamedRangeRows(sheet, "NDASection1");
                WorksheetUtilities.DeleteNamedRangeRows(sheet, "NDASection2");
            }
            else if (!isPAV)
            {
                WorksheetUtilities.DeleteNamedRangeRows(sheet, "PAVSection1");
            }

            // set Acceptance Criteria values
            SetImpurityAcceptanceCriteriaFromIndividualValues(
                sheet,
                "ImpurityAcceptanceCriteriaRange",
                strcmbOperator1Imp,
                valAC1Imp,
                valAC2Imp,
                strAutoOperator1Imp,
                valacceptancecriteria1Imp,
                strcmbOperator2Imp,
                valAC3Imp,
                strcmbOperator4Imp,
                valAC4Imp,
                strAutoOperator2Imp,
                valacceptancecriteria3Imp,
                strcmbOperator5Imp,
                valAC5Imp
            );

            // Process Sample Info
            string sampleInfoNamedRange = WorksheetUtilities.ProcessSampleInfo(sheet, strcmbProductType, numSamplesImp);

            // Adjust num rows in Raw Data table and Summary tables
            WorksheetUtilities.ResizeNamedRangeRows(sheet, "PrepNumsRawData1", numRepsGeneral);
            WorksheetUtilities.SetSequentialNumbersInNamedRange(sheet, "PrepNumsRawData1", numRepsGeneral);

            // Handle replicates of Impurity Peak column horizontally based on num of Peaks
            for (int i = 2; i <= numPeaksImp; i++)
            {
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"PeakCol{i - 1}", $"PeakCol{i}", 1, 1 + 1, XlPasteType.xlPasteAll);
                
                WorksheetUtilities.ResizeNamedRange(sheet, $"RawData1", 0, 1);
                WorksheetUtilities.ResizeNamedRange(sheet, $"RawDataStats1", 0, 1);
                WorksheetUtilities.ResizeNamedRange(sheet, $"RawDataConfidence1", 0, 1);
            }

            // Handle replicates in Raw Data tables vertically based on num Samples
            int namedRangeHeight = WorksheetUtilities.GetNamedRangeRowCount(sheet, "RawDataTable1");

            for (int i = 2; i <= numSamplesImp; i++)
            {
                // Handle Raw Data Tables, replicate horizontally
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"RawDataTable{i - 1}", $"RawDataTable{i}", namedRangeHeight + 1 /* gapBetweenTables */ + 1, 1, XlPasteType.xlPasteAll);

                // copy over named ranges internal to the Raw Data Table
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"RawBatch{i - 1}", $"RawBatch{i}", namedRangeHeight + 1 /* gapBetweenTables */ + 1, 1, XlPasteType.xlPasteAll);
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"RawHeader{i - 1}", $"RawHeader{i}", namedRangeHeight + 1 /* gapBetweenTables */ + 1, 1, XlPasteType.xlPasteAll);
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"RawData{i - 1}", $"RawData{i}", namedRangeHeight + 1 /* gapBetweenTables */ + 1, 1, XlPasteType.xlPasteAll);
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"RawDataStats{i - 1}", $"RawDataStats{i}", namedRangeHeight + 1 /* gapBetweenTables */ + 1, 1, XlPasteType.xlPasteAll);

                if (isNDA)
                {
                    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"RawDataConfidence{i - 1}", $"RawDataConfidence{i}", namedRangeHeight + 1 /* gapBetweenTables */ + 1, 1, XlPasteType.xlPasteAll);
                }

                // set label of the table
                WorksheetUtilities.SetNamedRangeValue(sheet, $"RawDataTable{i}", $"Raw Data Table {i}", 1, 1);

                // set column width of last columns of the table.
                WorksheetUtilities.SetColumnWidth(sheet, $"RawDataTable{i}", 1, 13);
                WorksheetUtilities.SetColumnWidth(sheet, $"RawDataTable{i}", 2, 13);
                WorksheetUtilities.SetColumnWidth(sheet, $"RawDataTable{i}", 3, 13);
                WorksheetUtilities.SetColumnWidth(sheet, $"RawDataTable{i}", 4, 13);
            }

            // can remove Autofit in this sheet as column widths are being set manually
            sheet.Columns.AutoFit();

            WorksheetUtilities.DeleteInvalidNamedRanges(sheet);
            WorksheetUtilities.ScrollToTopLeft(sheet);
            if (wasProtected) WorksheetUtilities.SetSheetProtection(sheet, null, true);
        }

        private static void UpdateContentUniformity(Worksheet sheet,
            string strcmbProtocolType, string strcmbProductType, string strcmbTestType, int numSamplesCU
        )
        {
            if (sheet == null)
                return;

            bool wasProtected = WorksheetUtilities.SetSheetProtection(sheet, null, false);

            WorksheetUtilities.SetMetadataValues(sheet, strcmbProtocolType, strcmbProductType, strcmbTestType);

            bool isNDA = (strcmbProtocolType == "NDA");
            bool isPAV = (strcmbProtocolType == "PAV");

            // Delete NDA sections if not NDA
            if (!isNDA)
            {
                WorksheetUtilities.DeleteNamedRangeRows(sheet, "NDASection1");
            }
            else if (!isPAV)
            {
                WorksheetUtilities.DeleteNamedRangeRows(sheet, "PAVSection1");
            }

            // Process Sample Info
            string sampleInfoNamedRange = WorksheetUtilities.ProcessSampleInfo(sheet, strcmbProductType, numSamplesCU);

            // Handle replicates in Raw Data tables.
            int rawFixedCols = 6;
            for (int i = 2; i <= numSamplesCU; i++)
            {
                // Handle Raw Data Tables, replicate horizontally
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"RawDataTable{i - 1}", $"RawDataTable{i}", 1, rawFixedCols + 1 /* gapBetweenTables */ + 1, XlPasteType.xlPasteAll);

                // copy over named ranges internal to the Raw Data Table
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"RawBatch{i - 1}", $"RawBatch{i}", 1, rawFixedCols + 1 /* gapBetweenTables */ + 1, XlPasteType.xlPasteAll);
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"RawHeader{i - 1}", $"RawHeader{i}", 1, rawFixedCols + 1 /* gapBetweenTables */ + 1, XlPasteType.xlPasteAll);
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"RawData{i - 1}", $"RawData{i}", 1, rawFixedCols + 1 /* gapBetweenTables */ + 1, XlPasteType.xlPasteAll);
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"RawDataStats{i - 1}", $"RawDataStats{i}", 1, rawFixedCols + 1 /* gapBetweenTables */ + 1, XlPasteType.xlPasteAll);
                // have a second set of stats so named it RawDataConfidence1
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"RawDataConfidence{i - 1}", $"RawDataConfidence{i}", 1, rawFixedCols + 1 /* gapBetweenTables */ + 1, XlPasteType.xlPasteAll);

                // set label of the table
                WorksheetUtilities.SetNamedRangeValue(sheet, $"RawDataTable{i}", $"Raw Data Table {i}", 1, 1);

                // set column width of last columns of the table.
                WorksheetUtilities.SetColumnWidth(sheet, $"RawDataTable{i}", 1, 13);
                WorksheetUtilities.SetColumnWidth(sheet, $"RawDataTable{i}", 2, 13);
                WorksheetUtilities.SetColumnWidth(sheet, $"RawDataTable{i}", 3, 13);
                WorksheetUtilities.SetColumnWidth(sheet, $"RawDataTable{i}", 4, 13);
            }

            // Link Batch refs to Sample Info Summary cells.
            LinkSampleInfoToBatchCells(sheet, numSamplesCU, sampleInfoNamedRange);

            if (isNDA)
            {
                // Handle replicates in Summary tables for NDA
                // Height of each summary block (based on ValidationCol1)
                int summaryHeight = WorksheetUtilities.GetNamedRangeRowCount(sheet, "ValidationCol1");

                // Number of 4‑column blocks needed (1–4, 5–8, 9–12, etc.)
                int numBlocks = (numSamplesCU + 3) / 4;   // ceiling division

                int startRow = WorksheetUtilities.GetNamedRangeStartRow(sheet, "ValidationCol1");
                int colIndex = WorksheetUtilities.GetNamedRangeStartColumn(sheet, "ValidationCol1");


                // ============================================================================
                // 1. VERTICAL REPLICATION
                //    e.g. Copies only the FIRST column of each block:
                //      ValidationCol1 → ValidationCol5
                //      ValidationCol5 → ValidationCol9
                //    Uses vertical offset = summaryHeight + 1 (Insert Copied Cells behavior)
                // ============================================================================

                for (int block = 1; block < numBlocks; block++)
                {
                    int dst = block * 4 + 1;   // 5, 9, 13, ...

                    // Move startRow to the new block
                    startRow = CopyBlockDown(sheet, startRow, summaryHeight, colIndex, dst);

                    // Copy summary block internal named ranges
                    CopySummaryBlock(sheet, dst - 4, dst, 1 + summaryHeight + 2, 1, isNDA);

                    // Link formulas to Raw Data tables
                    LinkSummaryBlock(sheet, dst, isNDA);
                }

                // ============================================================================
                // 2. HORIZONTAL REPLICATION
                //    e.g. For each block, copy 3 columns horizontally:
                //      ValidationCol1 → 2 → 3 → 4
                //      ValidationCol5 → 6 → 7 → 8
                //      ValidationCol9 stops early (no 10)
                // ============================================================================

                int fixedSummaryCols = 1;

                for (int block = 0; block < numBlocks; block++)
                {
                    int baseCol = block * 4 + 1;

                    for (int offset = 1; offset <= 3; offset++)
                    {
                        int src = baseCol + offset - 1;
                        int dst = baseCol + offset;

                        if (dst > numSamplesCU)
                            break;

                        // Copy ValidationColX horizontally
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"ValidationCol{src}", $"ValidationCol{dst}", 1, fixedSummaryCols + 1, XlPasteType.xlPasteAll);

                        // Copy summary block internal named ranges
                        CopySummaryBlock(sheet, src, dst, 1, fixedSummaryCols + 1, true /* isNDA */); // isNDA as true to copy over ValidationConfidence1

                        // Link formulas to Raw Data tables
                        LinkSummaryBlock(sheet, dst, true /* isNDA */); // isNDA as true to copy over ValidationConfidence1
                    }
                }
            }
            else if (isPAV)
            {
                // Handle PAV Summary table
                // Expand the results table
                for (int i = 2; i <= numSamplesCU; i++)
                {
                    WorksheetUtilities.InsertRowsAfterForNamedRange(sheet, $"CUSummary{i - 1}", 2, $"CUSummary{i - 1}", true, XlPasteType.xlPasteAll, $"CUSummary{i}");

                    int rowOffset = 1 + 1, colOffset = 1;
                    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"ValidationHeader{i - 1}", $"ValidationHeader{i}", rowOffset, colOffset, XlPasteType.xlPasteAll);
                    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"ValidationData{i - 1}", $"ValidationData{i}", rowOffset, colOffset, XlPasteType.xlPasteAll);
                    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"ValidationStats{i - 1}", $"ValidationStats{i}", rowOffset, colOffset, XlPasteType.xlPasteAll);
                    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"ValidationConfidence{i - 1}", $"ValidationConfidence{i}", rowOffset, colOffset, XlPasteType.xlPasteAll);
                }
            }

            // can remove Autofit in this sheet as column widths are being set manually
            sheet.Columns.AutoFit();

            WorksheetUtilities.DeleteInvalidNamedRanges(sheet);
            WorksheetUtilities.ScrollToTopLeft(sheet);
            if (wasProtected) WorksheetUtilities.SetSheetProtection(sheet, null, true);
        }

        private static void UpdateDissolution(Worksheet sheet,
            string strcmbProtocolType, string strcmbProductType, string strcmbTestType,
            int numRepsGeneral,
            int numSamplesDisso, int numRepsDisso,
            string strcmbOperator1Disso, decimal valAC1Disso,
            string strcmbOperator2Disso, decimal valAC2Disso,
            string strcmbOperator3Disso,
            string strcmbOperator4Disso, decimal valAC3Disso,
            string strcmbOperator5Disso, decimal valAC4Disso,
            string strcmbOperator6Disso
        )
        {
            if (sheet == null)
                return;

            bool wasProtected = WorksheetUtilities.SetSheetProtection(sheet, null, false);

            WorksheetUtilities.SetMetadataValues(sheet, strcmbProtocolType, strcmbProductType, strcmbTestType);

            // Process Sample Info
            string sampleInfoNamedRange = WorksheetUtilities.ProcessSampleInfo(sheet, strcmbProductType, numSamplesDisso);

            HandleDissolution(sheet, numRepsDisso, numSamplesDisso);

            WorksheetUtilities.DeleteInvalidNamedRanges(sheet);
            WorksheetUtilities.ScrollToTopLeft(sheet);
            if (wasProtected) WorksheetUtilities.SetSheetProtection(sheet, null, true);
        }

        private static void UpdateWaterContent(Worksheet sheet,
            string strcmbProtocolType, string strcmbProductType, string strcmbTestType,
            int numRepsGeneral,
            int numSamplesWC,
            string strcmbWaterContent1WC, decimal valAC6WC,
            string strcmbWaterContent3WC, decimal valAC7WC,
            string strcmbWaterContent2WC, decimal valAC8WC,
            string strcmbWaterContent4WC, decimal valAC9WC
        )
        {
            if (sheet == null)
                return;

            bool wasProtected = WorksheetUtilities.SetSheetProtection(sheet, null, false);

            WorksheetUtilities.SetMetadataValues(sheet, strcmbProtocolType, strcmbProductType, strcmbTestType);

            // set Acceptance Criteria values
            // Row 1
            WorksheetUtilities.SetNamedRangeValue(sheet, "AcceptanceCriteriaRange", strcmbWaterContent1WC, 1, 1);
            WorksheetUtilities.SetNamedRangeValue(sheet, "AcceptanceCriteriaRange", valAC6WC.ToString(), 1, 2);
            WorksheetUtilities.SetNamedRangeValue(sheet, "AcceptanceCriteriaRange", strcmbWaterContent2WC, 1, 3);
            WorksheetUtilities.SetNamedRangeValue(sheet, "AcceptanceCriteriaRange", valAC8WC.ToString(), 1, 4);

            // Row 2
            WorksheetUtilities.SetNamedRangeValue(sheet, "AcceptanceCriteriaRange", strcmbWaterContent3WC, 2, 1);
            WorksheetUtilities.SetNamedRangeValue(sheet, "AcceptanceCriteriaRange", valAC7WC.ToString(), 2, 2);
            WorksheetUtilities.SetNamedRangeValue(sheet, "AcceptanceCriteriaRange", strcmbWaterContent4WC, 2, 3);
            WorksheetUtilities.SetNamedRangeValue(sheet, "AcceptanceCriteriaRange", valAC9WC.ToString(), 2, 4);

            // Process Sample Info
            string sampleInfoNamedRange = WorksheetUtilities.ProcessSampleInfo(sheet, strcmbProductType, numSamplesWC);

            // Adjust num rows in Raw Data table and Summary tables
            WorksheetUtilities.ResizeNamedRangeRows(sheet, "PrepNumsRawData1", numRepsGeneral);
            WorksheetUtilities.SetSequentialNumbersInNamedRange(sheet, "PrepNumsRawData1", numRepsGeneral);

            WorksheetUtilities.ResizeNamedRangeRows(sheet, "ValidationDataRows", numRepsGeneral);
            WorksheetUtilities.SetSequentialNumbersInNamedRange(sheet, "PrepNumsValidationResults", numRepsGeneral);

            // Handle replicates in Raw Data tables.
            int rawFixedCols = 2;
            for (int i = 2; i <= numSamplesWC; i++)
            {
                // Handle Raw Data Tables, replicate horizontally
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"RawDataTable{i - 1}", $"RawDataTable{i}", 1, rawFixedCols + 1 /* gapBetweenTables */ + 1, XlPasteType.xlPasteAll);

                // copy over named ranges internal to the Raw Data Table
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"RawBatch{i - 1}", $"RawBatch{i}", 1, rawFixedCols + 1 /* gapBetweenTables */ + 1, XlPasteType.xlPasteAll);
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"RawHeader{i - 1}", $"RawHeader{i}", 1, rawFixedCols + 1 /* gapBetweenTables */ + 1, XlPasteType.xlPasteAll);
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"RawData{i - 1}", $"RawData{i}", 1, rawFixedCols + 1 /* gapBetweenTables */ + 1, XlPasteType.xlPasteAll);
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"RawDataStats{i - 1}", $"RawDataStats{i}", 1, rawFixedCols + 1 /* gapBetweenTables */ + 1, XlPasteType.xlPasteAll);

                // set label of the table
                WorksheetUtilities.SetNamedRangeValue(sheet, $"RawDataTable{i}", $"Raw Data Table {i}", 1, 1);

                // set column width of last columns of the table.
                WorksheetUtilities.SetColumnWidth(sheet, $"RawDataTable{i}", 1, 13);
                WorksheetUtilities.SetColumnWidth(sheet, $"RawDataTable{i}", 2, 13);
            }

            // Link Batch refs to Sample Info Summary cells.
            LinkSampleInfoToBatchCells(sheet, numSamplesWC, sampleInfoNamedRange);

            // Handle replicates in Summary tables.
            // Height of each summary block (based on ValidationCol1)
            int summaryHeight = WorksheetUtilities.GetNamedRangeRowCount(sheet, "ValidationCol1");

            // Number of 4‑column blocks needed (1–4, 5–8, 9–12, etc.)
            int numBlocks = (numSamplesWC + 3) / 4;   // ceiling division

            int startRow = WorksheetUtilities.GetNamedRangeStartRow(sheet, "ValidationCol1");
            int colIndex = WorksheetUtilities.GetNamedRangeStartColumn(sheet, "ValidationCol1");


            // ============================================================================
            // 1. VERTICAL REPLICATION
            //    e.g. Copies only the FIRST column of each block:
            //      ValidationCol1 → ValidationCol5
            //      ValidationCol5 → ValidationCol9
            //    Uses vertical offset = summaryHeight + 1 (Insert Copied Cells behavior)
            // ============================================================================

            for (int block = 1; block < numBlocks; block++)
            {
                int dst = block * 4 + 1;   // 5, 9, 13, ...

                // Move startRow to the new block
                startRow = CopyBlockDown(sheet, startRow, summaryHeight, colIndex, dst);

                // Copy summary block internal named ranges
                CopySummaryBlock(sheet, dst - 4, dst, 1 + summaryHeight + 2, 1, false);

                // Link formulas to Raw Data tables
                LinkSummaryBlock(sheet, dst, false);
            }

            // ============================================================================
            // 2. HORIZONTAL REPLICATION
            //    e.g. For each block, copy 3 columns horizontally:
            //      ValidationCol1 → 2 → 3 → 4
            //      ValidationCol5 → 6 → 7 → 8
            //      ValidationCol9 stops early (no 10)
            // ============================================================================

            int fixedSummaryCols = 1;

            for (int block = 0; block < numBlocks; block++)
            {
                int baseCol = block * 4 + 1;

                for (int offset = 1; offset <= 3; offset++)
                {
                    int src = baseCol + offset - 1;
                    int dst = baseCol + offset;

                    if (dst > numSamplesWC)
                        break;

                    // Copy ValidationColX horizontally
                    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"ValidationCol{src}", $"ValidationCol{dst}", 1, fixedSummaryCols + 1, XlPasteType.xlPasteAll);

                    // Copy summary block internal named ranges
                    CopySummaryBlock(sheet, src, dst, 1, fixedSummaryCols + 1, false);

                    // Link formulas to Raw Data tables
                    LinkSummaryBlock(sheet, dst, false);
                }
            }

            // can remove Autofit in this sheet as column widths are being set manually
            sheet.Columns.AutoFit();

            WorksheetUtilities.DeleteInvalidNamedRanges(sheet);
            WorksheetUtilities.ScrollToTopLeft(sheet);
            if (wasProtected) WorksheetUtilities.SetSheetProtection(sheet, null, true);
        }

        // helper methods
        // Assumes the namedrange of target series is RawBatch{i}
        private static void LinkSampleInfoToBatchCells(Worksheet sheet, int numSamplesAssay, string sampleInfoNamedRange)
        {
            // Determine last column of sampleInfoNamedRange
            Range sampleInfoRange = sheet.Range[sampleInfoNamedRange];
            int lastCol = sampleInfoRange.Columns.Count;

            for (int i = 1; i <= numSamplesAssay; i++)
            {
                // Get the target cell inside sampleInfoNamedRange (row i, last column)
                string targetAddress = WorksheetUtilities.GetCellAddress(sheet, sampleInfoNamedRange, i, lastCol);

                // Build formula referencing that cell
                string formula = WorksheetUtilities.GetSimpleReferenceFormula(targetAddress);

                // Set formula into RawBatch{i} cell (1,1)
                WorksheetUtilities.SetNamedRangeFormula(sheet, $"RawBatch{i}", formula, 1, 1);
            }
        }

        private static void CopySummaryBlock(Worksheet sheet, int src, int dst, int rowOffset, int colOffset, bool isNDA)
        {
            WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"ValidationHeader{src}", $"ValidationHeader{dst}", rowOffset, colOffset, XlPasteType.xlPasteAll);
            WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"ValidationData{src}", $"ValidationData{dst}", rowOffset, colOffset, XlPasteType.xlPasteAll);
            WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"ValidationStats{src}", $"ValidationStats{dst}", rowOffset, colOffset, XlPasteType.xlPasteAll);

            if (isNDA)
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"ValidationConfidence{src}", $"ValidationConfidence{dst}", rowOffset, colOffset, XlPasteType.xlPasteAll);
        }

        private static void LinkSummaryBlock(Worksheet sheet, int dst, bool isNDA)
        {
            WorksheetUtilities.LinkVerticalNamedRanges(sheet, $"RawHeader{dst}", $"ValidationHeader{dst}");
            WorksheetUtilities.LinkVerticalNamedRanges(sheet, $"RawData{dst}", $"ValidationData{dst}");
            WorksheetUtilities.LinkVerticalNamedRanges(sheet, $"RawDataStats{dst}", $"ValidationStats{dst}");

            if (isNDA)
                WorksheetUtilities.LinkVerticalNamedRanges(sheet, $"RawDataConfidence{dst}", $"ValidationConfidence{dst}");
        }

        private static int CopyBlockDown(Worksheet sheet, int startRow, int summaryHeight, int colIndex, int dst)
        {
            // Copy entire row block (all columns)
            Range blockToCopy = sheet.Rows[$"{startRow}:{startRow + summaryHeight + 1}"];
            blockToCopy.Copy();

            // Insert copied rows with one-row gap
            Range insertPoint = sheet.Rows[startRow + summaryHeight + 2];
            insertPoint.Insert(XlInsertShiftDirection.xlShiftDown);

            // Move startRow to the new block
            startRow += summaryHeight + 2;

            // Reassign named range ValidationColX
            Range newRange = sheet.Range[
                sheet.Cells[startRow, colIndex],
                sheet.Cells[startRow + summaryHeight - 1, colIndex]
            ];

            sheet.Names.Add(Name: $"ValidationCol{dst}", RefersTo: $"={newRange.Address}");

            return startRow;
        }

        private static void SetImpurityAcceptanceCriteriaFromIndividualValues(
            Worksheet sheet,
            string accCriteriaRangeName,
            string impurityOperator1,
            decimal impurityValue1,
            decimal impurityValue2,
            string impurityAutoOperator1,
            decimal impurityAutoValue1,
            string impurityOperator2,
            decimal impurityValue3,
            string impurityOperator4,
            decimal impurityValue4,
            string impurityAutoOperator2,
            decimal impurityAutoValue2,
            string impurityOperator5,
            decimal impurityValue5)
        {
            string rangeName = accCriteriaRangeName;

            // --- Row 1 ---
            WorksheetUtilities.SetNamedRangeValue(sheet, rangeName, impurityOperator1, 1, 4);
            WorksheetUtilities.SetNamedRangeValue(sheet, rangeName, impurityValue1.ToString(), 1, 5);
            WorksheetUtilities.SetNamedRangeValue(sheet, rangeName, impurityValue2.ToString(), 1, 7);

            // --- Row 2 ---
            WorksheetUtilities.SetNamedRangeValue(sheet, rangeName, impurityAutoOperator1, 2, 1);
            WorksheetUtilities.SetNamedRangeValue(sheet, rangeName, impurityAutoValue1.ToString(), 2, 2);
            WorksheetUtilities.SetNamedRangeValue(sheet, rangeName, impurityOperator2, 2, 4);
            WorksheetUtilities.SetNamedRangeValue(sheet, rangeName, impurityValue3.ToString(), 2, 5);
            WorksheetUtilities.SetNamedRangeValue(sheet, rangeName, impurityOperator4, 2, 6);
            WorksheetUtilities.SetNamedRangeValue(sheet, rangeName, impurityValue4.ToString(), 2, 7);

            // --- Row 3 ---
            WorksheetUtilities.SetNamedRangeValue(sheet, rangeName, impurityAutoOperator2, 3, 1);
            WorksheetUtilities.SetNamedRangeValue(sheet, rangeName, impurityAutoValue2.ToString(), 3, 2);
            WorksheetUtilities.SetNamedRangeValue(sheet, rangeName, impurityOperator5, 3, 6);
            WorksheetUtilities.SetNamedRangeValue(sheet, rangeName, impurityValue5.ToString(), 3, 7);
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

        private static void UpdateDissolutionFormulas(Worksheet sheet, string namedRangeBaseName, string referenceRange, int baseRow, int baseCol)
        {
            // Get main range using helper
            Range range = WorksheetUtilities.GetNamedRange(sheet, namedRangeBaseName);
            if (range == null)
                return;

            // Get reference range using helper
            Range refRange = WorksheetUtilities.GetNamedRange(sheet, referenceRange);
            if (refRange == null)
                return;

            int numRows = range.Rows.Count;
            int numCols = range.Columns.Count;

            for (int col = 0; col < numCols; col++)
            {
                // Determine reference column (1–6 map directly, 7+ map to 6)
                int refCol = (col + 1 < 7) ? col + 1 : 6;

                // Get reference cell address via helper
                string refAddress = WorksheetUtilities.GetCellAddress(sheet, referenceRange, 1, refCol, rowAbsolute: false, colAbsolute: false);

                for (int row = 1; row < numRows; row++)
                {
                    // Column 4 does not use formulas
                    if (col + 1 == 4)
                        continue;

                    Range cell = range.Cells[row + 1, col + 1] as Range;
                    if (cell == null)
                    {
                        WorksheetUtilities.ReleaseComObject(cell);
                        continue;
                    }

                    if (cell.HasFormula)
                    {
                        string formula = cell.Formula;

                        // Skip if already correct
                        if (!formula.Contains(refAddress))
                        {
                            // Extract the simple reference from the formula
                            string simpleRef = WorksheetUtilities.GetSimpleReferenceFormula(formula);

                            if (!string.IsNullOrEmpty(simpleRef) && simpleRef != refAddress)
                            {
                                string updatedFormula = formula.Replace(simpleRef, refAddress);

                                // Use helper to set the formula
                                WorksheetUtilities.SetNamedRangeFormula(sheet, namedRangeBaseName, updatedFormula, row + 1, col + 1);
                            }
                        }
                    }

                    WorksheetUtilities.ReleaseComObject(cell);
                }
            }

            // Cleanup
            WorksheetUtilities.ReleaseComObject(range);
            WorksheetUtilities.ReleaseComObject(refRange);
        }

        private static void UpdateFormulasInNamedRange(Worksheet sheet, string namedRangeBase, int rangeCount, string sourceString, string destString)
        {
            Range tableRange = null;
            Range cell = null;

            try
            {
                for (int index = 2; index <= rangeCount; index++)
                {
                    string fullRangeName = namedRangeBase + index;

                    // Use helper to get the named range
                    tableRange = WorksheetUtilities.GetNamedRange(sheet, fullRangeName);
                    if (tableRange == null)
                        continue;

                    int rows = tableRange.Rows.Count;
                    int cols = tableRange.Columns.Count;

                    for (int r = 1; r <= rows; r++)
                    {
                        for (int c = 1; c <= cols; c++)
                        {
                            cell = tableRange.Cells[r, c] as Range;

                            if (cell != null && cell.HasFormula)
                            {
                                string updatedFormula = cell.Formula.Replace(sourceString, destString + index);
                                cell.Formula = updatedFormula;
                            }

                            WorksheetUtilities.ReleaseComObject(cell);
                            cell = null;
                        }
                    }

                    WorksheetUtilities.ReleaseComObject(tableRange);
                    tableRange = null;
                }
            }
            finally
            {
                // Ensure cleanup even if exceptions occur
                WorksheetUtilities.ReleaseComObject(cell);
                WorksheetUtilities.ReleaseComObject(tableRange);

                sheet = null;
            }
        }
    }
}
