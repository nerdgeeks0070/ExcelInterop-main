using log4net.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using ExcelRange = Microsoft.Office.Interop.Excel.Range;

namespace Spreadsheet.Handler
{
    public class SampleRepeatabilityApi
    {
        private static Application _app;

        private const int DefaultNumReplicates = 6;
        private const int MinNumReplicates = 2;
        private const int DefaultNumDataTables = 1;
        private const string DefaultApiDosageUnits = "mg/ml";

        // This is the raw data table column count + 1 (as a column spacer)
        private const int RawDataColOffset = 7;

        private const string TempDirectoryName = "ABD_TempFiles";

        public static string UpdateSampleRepeatabilityApiSheet(string sourcePath,
            // general
            string protocolType,
            string productType,
            string testType,
            int numReplicates,
            // assay
            int assayNumSamples,
            string rsdOperator,
            decimal rsdValue,
            // impurity
            int impurityNumSamples,
            int impurityNumPeaks,
            string impurityQuantitationType,
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
            decimal impurityValue5,
            // water content
            int wcNumSamples,
            string wcOperator1,
            decimal wcValue1,
            string wcOperator2,
            decimal wcValue2,
            string wcOperator3,
            decimal wcValue3,
            string wcOperator4,
            decimal wcValue4)
        {
            string returnPath = "";
            try
            {
                returnPath = UpdateSampleRepeatabilityApiSheet2(
                    sourcePath,
                    numReplicates,
                    protocolType,
                    productType,
                    testType,
                    assayNumSamples,
                    rsdOperator,
                    rsdValue,
                    impurityNumSamples,
                    impurityQuantitationType,
                    impurityNumPeaks,
                    impurityOperator1,
                    impurityValue1,
                    impurityValue2,
                    impurityAutoOperator1,
                    impurityAutoValue1,
                    impurityOperator2,
                    impurityValue3,
                    impurityOperator4,
                    impurityValue4,
                    impurityAutoOperator2,
                    impurityAutoValue2,
                    impurityOperator5,
                    impurityValue5,
                    wcNumSamples,
                    wcOperator1,
                    wcValue1,
                    wcOperator2,
                    wcValue2,
                    wcOperator3,
                    wcValue3,
                    wcOperator4,
                    wcValue4
                );
            }
            catch (Exception ex)
            {
                Logger.LogMessage("An error occurred in the call to SampleRepeatability.UpdateSampleRpeatabilityAPI. Message and stack trace are:\r\n" + ex.Message + "\r\n" + ex.StackTrace, Level.Error);

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
                            Logger.LogMessage("An error occurred in the call to SampleRepeatability.UpdateSampleRpeatabilityAPI. Failed to save current workbook changes and to get path.", Level.Error);
                        }

                        _app.Workbooks.Close();
                    }
                    _app = null;
                }
                catch
                {
                    Logger.LogMessage("An error occurred in the call to SampleRepeatability.UpdateSampleRpeatabilityAPI. Application failed to close workbooks. Message and stack trace are:\r\n" + ex.Message + "\r\n" + ex.StackTrace, Level.Error);
                }
                finally
                {
                    WorksheetUtilities.ReleaseExcelApp();
                }
            }
            return returnPath;
        }

        private static string UpdateSampleRepeatabilityApiSheet2(string sourcePath,
            int numReplicates,
            string protocolType,
            string productType,
            string testType,
            // assay
            int assayNumSamples,
            string rsdOperator,
            decimal rsdValue,
            // impurity
            int impurityNumSamples,
            string impurityQuantitationType,
            int impurityNumPeaks,
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
            decimal impurityValue5,
            // water content
            int wcNumSamples,
            string wcOperator1,
            decimal wcValue1,
            string wcOperator2,
            decimal wcValue2,
            string wcOperator3,
            decimal wcValue3,
            string wcOperator4,
            decimal wcValue4,
            bool chkAssay = true,
            bool chkImpurity = true,
            bool chkWaterContent = true)
        {
            if (!File.Exists(sourcePath))
            {
                Logger.LogMessage("Error in call to SampleRepeatability.UpdateSampleRpeatabilityAPI. Invalid source file path specified.", Level.Error);
                return "";
            }

            // Generate an random temp path to save new workbook
            string savePath = WorksheetUtilities.CopyWorkbook(sourcePath, TempDirectoryName, "Sample Repeatability API Results.xls");
            if (String.IsNullOrEmpty(savePath)) return "";

            // Try to open the file
            _app = WorksheetUtilities.GetExcelApp();
            _app.Workbooks.Open(savePath, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            Workbook book = _app.Workbooks[1];
            Worksheet sheet = book.Worksheets[1] as Worksheet;

            if (sheet != null)
            {
                bool wasProtected = WorksheetUtilities.SetSheetProtection(sheet, null, false);
                if (chkAssay)
                {
                    HandleAssay(sheet, numReplicates, assayNumSamples, rsdOperator, rsdValue);
                    HandleAssayValidation(sheet, numReplicates, assayNumSamples);
                }
                if (chkImpurity)
                {
                    Console.WriteLine($"=== Calling HandleImpurity with numReplicates={numReplicates}, impurityNumSamples={impurityNumSamples}, numPeaks={impurityNumPeaks}, ===");

                    HandleImpurity(
                        sheet,
                        numReplicates,
                        impurityNumSamples,
                        impurityNumPeaks,
                        impurityQuantitationType,

                        impurityOperator1,
                        impurityValue1,
                        impurityValue2,

                        impurityAutoOperator1,
                        impurityAutoValue1,

                        impurityOperator2,
                        impurityValue3,

                        impurityOperator4,
                        impurityValue4,

                        impurityAutoOperator2,
                        impurityAutoValue2,

                        impurityOperator5,
                        impurityValue5);
                }
                else
                {
                    Console.WriteLine("=== Impurity section SKIPPED (chkImpurity=false) ===");
                }

                if (chkWaterContent)
                {
                    HandleWaterContent(sheet, numReplicates,
                        wcNumSamples,
                        wcOperator1,
                        wcValue1,
                        wcOperator2,
                        wcValue2,
                        wcOperator3,
                        wcValue3,
                        wcOperator4,
                        wcValue4);
                    HandleWaterContentSummary(sheet, numReplicates, wcNumSamples);
                }

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

                //Remove Water Content name range if Water Content checkbox is false
                if (chkWaterContent != true)
                {
                    WorksheetUtilities.DeleteNamedRangeRows(sheet, "WaterContent");
                    WorksheetUtilities.DeleteNamedRange(sheet, "WaterContent");
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

        private static void HandleAssay(Worksheet sheet, int numReplicates, int assayNumSamples, string assayRsdOperator, decimal assayRsdValue)
        {
            Console.WriteLine($"\n=== Processing Assay section with {numReplicates} replicates and {assayNumSamples} samples ===");

            SetSingleCellValue(sheet, "AcceptanceCriteriaAssayOperator", assayRsdOperator);
            SetSingleCellValue(sheet, "AcceptanceCriteriaAssayRSD", assayRsdValue.ToString());

            // Step 1: Row Expansion FIRST (for both raw data AND validation)
            if (numReplicates > 6) // Default is 6, expand if more
            {
                int numRowsToInsert = numReplicates - 6;
                Console.WriteLine($"Expanding tables by {numRowsToInsert} rows...");

                // Expand raw data prep numbers
                WorksheetUtilities.InsertRowsIntoNamedRange(numRowsToInsert, sheet, "PrepNumsRawData1", true, XlDirection.xlDown, XlPasteType.xlPasteFormulas);

                // // Expand validation prep numbers (DON'T fill rows - we'll renumber later)
                // if (WorksheetUtilities.NamedRangeExist(sheet, "PrepNumsValidationResults"))
                // {
                //     WorksheetUtilities.InsertRowsIntoNamedRange(numRowsToInsert, sheet, "PrepNumsValidationResults", false, XlDirection.xlDown, XlPasteType.xlPasteFormulas);
                //     Console.WriteLine($"✅ Expanded validation prep numbers by {numRowsToInsert} rows");
                // }

                // // Expand validation data (DON'T fill rows - let the linking handle this)
                // if (WorksheetUtilities.NamedRangeExist(sheet, "ValidationTableValues"))
                // {
                //     WorksheetUtilities.InsertRowsIntoNamedRange(numRowsToInsert, sheet, "ValidationTableValues", false, XlDirection.xlDown, XlPasteType.xlPasteFormulas);
                //     Console.WriteLine($"✅ Expanded validation data by {numRowsToInsert} rows");
                // }
            }
            else if (numReplicates < 6) // Contract if less than 6
            {
                int numRowsToRemove = 6 - numReplicates;
                if (numRowsToRemove > 4) numRowsToRemove = 4; // Keep minimum 2 rows
                Console.WriteLine($"Contracting tables by {numRowsToRemove} rows...");

                WorksheetUtilities.DeleteRowsFromNamedRange(numRowsToRemove, sheet, "PrepNumsRawData1", XlDirection.xlDown);
            }

            // Step 2: Update preparation numbers (1, 2, 3, ..., numReplicates) - BOTH raw data AND validation
            if (numReplicates != 6) // Only if we changed the row count
            {
                Console.WriteLine($"Updating preparation numbers 1-{numReplicates}...");
                List<string> prepNumbers = new List<string>();
                for (int i = 1; i <= numReplicates; i++)
                {
                    prepNumbers.Add(i.ToString());
                }

                // Update raw data prep numbers
                WorksheetUtilities.SetNamedRangeValues(sheet, "PrepNumsRawData1", prepNumbers);
            }

            // Step 3: Raw Data Table Copying (Horizontal Table Copying)
            if (assayNumSamples > 1)
            {
                Console.WriteLine($"Creating {assayNumSamples - 1} additional sample tables...");

                var table1Range = WorksheetUtilities.GetNamedRange(sheet, "RawDataTable1");
                if (table1Range == null)
                {
                    Console.WriteLine("❌ Error: Could not find RawDataTable1 named range");
                    return;
                }

                int table1StartCol = table1Range.Column;
                int table1Width = table1Range.Columns.Count;
                int spacerColumns = 1;

                for (int sampleNum = 2; sampleNum <= assayNumSamples; sampleNum++)
                {
                    int colOffset = (sampleNum - 1) * (table1Width + spacerColumns);
                    int expectedStartCol = table1StartCol + colOffset;

                    Console.WriteLine($"Creating Sample {sampleNum} table at column {expectedStartCol}...");

                    string destTableName = $"RawDataTable{sampleNum}";

                    try
                    {
                        WorksheetUtilities.CopyNamedRangeToNewLocationWithNewNamedRange(sheet, "RawDataTable1", destTableName, table1Range.Row, expectedStartCol, XlPasteType.xlPasteAll);

                        CopyColumnWidths(sheet, "RawDataTable1", destTableName);

                        // Update header text
                        var headerRange = WorksheetUtilities.GetNamedRange(sheet, destTableName);
                        if (headerRange != null)
                        {
                            bool headerFixed = false;
                            for (int row = 1; row <= 3 && !headerFixed; row++)
                            {
                                for (int col = 1; col <= headerRange.Columns.Count && !headerFixed; col++)
                                {
                                    var cell = headerRange.Cells[row, col] as ExcelRange;
                                    if (cell.Value2.ToString().Contains("Raw Data Table") == true)
                                    {
                                        cell.Value2 = $"Raw Data Table {sampleNum}";
                                        Console.WriteLine($"✅ Updated header to 'Raw Data Table {sampleNum}'");
                                        headerFixed = true;
                                    }
                                }
                            }
                        }

                        // Create individual named ranges for the copied table components
                        CreateIndividualNamedRangesForTable(sheet, sampleNum, colOffset);

                        Console.WriteLine($"✅ Successfully created Sample {sampleNum} table");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"❌ Error creating Sample {sampleNum} table: {ex.Message}");
                    }
                }
            }

            Console.WriteLine("✅ Assay section processing complete!");
        }

        // Add this debug method to see what's actually happening with the tables

        private static void DebugTablePositions(Worksheet sheet)
        {
            Console.WriteLine("\n🔍 DEBUG: Analyzing table positions...");

            var table1 = WorksheetUtilities.GetNamedRange(sheet, "RawDataTable1");
            if (table1 != null)
            {
                Console.WriteLine($"RawDataTable1: Column {table1.Column} to {table1.Column + table1.Columns.Count - 1} (Width: {table1.Columns.Count})");
                Console.WriteLine($"  Address: {table1.Address}");
            }

            var table2 = WorksheetUtilities.GetNamedRange(sheet, "RawDataTable2");
            if (table2 != null)
            {
                Console.WriteLine($"RawDataTable2: Column {table2.Column} to {table2.Column + table2.Columns.Count - 1} (Width: {table2.Columns.Count})");
                Console.WriteLine($"  Address: {table2.Address}");

                if (table1 != null)
                {
                    int gap = table2.Column - (table1.Column + table1.Columns.Count);
                    Console.WriteLine($"  Gap between Table1 and Table2: {gap} columns");
                }
            }

            var table3 = WorksheetUtilities.GetNamedRange(sheet, "RawDataTable3");
            if (table3 != null)
            {
                Console.WriteLine($"RawDataTable3: Column {table3.Column} to {table3.Column + table3.Columns.Count - 1} (Width: {table3.Columns.Count})");
                Console.WriteLine($"  Address: {table3.Address}");

                if (table2 != null)
                {
                    int gap = table3.Column - (table2.Column + table2.Columns.Count);
                    Console.WriteLine($"  Gap between Table2 and Table3: {gap} columns");
                }
            }
        }

        private static void HandleAssayValidation(Worksheet sheet, int numReplicates, int assayNumSamples)
        {
            Console.WriteLine($"\n=== Processing Assay Validation section with {numReplicates} replicates ===");

            // Step 1: Row Expansion based on numReplicates (FIRST)
            if (numReplicates > 6) // Default is 6, expand if more
            {
                int numRowsToInsert = numReplicates - 6;
                Console.WriteLine($"Expanding Assay Validation by {numRowsToInsert} rows...");

                // Only expand the prep column - other columns expand automatically
                WorksheetUtilities.InsertRowsIntoNamedRange(numRowsToInsert, sheet, "PrepNumsValidationResults", true, XlDirection.xlDown, XlPasteType.xlPasteFormulas);
            }
            else if (numReplicates < 6) // Contract if less than 6
            {
                int numRowsToRemove = 6 - numReplicates;
                if (numRowsToRemove > 4) numRowsToRemove = 4; // Keep minimum 2 rows
                Console.WriteLine($"Contracting Assay Validation by {numRowsToRemove} rows...");

                WorksheetUtilities.DeleteRowsFromNamedRange(numRowsToRemove, sheet, "PrepNumsValidationResults", XlDirection.xlDown);
            }

            // Step 2: Update preparation numbers
            if (numReplicates != 6) // Only if we changed the row count
            {
                Console.WriteLine($"Updating Assay Validation preparation numbers 1-{numReplicates}...");

                List<string> prepNumbers = new List<string>();
                for (int i = 1; i <= numReplicates; i++)
                {
                    prepNumbers.Add(i.ToString());
                }

                WorksheetUtilities.SetNamedRangeValues(sheet, "PrepNumsValidationResults", prepNumbers);
            }

            // Step 3: Column Copying based on assayNumSamples (NEW)
            if (assayNumSamples > 1)
            {
                Console.WriteLine($"Creating {assayNumSamples - 1} additional Assay validation columns...");

                for (int sampleNum = 2; sampleNum <= assayNumSamples; sampleNum++)
                {
                    int colOffset = sampleNum - 1; // Sample 2 = 1 col right (D), Sample 3 = 2 cols right (E)

                    Console.WriteLine($"Creating Assay Validation Sample {sampleNum} column at column offset {colOffset}...");

                    try
                    {
                        // Copy entire validation column (Assay__Validation → Assay__Validation2, etc.)
                        string destValidationName = $"Assay__Validation{sampleNum}";
                        Console.WriteLine($"Copying: Assay__Validation → {destValidationName}");

                        // Assay__Validation: Copy EVERYTHING (formulas + formats + values)
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "Assay__Validation", destValidationName, 1, colOffset + 1, XlPasteType.xlPasteFormats);
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "Assay__Validation", destValidationName, 1, colOffset + 1, XlPasteType.xlPasteValues);

                        // ValidationTableValues: Copy FORMATS + VALUES (no formulas)
                        if (WorksheetUtilities.NamedRangeExist(sheet, "ValidationTableValues"))
                        {
                            string destDataRange = $"ValidationTableValues{sampleNum}";

                            // Copy formats first
                            WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ValidationTableValues", destDataRange, 1, colOffset + 1, XlPasteType.xlPasteFormats);
                            // Copy values second
                            WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ValidationTableValues", destDataRange, 1, colOffset + 1, XlPasteType.xlPasteValues);

                            Console.WriteLine($"✅ Created {destDataRange} (formats + values) at column offset {colOffset + 1}");
                        }
                        else
                        {
                            Console.WriteLine($"⚠️ Warning: ValidationTableValues not found, skipping ValidationTableValues{sampleNum}");
                        }

                        // Copy column width
                        var sourceColumn = sheet.Columns[3] as ExcelRange; // Column C
                        var destColumn = sheet.Columns[3 + colOffset + 1] as ExcelRange; // Target column
                        if (sourceColumn != null && destColumn != null)
                            destColumn.ColumnWidth = sourceColumn.ColumnWidth;

                        Console.WriteLine($"✅ Successfully created Assay Validation Sample {sampleNum} column");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"❌ Error creating Assay Validation Sample {sampleNum} column: {ex.Message}");
                    }
                }
            }
            else
            {
                Console.WriteLine("No additional Assay validation columns needed (assayNumSamples = 1)");
            }
            UpdateValidationFormulas(sheet, numReplicates, assayNumSamples);
            UpdateValidationStatsFormulas(sheet, assayNumSamples);           // Links statistics with correct precision names
            UpdateValidationConfidenceFormulas(sheet, assayNumSamples);      // Links confidence intervals (NEW)

            Console.WriteLine("✅ Assay Validation section processing complete!");
        }

        private static void CreateIndividualNamedRangesForTable(Worksheet sheet, int sampleNum, int colOffset)
        {
            // Create named ranges for the individual components of the copied table
            // Based on the original pattern: PrepNumsRawData1, RawData1, RawDataResultID1, RawDataPeakName1

            try
            {
                // Get the original RawDataTable1 range to understand the structure
                var originalTable = WorksheetUtilities.GetNamedRange(sheet, "RawDataTable1");
                if (originalTable == null) return;

                // Calculate the new positions based on the original named ranges + column offset
                var originalPrepNums = WorksheetUtilities.GetNamedRange(sheet, "PrepNumsRawData1");
                var originalRawData = WorksheetUtilities.GetNamedRange(sheet, "RawData1");
                var originalResultID = WorksheetUtilities.GetNamedRange(sheet, "RawDataResultID1");
                var originalPeakName = WorksheetUtilities.GetNamedRange(sheet, "RawDataPeakName1");
                var originalAssayAsColumn = WorksheetUtilities.GetNamedRange(sheet, "AssayAs_Column");
                // var originalValidationValues = WorksheetUtilities.GetNamedRange(sheet, "ValidationTableValues");
                var originalRawDataStats = WorksheetUtilities.GetNamedRange(sheet, "RawDataStats1");
                var originalValidationStats = WorksheetUtilities.GetNamedRange(sheet, "ValidationStats");
                var originalRawDataConfidence = WorksheetUtilities.GetNamedRange(sheet, "RawDataConfidence1");
                var originalValidationConfidence = WorksheetUtilities.GetNamedRange(sheet, "ValidationConfidence");

                if (originalPrepNums != null)
                {
                    CreateShiftedNamedRange(sheet, originalPrepNums, $"PrepNumsRawData{sampleNum}", 0, colOffset);
                }

                if (originalRawData != null)
                {
                    CreateShiftedNamedRange(sheet, originalRawData, $"RawData{sampleNum}", 0, colOffset);
                }

                if (originalResultID != null)
                {
                    CreateShiftedNamedRange(sheet, originalResultID, $"RawDataResultID{sampleNum}", 0, colOffset);
                }

                if (originalPeakName != null)
                {
                    CreateShiftedNamedRange(sheet, originalPeakName, $"RawDataPeakName{sampleNum}", 0, colOffset);
                }
                // ADD THIS: Create AssayAs_Column2, AssayAs_Column3, etc.
                if (originalAssayAsColumn != null)
                {
                    CreateShiftedNamedRange(sheet, originalAssayAsColumn, $"AssayAs_Column{sampleNum}", 0, colOffset);
                    Console.WriteLine($"✅ Created named range: AssayAs_Column{sampleNum}");
                    WorksheetUtilities.ReleaseComObject(originalAssayAsColumn);
                }
                else
                {
                    Console.WriteLine($"⚠️  Warning: AssayAs_Column not found, skipping AssayAs_Column{sampleNum}");
                }

                // ADD THIS: Create RawDataStats2, RawDataStats3, etc.
                if (originalRawDataStats != null)
                {
                    CreateShiftedNamedRange(sheet, originalRawDataStats, $"RawDataStats{sampleNum}", 0, colOffset);
                    Console.WriteLine($"✅ Created named range: RawDataStats{sampleNum}");
                    WorksheetUtilities.ReleaseComObject(originalRawDataStats);
                }
                else
                {
                    Console.WriteLine($"⚠️  Warning: RawDataStats1 not found, skipping RawDataStats{sampleNum}");
                }

                // Create ValidationStats2, ValidationStats3, etc. (using validation column offset)
                if (originalValidationStats != null)
                {
                    // Use consecutive offset for validation (no gaps)
                    int validationColOffset = sampleNum - 1;
                    CreateShiftedNamedRange(sheet, originalValidationStats, $"ValidationStats{sampleNum}", 0, validationColOffset);
                    Console.WriteLine($"✅ Created named range: ValidationStats{sampleNum}");
                    WorksheetUtilities.ReleaseComObject(originalValidationStats);
                }

                // Create RawDataConfidence2, RawDataConfidence3, etc.
                if (originalRawDataConfidence != null)
                {
                    CreateShiftedNamedRange(sheet, originalRawDataConfidence, $"RawDataConfidence{sampleNum}", 0, colOffset);
                    Console.WriteLine($"✅ Created named range: RawDataConfidence{sampleNum}");
                    WorksheetUtilities.ReleaseComObject(originalRawDataConfidence);
                }
                else
                {
                    Console.WriteLine($"⚠️  Warning: RawDataConfidence1 not found, skipping RawDataConfidence{sampleNum}");
                }

                // Create ValidationConfidence2, ValidationConfidence3, etc. (using validation column offset)
                if (originalValidationConfidence != null)
                {
                    // Use consecutive offset for validation (no gaps)
                    int validationColOffset = sampleNum - 1;
                    CreateShiftedNamedRange(sheet, originalValidationConfidence, $"ValidationConfidence{sampleNum}", 0, validationColOffset);
                    Console.WriteLine($"✅ Created named range: ValidationConfidence{sampleNum}");
                    WorksheetUtilities.ReleaseComObject(originalValidationConfidence);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not create individual named ranges for table {sampleNum}: {ex.Message}");
            }
        }

        private static void CreateShiftedNamedRange(Worksheet sheet, ExcelRange originalRange, string newRangeName, int rowOffset, int colOffset)
        {
            try
            {
                // Calculate new range position
                int newStartRow = originalRange.Row + rowOffset;
                int newStartCol = originalRange.Column + colOffset;
                int newEndRow = newStartRow + originalRange.Rows.Count - 1;
                int newEndCol = newStartCol + originalRange.Columns.Count - 1;

                // Create the new range reference
                var startCell = sheet.Cells[newStartRow, newStartCol] as ExcelRange;
                var endCell = sheet.Cells[newEndRow, newEndCol] as ExcelRange;

                if (startCell != null && endCell != null)
                {
                    string refersToLocal = $"='{sheet.Name}'!" +
                                         startCell.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing) + ":" +
                                         endCell.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

                    sheet.Names.Add(newRangeName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                   Type.Missing, Type.Missing, refersToLocal, Type.Missing, Type.Missing, Type.Missing);

                    Console.WriteLine($"✅ Created named range: {newRangeName}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error creating named range {newRangeName}: {ex.Message}");
            }
        }

        private static void SetSingleCellValue(Worksheet sheet, string namedRange, string value)
        {
            if (WorksheetUtilities.NamedRangeExist(sheet, namedRange))
            {
                WorksheetUtilities.SetNamedRangeValue(sheet, namedRange, value, 1, 1);
                Console.WriteLine($"✅ {namedRange} = '{value}'");
            }
        }

        private static void CopyColumnWidths(Worksheet sheet, string sourceTableName, string destTableName)
        {
            try
            {
                var sourceRange = WorksheetUtilities.GetNamedRange(sheet, sourceTableName);
                var destRange = WorksheetUtilities.GetNamedRange(sheet, destTableName);

                if (sourceRange != null && destRange != null)
                {
                    // Copy column widths
                    for (int col = 1; col <= sourceRange.Columns.Count; col++)
                    {
                        var sourceCol = sourceRange.Columns[col] as ExcelRange;
                        var destCol = destRange.Columns[col] as ExcelRange;

                        if (sourceCol != null && destCol != null)
                        {
                            destCol.ColumnWidth = sourceCol.ColumnWidth;
                        }
                    }
                    Console.WriteLine($"✅ Copied column widths for {destTableName}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not copy column widths for {destTableName}: {ex.Message}");
            }
        }

        private static void UpdateValidationFormulas(Worksheet sheet, int numReplicates, int assayNumSamples)
        {
            Console.WriteLine($"🔗 Linking validation tables to raw data tables and applying formulas...");

            try
            {
                // Loop through each sample to create the mapping AND formulas
                for (int sampleNum = 1; sampleNum <= assayNumSamples; sampleNum++)
                {
                    // Step 1: Identify the corresponding ranges
                    string validationRange = sampleNum == 1 ? "ValidationTableValues" : $"ValidationTableValues{sampleNum}";
                    string rawDataRange = sampleNum == 1 ? "RawData1" : $"RawData{sampleNum}";

                    Console.WriteLine($"📝 Mapping: {validationRange} ↔ {rawDataRange}");

                    var validationTableRange = WorksheetUtilities.GetNamedRange(sheet, validationRange);
                    if (validationTableRange != null)
                    {
                        // Step 2: Apply INDEX formulas with AssayPrecision to each row
                        for (int row = 1; row <= numReplicates; row++)
                        {
                            var cell = validationTableRange.Cells[row, 1] as ExcelRange;
                            if (cell != null)
                            {
                                // Create formula that links ValidationTableValues to RawData + uses AssayPrecision
                                string formula = $"=IF(INDEX({rawDataRange},{row},1)=\"\",\"\",FIXED(INDEX({rawDataRange},{row},1),AssayPrecision))";
                                cell.Formula = formula;

                                Console.WriteLine($"  ✅ Row {row}: {validationRange} → {rawDataRange} (with precision)");
                            }
                        }

                        WorksheetUtilities.ReleaseComObject(validationTableRange);
                        Console.WriteLine($"✅ Completed mapping and formulas for {validationRange}");
                    }
                    else
                    {
                        Console.WriteLine($"❌ Warning: {validationRange} not found");
                    }
                }

                Console.WriteLine($"🎉 All validation tables linked to raw data with precision formulas!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error updating validation formulas: {ex.Message}");
            }
        }

        private static void UpdateValidationStatsFormulas(Worksheet sheet, int assayNumSamples)
        {
            Console.WriteLine($"🔗 Linking validation statistics to raw data statistics...");

            try
            {
                // Loop through each sample to map statistics
                for (int sampleNum = 1; sampleNum <= assayNumSamples; sampleNum++)
                {
                    // Identify corresponding ranges
                    string validationStatsRange = sampleNum == 1 ? "ValidationStats" : $"ValidationStats{sampleNum}";
                    string rawDataStatsRange = sampleNum == 1 ? "RawDataStats1" : $"RawDataStats{sampleNum}";

                    Console.WriteLine($"📊 Mapping: {validationStatsRange} ↔ {rawDataStatsRange}");

                    var validationStatsTableRange = WorksheetUtilities.GetNamedRange(sheet, validationStatsRange);
                    if (validationStatsTableRange != null)
                    {
                        // Cell 1: Mean with AssayPrecision
                        var meanCell = validationStatsTableRange.Cells[1, 1] as ExcelRange;
                        if (meanCell != null)
                        {
                            string meanFormula = $"=IF(INDEX({rawDataStatsRange},1,1)=\"\",\"\",FIXED(INDEX({rawDataStatsRange},1,1),AssayPrecision))";
                            meanCell.Formula = meanFormula;
                            Console.WriteLine($"  ✅ Mean: {meanFormula}");
                        }

                        // Cell 2: Std Dev with StdDevPrecision
                        var stdDevCell = validationStatsTableRange.Cells[2, 1] as ExcelRange;
                        if (stdDevCell != null)
                        {
                            string stdDevFormula = $"=IF(INDEX({rawDataStatsRange},2,1)=\"\",\"\",FIXED(INDEX({rawDataStatsRange},2,1),StdDevPrecision))";
                            stdDevCell.Formula = stdDevFormula;
                            Console.WriteLine($"  ✅ Std Dev: {stdDevFormula}");
                        }

                        // Cell 3: RSD with RSDPrecision
                        var rsdCell = validationStatsTableRange.Cells[3, 1] as ExcelRange;
                        if (rsdCell != null)
                        {
                            string rsdFormula = $"=IF(INDEX({rawDataStatsRange},3,1)=\"\",\"\",FIXED(INDEX({rawDataStatsRange},3,1),RSDPrecision))";
                            rsdCell.Formula = rsdFormula;
                            Console.WriteLine($"  ✅ RSD: {rsdFormula}");
                        }

                        WorksheetUtilities.ReleaseComObject(validationStatsTableRange);
                        Console.WriteLine($"✅ Completed statistics mapping for {validationStatsRange}");
                    }
                    else
                    {
                        Console.WriteLine($"❌ Warning: {validationStatsRange} not found");
                    }
                }

                Console.WriteLine($"🎉 All validation statistics linked to raw data with individual precision controls!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error updating validation statistics formulas: {ex.Message}");
            }
        }

        private static void UpdateValidationConfidenceFormulas(Worksheet sheet, int assayNumSamples)
        {
            Console.WriteLine($"🔗 Linking validation confidence intervals to raw data confidence intervals...");

            try
            {
                // Loop through each sample to map confidence intervals
                for (int sampleNum = 1; sampleNum <= assayNumSamples; sampleNum++)
                {
                    // Identify corresponding ranges
                    string validationConfidenceRange = sampleNum == 1 ? "ValidationConfidence" : $"ValidationConfidence{sampleNum}";
                    string rawDataConfidenceRange = sampleNum == 1 ? "RawDataConfidence1" : $"RawDataConfidence{sampleNum}";

                    Console.WriteLine($"📊 Mapping: {validationConfidenceRange} ↔ {rawDataConfidenceRange}");

                    var validationConfidenceTableRange = WorksheetUtilities.GetNamedRange(sheet, validationConfidenceRange);
                    if (validationConfidenceTableRange != null)
                    {
                        // Cell 1: Lower 95% Confidence Interval with LowerPrecision
                        var lowerConfidenceCell = validationConfidenceTableRange.Cells[1, 1] as ExcelRange;
                        if (lowerConfidenceCell != null)
                        {
                            string lowerFormula = $"=IF(INDEX({rawDataConfidenceRange},1,1)=\"\",\"\",FIXED(INDEX({rawDataConfidenceRange},1,1),LowerPrecision))";
                            lowerConfidenceCell.Formula = lowerFormula;
                            Console.WriteLine($"  ✅ Lower 95%: {lowerFormula}");
                        }

                        // Cell 2: Upper 95% Confidence Interval with UpperPrecision
                        var upperConfidenceCell = validationConfidenceTableRange.Cells[2, 1] as ExcelRange;
                        if (upperConfidenceCell != null)
                        {
                            string upperFormula = $"=IF(INDEX({rawDataConfidenceRange},2,1)=\"\",\"\",FIXED(INDEX({rawDataConfidenceRange},2,1),UpperPrecision))";
                            upperConfidenceCell.Formula = upperFormula;
                            Console.WriteLine($"  ✅ Upper 95%: {upperFormula}");
                        }

                        WorksheetUtilities.ReleaseComObject(validationConfidenceTableRange);
                        Console.WriteLine($"✅ Completed confidence interval mapping for {validationConfidenceRange}");
                    }
                    else
                    {
                        Console.WriteLine($"❌ Warning: {validationConfidenceRange} not found");
                    }
                }

                Console.WriteLine($"🎉 All validation confidence intervals linked to raw data with specific precision controls!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error updating validation confidence formulas: {ex.Message}");
            }
        }

        //this is old
        private static void HandleImpurity(
                Worksheet sheet,
                int numReplicates,
                int impurityNumSamples,
                int numOfPeaks,
                string quantitationType,

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
            Console.WriteLine($"🔍 Setting Number_of_ImpSample_Preparation = {impurityNumSamples}");
            SetSingleCellValue(sheet, "Number_of_ImpSample_Preparation", impurityNumSamples.ToString());

            Console.WriteLine($"🔍 Setting Quantitation Type = '{quantitationType}'");

            SetSingleCellValue(sheet, "QuantitationType", quantitationType);
            // At the end of HandleImpurity method:
            SetImpurityAcceptanceCriteriaFromIndividualValues(
                sheet,
                impurityOperator1,
                impurityValue1,
                impurityValue2,

                impurityAutoOperator1,
                impurityAutoValue1,

                impurityOperator2,
                impurityValue3,

                impurityOperator4,
                impurityValue4,

                impurityAutoOperator2,
                impurityAutoValue2,

                impurityOperator5,
                impurityValue5);
            //SetImpurityAcceptanceCriteriaFromIndividualValues(sheet, impurityValue1);

            if (impurityNumSamples <= 0)
            {
                WorksheetUtilities.DeleteNamedRangeRows(sheet, "ImpurityAndImpuritySummary");
                WorksheetUtilities.DeleteNamedRange(sheet, "ImpurityAndImpuritySummary");
                return;
            }

            // EXISTING STEP 1: Expand or contract the tables based on the number of replicates
            if (numReplicates > DefaultNumReplicates)
            {
                int numRowsToInsert = numReplicates - DefaultNumReplicates;
                WorksheetUtilities.InsertRowsIntoNamedRange(numRowsToInsert, sheet, "ImpurityPrep1", true, XlDirection.xlDown, XlPasteType.xlPasteFormulas);
            }
            else if (numReplicates < DefaultNumReplicates)
            {
                int numRowsToRemove = DefaultNumReplicates - numReplicates;
                // There needs to be at least 2 rows in order to not corrupt the sheet's formulas
                if (DefaultNumReplicates - numRowsToRemove < MinNumReplicates) numRowsToRemove = DefaultNumReplicates - MinNumReplicates;
                WorksheetUtilities.DeleteRowsFromNamedRange(numRowsToRemove, sheet, "ImpurityPrep1", XlDirection.xlDown);
            }

            // EXISTING STEP 2: Re-number the preps in the rawdata and validation results tables
            if (numReplicates < MinNumReplicates) numReplicates = MinNumReplicates;
            List<string> prepNumbers = new List<string>(0);
            for (int i = 1; i <= numReplicates; i++) prepNumbers.Add(i.ToString());
            WorksheetUtilities.SetNamedRangeValues(sheet, "ImpurityPrep1", prepNumbers);

            // EXISTING STEP 3: Column Copying based on numOfPeaks (SECOND - after row expansion)
            if (numOfPeaks > 1)
            {
                Console.WriteLine($"Creating {numOfPeaks - 1} additional Result ID columns...");

                for (int peakNum = 2; peakNum <= numOfPeaks; peakNum++)
                {
                    int colOffset = peakNum - 1; // Peak 2 = 1 col right (E), Peak 3 = 2 cols right (F)

                    Console.WriteLine($"Creating Result ID Peak {peakNum} column at column offset {colOffset}...");

                    try
                    {
                        // Step 3.1: Copy headers first (Result ID header)

                        string destHeaderName = $"ImpurityImpurityHeader1{peakNum}";
                        Console.WriteLine($"Copying headers: ImpurityImpurityHeader1 → {destHeaderName}");
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ImpurityImpurityHeader1", destHeaderName, 1, colOffset + 1, XlPasteType.xlPasteAll);

                        string destDataName = $"ImpurityData{peakNum}";
                        Console.WriteLine($"Copying data: ImpurityData1 → {destDataName}");
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ImpurityData1", destDataName, 1, colOffset + 1, XlPasteType.xlPasteAll);

                        // Step 3.3: Copy statistics (ImpurityStats1 → ImpurityStats2, etc.)
                        string destStatsName = $"ImpurityStats{peakNum}";
                        Console.WriteLine($"Copying statistics: ImpurityStats1 → {destStatsName}");
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ImpurityStats1", destStatsName, 1, colOffset + 1, XlPasteType.xlPasteAll);

                        Console.WriteLine($"✅ Successfully created Impurity {peakNum} statistics column");

                        // Copy column width
                        var sourceColumn = sheet.Columns[4] as ExcelRange; // Column D (Result ID)
                        var destColumn = sheet.Columns[4 + colOffset + 1] as ExcelRange; // Target column
                        if (sourceColumn != null && destColumn != null) destColumn.ColumnWidth = sourceColumn.ColumnWidth;

                        var headerRange = WorksheetUtilities.GetNamedRange(sheet, destHeaderName);
                        if (headerRange != null)
                        {
                            headerRange.Value2 = $"Impurity {peakNum}";
                            Console.WriteLine($"✅ Updated header to 'Impurity {peakNum}'");
                        }

                        Console.WriteLine($"✅ Successfully created Impurity {peakNum} column");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"❌ Error creating Impurity {peakNum} column: {ex.Message}");
                    }
                }

                // RESIZE the ImpurityMainTable1 named range ONCE after all columns are copied
                int additionalColumns = numOfPeaks - 1; // How many extra columns we added
                try
                {
                    WorksheetUtilities.ResizeNamedRange(sheet, "ImpurityMainTable1", 0, additionalColumns);
                    Console.WriteLine($"✅ Updated ImpurityMainTable1 to include {numOfPeaks} Impurity Columns");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Error resizing ImpurityMainTable1: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("No additional Impurity columns needed (numOfPeaks = 1)");
            }

            Console.WriteLine("✅ Impurity column copying complete!");

            // UPDATED STEP 4: VERTICAL Table Copying based on impurityNumSamples (THIRD - after row + column expansion)
            if (impurityNumSamples > 1)
            {
                Console.WriteLine($"Creating {impurityNumSamples - 1} additional impurity tables VERTICALLY...");

                // STEP 1: Calculate space requirements
                var originalTable = WorksheetUtilities.GetNamedRange(sheet, "ImpurityMainTable1");
                if (originalTable == null)
                {
                    Console.WriteLine("❌ Error: ImpurityMainTable1 not found, skipping vertical copying");
                    return;
                }

                int tableHeight = originalTable.Rows.Count;
                int spacing = 2; // 2 rows between tables as requested
                int additionalTables = impurityNumSamples - 1; // How many new tables to create
                int totalRowsNeeded = additionalTables * (tableHeight + spacing);

                Console.WriteLine($"📊 Space Calculation:");
                Console.WriteLine($"   Table height: {tableHeight} rows");
                Console.WriteLine($"   Tables to create: {additionalTables}");
                Console.WriteLine($"   Spacing between tables: {spacing} rows");
                Console.WriteLine($"   Total rows needed: {totalRowsNeeded}");

                // STEP 2: Find insertion point (after ImpurityMainTable1)
                int lastRowOfTable1 = originalTable.Row + originalTable.Rows.Count - 1;
                int insertionPoint = lastRowOfTable1 + spacing; // Add spacing after table 1

                Console.WriteLine($"🔧 Original table ends at row {lastRowOfTable1}");
                Console.WriteLine($"🔧 Will insert {totalRowsNeeded} rows starting at row {insertionPoint}...");

                // STEP 3: INSERT ALL REQUIRED ROWS AT ONCE (pushes Summary + WaterContent down)
                try
                {
                    var insertRange = sheet.Range[$"{insertionPoint}:{insertionPoint + totalRowsNeeded - 1}"];
                    insertRange.EntireRow.Insert(XlInsertShiftDirection.xlShiftDown);
                    Console.WriteLine($"✅ Successfully inserted {totalRowsNeeded} rows, all sections pushed down!");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Failed to insert rows: {ex.Message}");
                    return;
                }

                // STEP 4: Copy tables to the new empty space sequentially
                for (int impNum = 2; impNum <= impurityNumSamples; impNum++)
                {
                    try
                    {
                        // Calculate target row for this table
                        int tableIndex = impNum - 2; // Table 2 = index 0, Table 3 = index 1, etc.
                        int targetRow = insertionPoint + (tableIndex * (tableHeight + spacing));

                        Console.WriteLine($"📋 Copying Impurity Table {impNum} to row {targetRow}...");

                        // Copy the entire table using absolute positioning
                        string destTableName = $"ImpurityMainTable{impNum}";
                        WorksheetUtilities.CopyNamedRangeToNewLocationWithNewNamedRange(
                            sheet,
                            "ImpurityMainTable1",
                            destTableName,
                            targetRow,                    // Absolute row position
                            originalTable.Column,         // Same column as original
                            XlPasteType.xlPasteAll
                        );

                        // Copy column widths for the entire table
                        CopyColumnWidths(sheet, "ImpurityMainTable1", destTableName);

                        // Create individual named ranges for the copied table components
                        int rowOffset = targetRow - originalTable.Row; // Calculate offset from original
                        CreateIndividualNamedRangesForImpurityTable(sheet, impNum, 0, rowOffset, numOfPeaks);

                        Console.WriteLine($"✅ Successfully created Impurity Table {impNum} at row {targetRow}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"❌ Error creating Impurity Table {impNum}: {ex.Message}");
                    }
                }

                Console.WriteLine($"🎯 Vertical copying complete: {impurityNumSamples} impurity tables created!");
            }
            else
            {
                Console.WriteLine("No additional Impurity tables needed (impurityNumSamples = 1)");
            }

            // Step 5C: Row Expansion for Summary Impurity rows based on numOfPeaks
            if (numOfPeaks > 1) // Default is 1 impurity row, expand if more peaks
            {
                int numRowsToInsert = numOfPeaks - 1;
                Console.WriteLine($"Expanding Summary Impurity rows by {numRowsToInsert} rows for {numOfPeaks} peaks...");

                try
                {
                    // Insert rows into the impurity section - all copied rows will say "Impurity"
                    WorksheetUtilities.InsertRowsIntoNamedRange(numRowsToInsert, sheet, "ImpuritySummaryRow", true, XlDirection.xlDown, XlPasteType.xlPasteAll);
                    Console.WriteLine($"✅ Successfully created {numOfPeaks} impurity rows in summary table");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Error expanding impurity rows: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("No additional impurity rows needed (numOfPeaks = 1)");
            }
            // Step 5D: Update Impurity Labels (Impurity 1, Impurity 2, Impurity 3)
            if (numOfPeaks > 1)
            {
                Console.WriteLine($"Updating impurity labels for {numOfPeaks} peaks...");

                var summaryTable = WorksheetUtilities.GetNamedRange(sheet, "ImpuritySummaryTable");
                if (summaryTable != null)
                {
                    int headerRow = summaryTable.Row + 1;
                    int startDataRow = headerRow + 2;

                    for (int peakNum = 1; peakNum <= numOfPeaks; peakNum++)
                    {
                        int dataRow = startDataRow + (peakNum - 1);
                        var labelCell = sheet.Cells[dataRow, 2] as ExcelRange; // Column B
                        if (labelCell != null)
                        {
                            labelCell.Value2 = $"Impurity {peakNum}";
                            Console.WriteLine($"✅ Updated row {dataRow} to 'Impurity {peakNum}'");
                        }
                    }
                }
            }

            // Step 5A: Column Expansion with Correct Statistics Preservation
            const int TemplatePrepColumns = 2; // Template starts with 2 Prep columns

            if (numReplicates > TemplatePrepColumns)
            {
                Console.WriteLine($"Creating {numReplicates - TemplatePrepColumns} additional Prep columns using direct cell copying...");

                // Get the summary table to find actual boundaries
                var summaryTable = WorksheetUtilities.GetNamedRange(sheet, "ImpuritySummaryTable");
                if (summaryTable != null)
                {
                    int headerRow = summaryTable.Row + 1;
                    int columnHeaderRow = headerRow + 1;       // Row 109 (column headers) ← ADD THIS
                    int startDataRow = headerRow + 2;
                    int totalImpurityRows = numOfPeaks;
                    int endDataRow = startDataRow + totalImpurityRows - 1;

                    // Template columns: Component=A, Prep-1=C, Prep-2=D, Mean=E, StdDev=F, RSD=G
                    int componentColumn = 1; // Column A
                    int prep1Column = 3;     // Column C
                    int prep2Column = 4;     // Column D
                    int originalMeanColumn = 5;   // Column E
                    int originalStdDevColumn = 6; // Column F
                    int originalRSDColumn = 7;    // Column G
                    int originalLCI = 8;
                    int originalUCI = 9;

                    Console.WriteLine($"DEBUG: Table structure - HeaderRow={headerRow}");
                    Console.WriteLine($"DEBUG: DataRows={startDataRow} to {endDataRow} (Total impurity rows: {totalImpurityRows})");
                    Console.WriteLine($"DEBUG: Template columns - Prep-1=Column {prep1Column}, Prep-2=Column {prep2Column}");
                    Console.WriteLine($"DEBUG: Original statistics - Mean={originalMeanColumn}, StdDev={originalStdDevColumn}, RSD={originalRSDColumn}");

                    // STEP 1: Calculate final positions for statistics columns
                    // They should be AFTER all Prep columns (Prep-1, Prep-2, Prep-3, ..., Prep-numReplicates)
                    int finalMeanColumn = prep1Column + numReplicates; // After all numReplicates Prep columns
                    int finalStdDevColumn = finalMeanColumn + 1;
                    int finalRSDColumn = finalStdDevColumn + 1;
                    int finaloriginalLCI = finalRSDColumn + 1;
                    int finaloriginalUCI = finaloriginalLCI + 1;

                    Console.WriteLine($"Moving statistics columns to final positions...");
                    Console.WriteLine($"Final positions - Mean={finalMeanColumn}({GetColumnLetter(finalMeanColumn)}), StdDev={finalStdDevColumn}({GetColumnLetter(finalStdDevColumn)}), RSD={finalRSDColumn}({GetColumnLetter(finalRSDColumn)})");

                    // STEP 2: Move statistics columns to the end (in REVERSE order to avoid overwriting)
                    // Move %RSD first, then StdDev, then Mean
                    MoveColumnRange(sheet, columnHeaderRow, endDataRow, originalRSDColumn, finalRSDColumn, "%RSD");
                    MoveColumnRange(sheet, columnHeaderRow, endDataRow, originalStdDevColumn, finalStdDevColumn, "Std Dev");
                    MoveColumnRange(sheet, columnHeaderRow, endDataRow, originalMeanColumn, finalMeanColumn, "Mean(%)");
                    MoveColumnRange(sheet, columnHeaderRow, endDataRow, originalLCI, finaloriginalLCI, "Lower 95% Confidence Interval");
                    MoveColumnRange(sheet, columnHeaderRow, endDataRow, originalUCI, finaloriginalUCI, "Upper 95% Confidence Interval");

                    // STEP 3: Now create the Prep columns (starting from where Mean used to be)
                    for (int prepNum = 3; prepNum <= numReplicates; prepNum++)
                    {
                        int targetColumn = prep2Column + (prepNum - 2); // Prep-3=Column E, Prep-4=Column F, etc.

                        Console.WriteLine($"Creating Prep-{prepNum} in column {targetColumn} ({GetColumnLetter(targetColumn)})...");

                        try
                        {
                            // Copy header cell (from Prep-2 header to new column)
                            var sourceHeaderCell = sheet.Cells[columnHeaderRow, prep2Column] as ExcelRange;
                            var targetHeaderCell = sheet.Cells[columnHeaderRow, targetColumn] as ExcelRange;
                            if (sourceHeaderCell != null && targetHeaderCell != null)
                            {
                                sourceHeaderCell.Copy(Type.Missing);
                                targetHeaderCell.PasteSpecial(XlPasteType.xlPasteAll, XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                                targetHeaderCell.Value2 = $"Prep-{prepNum}";
                                Console.WriteLine($"✅ Created header 'Prep-{prepNum}' in {GetColumnLetter(targetColumn)}{headerRow}");
                            }

                            // Copy all data cells in this column
                            int cellsCopied = 0;
                            for (int row = startDataRow; row <= endDataRow; row++)
                            {
                                var sourceDataCell = sheet.Cells[row, prep2Column] as ExcelRange;
                                var targetDataCell = sheet.Cells[row, targetColumn] as ExcelRange;
                                if (sourceDataCell != null && targetDataCell != null)
                                {
                                    bool success = false;
                                    int attempts = 0;
                                    int maxAttempts = 3;

                                    while (!success && attempts < maxAttempts)
                                    {
                                        try
                                        {
                                            attempts++;
                                            _app.CutCopyMode = 0; // Clear clipboard
                                            sourceDataCell.Copy(Type.Missing);
                                            System.Threading.Thread.Sleep(10); // Small delay
                                            targetDataCell.PasteSpecial(XlPasteType.xlPasteAll, XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                                            _app.CutCopyMode = 0; // Clear clipboard
                                            success = true;
                                            cellsCopied++;
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine($"⚠️ Attempt {attempts} failed for row {row}: {ex.Message}");
                                            if (attempts < maxAttempts)
                                            {
                                                System.Threading.Thread.Sleep(100); // Longer delay before retry
                                            }
                                            else
                                            {
                                                Console.WriteLine($"❌ Failed to copy cell after {maxAttempts} attempts");
                                            }
                                        }
                                    }
                                }
                            }

                            // Copy column width
                            var sourceColumn = sheet.Columns[prep2Column] as ExcelRange;
                            var targetColumn_Range = sheet.Columns[targetColumn] as ExcelRange;
                            if (sourceColumn != null && targetColumn_Range != null)
                                targetColumn_Range.ColumnWidth = sourceColumn.ColumnWidth;

                            Console.WriteLine($"✅ Successfully created Prep-{prepNum} column");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"❌ Error creating Prep-{prepNum} column: {ex.Message}");
                        }
                    }

                    // Exclude both Column A AND the header row from formatting
                    int startFormattingColumn = componentColumn + 1; // Start from column B (exclude column A)
                    int startFormattingRow = headerRow + 1; // Start from row after header (exclude header row)

                    // Get the header row from named range to make it dynamic
                    var headerRowRange = WorksheetUtilities.GetNamedRange(sheet, "ImpuritySummaryHeaderRow");
                    if (headerRowRange != null)
                    {
                        startFormattingRow = headerRowRange.Row + 1; // Start formatting from row after the header
                        Console.WriteLine($"📍 Excluding header row {headerRowRange.Row} from table formatting");
                    }

                    // Apply formatting to exclude BOTH column A AND header row
                    var tableRange = sheet.Range[
                        sheet.Cells[startFormattingRow, startFormattingColumn],
                        sheet.Cells[endDataRow, finaloriginalUCI]
                    ];
                    tableRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                    tableRange.Borders.Weight = XlBorderWeight.xlThin;

                    Console.WriteLine($"✅ Applied table formatting to range {GetColumnLetter(startFormattingColumn)}{startFormattingRow}:{GetColumnLetter(finaloriginalUCI)}{endDataRow}");

                    // Update the summary table range to include new columns and rows
                    int originalImpurityRows = 1; // Template had 1 impurity row
                    int currentImpurityRows = numOfPeaks; // Now we have 3 impurity rows
                    int rowChange = currentImpurityRows - originalImpurityRows; // +2 rows
                    int totalColumns = numReplicates + 6; // Component + numReplicates Prep columns + 3 statistics columns
                    int colChange = totalColumns - TemplatePrepColumns - 1;

                    try
                    {
                        WorksheetUtilities.ResizeNamedRange(sheet, "ImpuritySummaryTable", rowChange, colChange);
                        Console.WriteLine($"✅ Updated ImpuritySummaryTable to include {numReplicates} Prep columns + 3 statistics columns");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"❌ Error resizing ImpuritySummaryTable: {ex.Message}");
                    }
                }
                else
                {
                    Console.WriteLine("❌ ERROR: ImpuritySummaryTable named range not found!");
                }
            }
            else
            {
                Console.WriteLine($"No additional Prep columns needed (numReplicates = {numReplicates} <= template columns = {TemplatePrepColumns})");
            }

            Console.WriteLine("✅ Summary table Prep column copying complete!");

            //step 6 Copying summary table

            if (impurityNumSamples > 1)
            {
                Console.WriteLine($"Creating {impurityNumSamples - 1} additional summary tables using simple row insertion...");

                var summaryTable = WorksheetUtilities.GetNamedRange(sheet, "ImpuritySummaryTable");
                if (summaryTable != null)
                {
                    // STEP 1: Calculate space needed
                    int actualTableHeight = summaryTable.Rows.Count; // Just use actual current height
                    int spacing = 2; // 2 rows gap between tables for safety
                    int additionalTables = impurityNumSamples - 1;
                    int totalRowsToInsert = (actualTableHeight + spacing) * additionalTables;

                    Console.WriteLine($"📊 Simple Calculation:");
                    Console.WriteLine($"   Current table height: {actualTableHeight} rows");
                    Console.WriteLine($"   Tables to create: {additionalTables}");
                    Console.WriteLine($"   Total rows needed: {totalRowsToInsert}");

                    // STEP 2: Insert empty rows AFTER the summary table (with proper gap)
                    int lastRowOfSummary = summaryTable.Row + summaryTable.Rows.Count; // Last row of the actual table
                    int insertionPoint = lastRowOfSummary + spacing; // Add spacing gap BEFORE insertion

                    Console.WriteLine($"🔧 Summary table ends at row {lastRowOfSummary}");
                    Console.WriteLine($"🔧 Will insert {totalRowsToInsert} rows starting at row {insertionPoint} (after {spacing}-row gap)...");

                    try
                    {
                        // Insert rows using Excel's native method - AFTER the table with gap
                        ExcelRange insertPoint = sheet.Range[$"{insertionPoint}:{insertionPoint + totalRowsToInsert - 1}"] as ExcelRange;
                        if (insertPoint != null)
                        {
                            insertPoint.EntireRow.Insert(XlInsertShiftDirection.xlShiftDown);
                            Console.WriteLine($"✅ Successfully inserted rows starting at {insertionPoint}, Water Content pushed down!");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"❌ Row insertion failed: {ex.Message}");
                        return;
                    }

                    // STEP 3: Copy tables using Excel's native copy/paste - AFTER the insertion point
                    for (int impNum = 2; impNum <= impurityNumSamples; impNum++)
                    {
                        try
                        {
                            // Calculate target row - starts after insertion point
                            int targetRow = insertionPoint + ((impNum - 2) * (actualTableHeight + spacing));
                            Console.WriteLine($"📋 Copying Summary Table {impNum} to row {targetRow}...");

                            // Turn off Excel alerts and interactions
                            if (_app != null)
                            {
                                _app.DisplayAlerts = false;
                                _app.Interactive = false;
                            }

                            summaryTable.Copy();

                            // Paste to target location
                            ExcelRange pasteTarget = sheet.Cells[targetRow, summaryTable.Column] as ExcelRange;
                            if (pasteTarget != null)
                            {
                                pasteTarget.PasteSpecial(XlPasteType.xlPasteAll, XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                                // Create named range for the copied table
                                string destTableName = $"ImpuritySummaryTable{impNum}";
                                int endRow = targetRow + actualTableHeight - 1;
                                int endCol = summaryTable.Column + summaryTable.Columns.Count - 1;

                                ExcelRange startCell = sheet.Cells[targetRow, summaryTable.Column] as ExcelRange;
                                ExcelRange endCell = sheet.Cells[endRow, endCol] as ExcelRange;

                                if (startCell != null && endCell != null)
                                {
                                    string refersTo = $"='{sheet.Name}'!" +
                                                    startCell.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing) +
                                                    ":" +
                                                    endCell.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

                                    sheet.Names.Add(destTableName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, refersTo, Type.Missing, Type.Missing, Type.Missing);

                                    Console.WriteLine($"✅ Successfully created {destTableName}");
                                }
                            }

                            // Clear clipboard
                            if (_app != null)
                            {
                                _app.CutCopyMode = 0;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"❌ Error creating Summary Table {impNum}: {ex.Message}");
                        }
                    }

                    // Restore Excel settings
                    if (_app != null)
                    {
                        _app.DisplayAlerts = true;
                        _app.Interactive = false; // Keep this false for automation
                    }

                    Console.WriteLine($"🎯 Simple copy complete: {impurityNumSamples} summary tables created!");
                }
                else
                {
                    Console.WriteLine("❌ ERROR: ImpuritySummaryTable named range not found!");
                }
            }
            else
            {
                Console.WriteLine("ℹ️  Only 1 impurity table needed, skipping summary table copying");
            }

            try
            {
                UpdateImpuritySummaryTableFormulasUsingNamedRanges(sheet, impurityNumSamples, numOfPeaks);
            }
            catch (Exception)
            {
                Logger.LogMessage("Scroll of sheet failed in SampleRepeatability.UpdateImpuritySummaryTableFormulas!", Level.Error);
            }

            try
            {
                UpdateImpurityMeanFormulas(sheet, numReplicates, impurityNumSamples, impurityNumSamples); // numTables = impurityNumSamples for now
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not update mean formulas: {ex.Message}");
            }

            // Add this after the Mean formulas call:
            try
            {
                UpdateImpurityStdDevFormulas(sheet, numReplicates, impurityNumSamples, impurityNumSamples);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not update std dev formulas: {ex.Message}");
            }

            // NEW: Update RSD formulas specifically
            try
            {
                UpdateImpurityRSDFormulas(sheet, numReplicates, impurityNumSamples, impurityNumSamples);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not update RSD formulas: {ex.Message}");
            }

            // NEW: Update Lower Confidence Interval formulas specifically
            try
            {
                UpdateImpurityLowerConfidenceFormulas(sheet, numReplicates, impurityNumSamples, impurityNumSamples);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not update lower confidence formulas: {ex.Message}");
            }

            // NEW: Update Upper Confidence Interval formulas specifically
            try
            {
                UpdateImpurityUpperConfidenceFormulas(sheet, numReplicates, impurityNumSamples, impurityNumSamples);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not update upper confidence formulas: {ex.Message}");
            }

            Console.WriteLine("✅ Impurity section processing complete!");
        }

        private static void SetImpurityAcceptanceCriteriaFromIndividualValues(
            Worksheet sheet,
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
            try
            {
                var range = WorksheetUtilities.GetNamedRange(sheet, "ImpurityAcceptanceCriteriaRange");
                if (range != null)
                {
                    // Row 1: Assign individual variables to specific positions
                    (range.Cells[1, 4] as ExcelRange).Value2 = impurityOperator1;
                    (range.Cells[1, 5] as ExcelRange).Value2 = impurityValue1;
                    (range.Cells[1, 7] as ExcelRange).Value2 = impurityValue2;

                    // Row 2: Assign individual variables to specific positions
                    (range.Cells[2, 1] as ExcelRange).Value2 = impurityAutoOperator1;
                    (range.Cells[2, 2] as ExcelRange).Value2 = impurityAutoValue1;
                    (range.Cells[2, 4] as ExcelRange).Value2 = impurityOperator2;
                    (range.Cells[2, 5] as ExcelRange).Value2 = impurityValue3;
                    (range.Cells[2, 6] as ExcelRange).Value2 = impurityOperator4;
                    (range.Cells[2, 7] as ExcelRange).Value2 = impurityValue4;

                    //row3
                    (range.Cells[3, 1] as ExcelRange).Value2 = impurityAutoOperator2;
                    (range.Cells[3, 2] as ExcelRange).Value2 = impurityAutoValue2;
                    (range.Cells[3, 6] as ExcelRange).Value2 = impurityOperator5;
                    (range.Cells[3, 7] as ExcelRange).Value2 = impurityValue5;

                    Console.WriteLine("✅ Set all individual impurity criteria values");
                    WorksheetUtilities.ReleaseComObject(range);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error setting individual values: {ex.Message}");
            }
        }

        private static void CreateIndividualNamedRangesForImpurityTable(Worksheet sheet, int tableNum, int colOffset, int rowOffset, int numOfPeaks)
        {
            // Create named ranges for the individual components of the copied impurity table
            // Based on the original pattern: ImpurityData1, ImpurityData2, ImpurityData3, etc.
            // DYNAMIC: Works for any number of peaks

            try
            {
                // Get the original ImpurityMainTable1 range to understand the structure
                var originalTable = WorksheetUtilities.GetNamedRange(sheet, "ImpurityMainTable1");
                if (originalTable == null) return;

                Console.WriteLine($"Creating named ranges for Table {tableNum} with {numOfPeaks} peaks (rowOffset={rowOffset}, colOffset={colOffset})...");

                // STEP 1: Copy MeanPrecision cell for this table (only for tables > 1)
                if (tableNum > 1)
                {
                    try
                    {
                        string sourcePrecisionName = "MeanPrecision"; // Original precision cell
                        string destPrecisionName = $"MeanPrecision{tableNum}"; // New precision cell for this table

                        Console.WriteLine($"   📋 Creating precision cell: {sourcePrecisionName} → {destPrecisionName}");

                        // Check if source precision cell exists
                        if (WorksheetUtilities.NamedRangeExist(sheet, sourcePrecisionName))
                        {
                            var originalPrecision = WorksheetUtilities.GetNamedRange(sheet, sourcePrecisionName);
                            if (originalPrecision != null)
                            {
                                // Create shifted named range for the precision cell
                                CreateShiftedNamedRange(sheet, originalPrecision, destPrecisionName, rowOffset, colOffset);
                                Console.WriteLine($"   ✅ Created {destPrecisionName} (row+{rowOffset}, col+{colOffset})");

                                WorksheetUtilities.ReleaseComObject(originalPrecision);
                            }
                        }
                        else
                        {
                            Console.WriteLine($"   ⚠️  Warning: {sourcePrecisionName} not found, skipping precision copy");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"   ❌ Error copying MeanPrecision for table {tableNum}: {ex.Message}");
                    }
                }

                // DYNAMIC: Loop through all peaks based on numOfPeaks
                for (int peakNum = 1; peakNum <= numOfPeaks; peakNum++)
                {
                    // Get the original named ranges for this peak
                    string originalDataName = $"ImpurityData{peakNum}";
                    string originalStatsName = $"ImpurityStats{peakNum}";

                    var originalData = WorksheetUtilities.GetNamedRange(sheet, originalDataName);
                    var originalStats = WorksheetUtilities.GetNamedRange(sheet, originalStatsName);

                    // Create shifted named ranges for the new table
                    if (originalData != null)
                    {
                        string newDataName = $"ImpurityData{peakNum}_T{tableNum}";
                        CreateShiftedNamedRange(sheet, originalData, newDataName, rowOffset, colOffset);
                        Console.WriteLine($"✅ Created {newDataName} (row+{rowOffset}, col+{colOffset})");
                    }
                    else
                    {
                        Console.WriteLine($"⚠️  Warning: {originalDataName} not found, skipping data range for peak {peakNum}");
                    }

                    if (originalStats != null)
                    {
                        string newStatsName = $"ImpurityStats{peakNum}_T{tableNum}";
                        CreateShiftedNamedRange(sheet, originalStats, newStatsName, rowOffset, colOffset);
                        Console.WriteLine($"✅ Created {newStatsName} (row+{rowOffset}, col+{colOffset})");
                    }
                    else
                    {
                        Console.WriteLine($"⚠️  Warning: {originalStatsName} not found, skipping stats range for peak {peakNum}");
                    }
                }

                Console.WriteLine($"✅ Created impurity named ranges for Table {tableNum} ({numOfPeaks} peaks)");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error creating individual named ranges for impurity table {tableNum}: {ex.Message}");
            }
        }

        private static string GetColumnLetter(int columnNumber)
        {
            string columnLetter = "";
            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnLetter = Convert.ToChar('A' + modulo) + columnLetter;
                columnNumber = (columnNumber - modulo) / 26;
            }
            return columnLetter;
        }

        // Helper method to move a column range
        private static void MoveColumnRange(Worksheet sheet, int startRow, int endRow, int sourceColumn, int targetColumn, string columnName)
        {
            try
            {
                Console.WriteLine($"Moving {columnName} from column {sourceColumn} to column {targetColumn}...");

                // Copy the entire column range
                var sourceRange = sheet.Range[
                    sheet.Cells[startRow, sourceColumn],
                    sheet.Cells[endRow, sourceColumn]
                ];

                var targetRange = sheet.Range[
                    sheet.Cells[startRow, targetColumn],
                    sheet.Cells[endRow, targetColumn]
                ];

                sourceRange.Copy(Type.Missing);
                targetRange.PasteSpecial(XlPasteType.xlPasteAll, XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                // Clear the original location
                sourceRange.ClearContents();

                // Copy column width
                var sourceCol = sheet.Columns[sourceColumn] as ExcelRange;
                var targetCol = sheet.Columns[targetColumn] as ExcelRange;
                if (sourceCol != null && targetCol != null)
                    targetCol.ColumnWidth = sourceCol.ColumnWidth;

                Console.WriteLine($"✅ Successfully moved {columnName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error moving {columnName}: {ex.Message}");
            }
        }

        private static void UpdateImpuritySummaryTableFormulasUsingNamedRanges(Worksheet sheet, int impurityNumSamples, int numOfPeaks)
        {
            Console.WriteLine($"Fixing formulas using named ranges for {impurityNumSamples} tables and {numOfPeaks} peaks...");

            for (int tableNum = 1; tableNum <= impurityNumSamples; tableNum++)
            {
                string summaryTableName = $"ImpuritySummaryTable{(tableNum == 1 ? "" : tableNum.ToString())}";
                var summaryTable = WorksheetUtilities.GetNamedRange(sheet, summaryTableName);

                if (summaryTable != null)
                {
                    Console.WriteLine($"\n--- Fixing Table {tableNum} ({summaryTableName}) ---");

                    // Calculate precision reference (always Column C, first row of summary table)
                    int precisionRow = summaryTable.Row;
                    string precisionRef = $"C{precisionRow}";
                    Console.WriteLine($"📍 Precision reference: {precisionRef}");

                    // Calculate data rows in summary table
                    int headerRow = summaryTable.Row + 1;
                    int startDataRow = headerRow + 2; // Skip precision + title + column headers

                    for (int peakNum = 1; peakNum <= numOfPeaks; peakNum++)
                    {
                        int summaryDataRow = startDataRow + (peakNum - 1);

                        // Get the named range for this peak's data
                        string dataRangeName = $"ImpurityData{peakNum}";
                        if (tableNum > 1)
                        {
                            dataRangeName += $"_T{tableNum}";
                        }

                        var impurityDataRange = WorksheetUtilities.GetNamedRange(sheet, dataRangeName);
                        if (impurityDataRange != null)
                        {
                            Console.WriteLine($"Fixing Impurity {peakNum} (row {summaryDataRow}) using {dataRangeName}");

                            // Update each prep column formula dynamically
                            int totalPrepColumns = summaryTable.Columns.Count - 6; // Total columns minus component + 3 statistics

                            for (int prepIndex = 0; prepIndex < totalPrepColumns; prepIndex++)
                            {
                                int prepCol = 3 + prepIndex; // Start from column C (3)
                                var cell = sheet.Cells[summaryDataRow, prepCol] as ExcelRange;
                                if (cell != null)
                                {
                                    try
                                    {
                                        // FIXED: Add column parameter (1) to INDEX function
                                        string newFormula = $"=FIXED(INDEX({dataRangeName}, {prepIndex + 1}, 1), {precisionRef})";
                                        cell.Formula = newFormula;

                                        Console.WriteLine($"  {GetColumnLetter(prepCol)}{summaryDataRow}: → {newFormula}");
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"  ❌ Error updating cell {GetColumnLetter(prepCol)}{summaryDataRow}: {ex.Message}");
                                    }
                                }
                            }
                        }
                        else
                        {
                            Console.WriteLine($"❌ Named range {dataRangeName} not found");
                        }
                    }

                    Console.WriteLine($"✅ Completed formula fixes for Table {tableNum}");
                }
            }

            Console.WriteLine("✅ Impurity section processing complete!");
        }

        private static void UpdateImpurityMeanFormulas(Worksheet sheet, int numReplicates, int impurityNumSamples, int numTables)
        {
            Console.WriteLine($"🔧 Updating Mean formulas for {numTables} tables, {impurityNumSamples} impurities, {numReplicates} reps...");

            // Calculate the Mean column position (numReplicates + 1)
            int meanColumn = numReplicates + 2;
            Console.WriteLine($"📊 Mean column calculated at position: {meanColumn}");

            try
            {
                // Loop through each summary table
                for (int tableNum = 1; tableNum <= numTables; tableNum++)
                {
                    string summaryTableName = tableNum == 1 ? "ImpuritySummaryTable" : $"ImpuritySummaryTable{tableNum}";

                    Console.WriteLine($"\n🎯 Processing {summaryTableName}...");

                    // Check if the summary table exists
                    if (!WorksheetUtilities.NamedRangeExist(sheet, summaryTableName))
                    {
                        Console.WriteLine($"❌ Warning: {summaryTableName} not found, skipping...");
                        continue;
                    }

                    // Get the summary table range - use a fresh reference each time
                    ExcelRange summaryTableRange = null;
                    try
                    {
                        summaryTableRange = WorksheetUtilities.GetNamedRange(sheet, summaryTableName);
                        if (summaryTableRange == null)
                        {
                            Console.WriteLine($"❌ Error: Could not get range for {summaryTableName}");
                            continue;
                        }

                        // Loop through each impurity row
                        for (int impNum = 1; impNum <= impurityNumSamples; impNum++)
                        {
                            ExcelRange targetCell = null;
                            try
                            {
                                // Calculate target row (row 4 = Impurity 1, row 5 = Impurity 2, etc.)
                                int targetRow = 3 + impNum; // Row 4, 5, 6, etc.

                                // Determine source stats range name
                                string statsRangeName;
                                if (tableNum == 1)
                                {
                                    statsRangeName = $"ImpurityStats{impNum}";
                                }
                                else
                                {
                                    statsRangeName = $"ImpurityStats{impNum}_T{tableNum}";
                                }

                                Console.WriteLine($"   📝 Setting [{targetRow}, {meanColumn}] = INDEX({statsRangeName}, 1, 1)");

                                // Check if source stats range exists
                                if (!WorksheetUtilities.NamedRangeExist(sheet, statsRangeName))
                                {
                                    Console.WriteLine($"   ⚠️  Warning: {statsRangeName} not found, skipping...");
                                    continue;
                                }

                                // Get the target cell in the summary table
                                targetCell = summaryTableRange.Cells[targetRow, meanColumn] as ExcelRange;
                                if (targetCell != null)
                                {
                                    // Set the formula using INDEX (same pattern as prep values)
                                    string formula = $"=INDEX({statsRangeName}, 1, 1)";
                                    targetCell.Formula = formula;

                                    Console.WriteLine($"   ✅ Successfully set formula: {formula}");
                                }
                                else
                                {
                                    Console.WriteLine($"   ❌ Error: Could not get target cell [{targetRow}, {meanColumn}]");
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"   ❌ Error setting formula for Impurity {impNum}: {ex.Message}");
                            }
                            finally
                            {
                                // Always release the target cell
                                if (targetCell != null)
                                {
                                    WorksheetUtilities.ReleaseComObject(targetCell);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"   ❌ Error processing {summaryTableName}: {ex.Message}");
                    }
                    finally
                    {
                        // Always release the summary table range
                        if (summaryTableRange != null)
                        {
                            WorksheetUtilities.ReleaseComObject(summaryTableRange);
                        }
                    }

                    Console.WriteLine($"✅ Completed {summaryTableName}");
                }

                Console.WriteLine($"🎉 Mean formula updates complete!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in UpdateImpurityMeanFormulas: {ex.Message}");
            }
        }

        private static void UpdateImpurityStdDevFormulas(Worksheet sheet, int numReplicates, int impurityNumSamples, int numTables)
        {
            Console.WriteLine($"🔧 Updating Std Dev formulas for {numTables} tables, {impurityNumSamples} impurities, {numReplicates} reps...");

            // Calculate the Std Dev column position (numReplicates + 3)
            int stdDevColumn = numReplicates + 3;
            Console.WriteLine($"📊 Std Dev column calculated at position: {stdDevColumn}");

            try
            {
                // Loop through each summary table
                for (int tableNum = 1; tableNum <= numTables; tableNum++)
                {
                    string summaryTableName = tableNum == 1 ? "ImpuritySummaryTable" : $"ImpuritySummaryTable{tableNum}";

                    Console.WriteLine($"\n🎯 Processing {summaryTableName}...");

                    // Check if the summary table exists
                    if (!WorksheetUtilities.NamedRangeExist(sheet, summaryTableName))
                    {
                        Console.WriteLine($"❌ Warning: {summaryTableName} not found, skipping...");
                        continue;
                    }

                    // Get the summary table range - use a fresh reference each time
                    ExcelRange summaryTableRange = null;
                    try
                    {
                        summaryTableRange = WorksheetUtilities.GetNamedRange(sheet, summaryTableName);
                        if (summaryTableRange == null)
                        {
                            Console.WriteLine($"❌ Error: Could not get range for {summaryTableName}");
                            continue;
                        }

                        // Loop through each impurity row
                        for (int impNum = 1; impNum <= impurityNumSamples; impNum++)
                        {
                            ExcelRange targetCell = null;
                            try
                            {
                                // Calculate target row (row 4 = Impurity 1, row 5 = Impurity 2, etc.)
                                int targetRow = 3 + impNum; // Row 4, 5, 6, etc.

                                // Determine source stats range name
                                string statsRangeName;
                                if (tableNum == 1)
                                {
                                    statsRangeName = $"ImpurityStats{impNum}";
                                }
                                else
                                {
                                    statsRangeName = $"ImpurityStats{impNum}_T{tableNum}";
                                }

                                Console.WriteLine($"   📝 Setting [{targetRow}, {stdDevColumn}] = INDEX({statsRangeName}, 5, 1)");

                                // Check if source stats range exists
                                if (!WorksheetUtilities.NamedRangeExist(sheet, statsRangeName))
                                {
                                    Console.WriteLine($"   ⚠️  Warning: {statsRangeName} not found, skipping...");
                                    continue;
                                }

                                // Get the target cell in the summary table
                                targetCell = summaryTableRange.Cells[targetRow, stdDevColumn] as ExcelRange;
                                if (targetCell != null)
                                {
                                    // Set the formula using INDEX - Std Dev is at row 5 in stats range
                                    string formula = $"=FIXED(INDEX({statsRangeName}, 5, 1), ImpStdDevPrecision)";

                                    //string formula = $"=INDEX({statsRangeName}, 5, 1)";
                                    targetCell.Formula = formula;

                                    Console.WriteLine($"   ✅ Successfully set formula: {formula}");
                                }
                                else
                                {
                                    Console.WriteLine($"   ❌ Error: Could not get target cell [{targetRow}, {stdDevColumn}]");
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"   ❌ Error setting formula for Impurity {impNum}: {ex.Message}");
                            }
                            finally
                            {
                                // Always release the target cell
                                if (targetCell != null)
                                {
                                    WorksheetUtilities.ReleaseComObject(targetCell);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"   ❌ Error processing {summaryTableName}: {ex.Message}");
                    }
                    finally
                    {
                        // Always release the summary table range
                        if (summaryTableRange != null)
                        {
                            WorksheetUtilities.ReleaseComObject(summaryTableRange);
                        }
                    }

                    Console.WriteLine($"✅ Completed {summaryTableName}");
                }

                Console.WriteLine($"🎉 Std Dev formula updates complete!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in UpdateImpurityStdDevFormulas: {ex.Message}");
            }
        }

        private static void UpdateImpurityRSDFormulas(Worksheet sheet, int numReplicates, int impurityNumSamples, int numTables)
        {
            Console.WriteLine($"🔧 Updating RSD formulas for {numTables} tables, {impurityNumSamples} impurities, {numReplicates} reps...");

            // Calculate the RSD column position (numReplicates + 4)
            int rsdColumn = numReplicates + 4;
            Console.WriteLine($"📊 RSD column calculated at position: {rsdColumn}");

            try
            {
                // Loop through each summary table
                for (int tableNum = 1; tableNum <= numTables; tableNum++)
                {
                    string summaryTableName = tableNum == 1 ? "ImpuritySummaryTable" : $"ImpuritySummaryTable{tableNum}";

                    Console.WriteLine($"\n🎯 Processing {summaryTableName}...");

                    // Check if the summary table exists
                    if (!WorksheetUtilities.NamedRangeExist(sheet, summaryTableName))
                    {
                        Console.WriteLine($"❌ Warning: {summaryTableName} not found, skipping...");
                        continue;
                    }

                    // Get the summary table range - use a fresh reference each time
                    ExcelRange summaryTableRange = null;
                    try
                    {
                        summaryTableRange = WorksheetUtilities.GetNamedRange(sheet, summaryTableName);
                        if (summaryTableRange == null)
                        {
                            Console.WriteLine($"❌ Error: Could not get range for {summaryTableName}");
                            continue;
                        }

                        // Loop through each impurity row
                        for (int impNum = 1; impNum <= impurityNumSamples; impNum++)
                        {
                            ExcelRange targetCell = null;
                            try
                            {
                                // Calculate target row (row 4 = Impurity 1, row 5 = Impurity 2, etc.)
                                int targetRow = 3 + impNum; // Row 4, 5, 6, etc.

                                // Determine source stats range name
                                string statsRangeName;
                                if (tableNum == 1)
                                {
                                    statsRangeName = $"ImpurityStats{impNum}";
                                }
                                else
                                {
                                    statsRangeName = $"ImpurityStats{impNum}_T{tableNum}";
                                }

                                Console.WriteLine($"   📝 Setting [{targetRow}, {rsdColumn}] = INDEX({statsRangeName}, 4, 1)");

                                // Check if source stats range exists
                                if (!WorksheetUtilities.NamedRangeExist(sheet, statsRangeName))
                                {
                                    Console.WriteLine($"   ⚠️  Warning: {statsRangeName} not found, skipping...");
                                    continue;
                                }

                                // Get the target cell in the summary table
                                targetCell = summaryTableRange.Cells[targetRow, rsdColumn] as ExcelRange;
                                if (targetCell != null)
                                {
                                    // Set the formula using INDEX - RSD is at row 4 in stats range
                                    string formula = $"=INDEX({statsRangeName}, 4, 1)";
                                    targetCell.Formula = formula;

                                    Console.WriteLine($"   ✅ Successfully set formula: {formula}");
                                }
                                else
                                {
                                    Console.WriteLine($"   ❌ Error: Could not get target cell [{targetRow}, {rsdColumn}]");
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"   ❌ Error setting formula for Impurity {impNum}: {ex.Message}");
                            }
                            finally
                            {
                                // Always release the target cell
                                if (targetCell != null)
                                {
                                    WorksheetUtilities.ReleaseComObject(targetCell);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"   ❌ Error processing {summaryTableName}: {ex.Message}");
                    }
                    finally
                    {
                        // Always release the summary table range
                        if (summaryTableRange != null)
                        {
                            WorksheetUtilities.ReleaseComObject(summaryTableRange);
                        }
                    }

                    Console.WriteLine($"✅ Completed {summaryTableName}");
                }

                Console.WriteLine($"🎉 RSD formula updates complete!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in UpdateImpurityRSDFormulas: {ex.Message}");
            }
        }

        private static void UpdateImpurityLowerConfidenceFormulas(Worksheet sheet, int numReplicates, int impurityNumSamples, int numTables)
        {
            Console.WriteLine($"🔧 Updating Lower Confidence Interval formulas for {numTables} tables, {impurityNumSamples} impurities, {numReplicates} reps...");

            // Calculate the Lower Confidence column position (numReplicates + 5)
            int lowerConfidenceColumn = numReplicates + 5;
            Console.WriteLine($"📊 Lower Confidence column calculated at position: {lowerConfidenceColumn}");

            try
            {
                // Loop through each summary table
                for (int tableNum = 1; tableNum <= numTables; tableNum++)
                {
                    string summaryTableName = tableNum == 1 ? "ImpuritySummaryTable" : $"ImpuritySummaryTable{tableNum}";

                    Console.WriteLine($"\n🎯 Processing {summaryTableName}...");

                    // Check if the summary table exists
                    if (!WorksheetUtilities.NamedRangeExist(sheet, summaryTableName))
                    {
                        Console.WriteLine($"❌ Warning: {summaryTableName} not found, skipping...");
                        continue;
                    }

                    // Get the summary table range - use a fresh reference each time
                    ExcelRange summaryTableRange = null;
                    try
                    {
                        summaryTableRange = WorksheetUtilities.GetNamedRange(sheet, summaryTableName);
                        if (summaryTableRange == null)
                        {
                            Console.WriteLine($"❌ Error: Could not get range for {summaryTableName}");
                            continue;
                        }

                        // Loop through each impurity row
                        for (int impNum = 1; impNum <= impurityNumSamples; impNum++)
                        {
                            ExcelRange targetCell = null;
                            try
                            {
                                // Calculate target row (row 4 = Impurity 1, row 5 = Impurity 2, etc.)
                                int targetRow = 3 + impNum; // Row 4, 5, 6, etc.

                                // Determine source stats range name
                                string statsRangeName;
                                if (tableNum == 1)
                                {
                                    statsRangeName = $"ImpurityStats{impNum}";
                                }
                                else
                                {
                                    statsRangeName = $"ImpurityStats{impNum}_T{tableNum}";
                                }

                                Console.WriteLine($"   📝 Setting [{targetRow}, {lowerConfidenceColumn}] = INDEX({statsRangeName}, 10, 1)");

                                // Check if source stats range exists
                                if (!WorksheetUtilities.NamedRangeExist(sheet, statsRangeName))
                                {
                                    Console.WriteLine($"   ⚠️  Warning: {statsRangeName} not found, skipping...");
                                    continue;
                                }

                                // Get the target cell in the summary table
                                targetCell = summaryTableRange.Cells[targetRow, lowerConfidenceColumn] as ExcelRange;
                                if (targetCell != null)
                                {
                                    // Set the formula using INDEX - Lower Confidence is at row 10 in stats range
                                    string formula = $"=INDEX({statsRangeName}, 10, 1)";
                                    targetCell.Formula = formula;

                                    Console.WriteLine($"   ✅ Successfully set formula: {formula}");
                                }
                                else
                                {
                                    Console.WriteLine($"   ❌ Error: Could not get target cell [{targetRow}, {lowerConfidenceColumn}]");
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"   ❌ Error setting formula for Impurity {impNum}: {ex.Message}");
                            }
                            finally
                            {
                                // Always release the target cell
                                if (targetCell != null)
                                {
                                    WorksheetUtilities.ReleaseComObject(targetCell);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"   ❌ Error processing {summaryTableName}: {ex.Message}");
                    }
                    finally
                    {
                        // Always release the summary table range
                        if (summaryTableRange != null)
                        {
                            WorksheetUtilities.ReleaseComObject(summaryTableRange);
                        }
                    }

                    Console.WriteLine($"✅ Completed {summaryTableName}");
                }

                Console.WriteLine($"🎉 Lower Confidence Interval formula updates complete!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in UpdateImpurityLowerConfidenceFormulas: {ex.Message}");
            }
        }

        private static void UpdateImpurityUpperConfidenceFormulas(Worksheet sheet, int numReplicates, int impurityNumSamples, int numTables)
        {
            Console.WriteLine($"🔧 Updating Upper Confidence Interval formulas for {numTables} tables, {impurityNumSamples} impurities, {numReplicates} reps...");

            // Calculate the Upper Confidence column position (numReplicates + 6)
            int upperConfidenceColumn = numReplicates + 6;
            Console.WriteLine($"📊 Upper Confidence column calculated at position: {upperConfidenceColumn}");

            try
            {
                // Loop through each summary table
                for (int tableNum = 1; tableNum <= numTables; tableNum++)
                {
                    string summaryTableName = tableNum == 1 ? "ImpuritySummaryTable" : $"ImpuritySummaryTable{tableNum}";

                    Console.WriteLine($"\n🎯 Processing {summaryTableName}...");

                    // Check if the summary table exists
                    if (!WorksheetUtilities.NamedRangeExist(sheet, summaryTableName))
                    {
                        Console.WriteLine($"❌ Warning: {summaryTableName} not found, skipping...");
                        continue;
                    }

                    // Get the summary table range - use a fresh reference each time
                    ExcelRange summaryTableRange = null;
                    try
                    {
                        summaryTableRange = WorksheetUtilities.GetNamedRange(sheet, summaryTableName);
                        if (summaryTableRange == null)
                        {
                            Console.WriteLine($"❌ Error: Could not get range for {summaryTableName}");
                            continue;
                        }

                        // Loop through each impurity row
                        for (int impNum = 1; impNum <= impurityNumSamples; impNum++)
                        {
                            ExcelRange targetCell = null;
                            try
                            {
                                // Calculate target row (row 4 = Impurity 1, row 5 = Impurity 2, etc.)
                                int targetRow = 3 + impNum; // Row 4, 5, 6, etc.

                                // Determine source stats range name
                                string statsRangeName;
                                if (tableNum == 1)
                                {
                                    statsRangeName = $"ImpurityStats{impNum}";
                                }
                                else
                                {
                                    statsRangeName = $"ImpurityStats{impNum}_T{tableNum}";
                                }

                                Console.WriteLine($"   📝 Setting [{targetRow}, {upperConfidenceColumn}] = INDEX({statsRangeName}, 11, 1)");

                                // Check if source stats range exists
                                if (!WorksheetUtilities.NamedRangeExist(sheet, statsRangeName))
                                {
                                    Console.WriteLine($"   ⚠️  Warning: {statsRangeName} not found, skipping...");
                                    continue;
                                }

                                // Get the target cell in the summary table
                                targetCell = summaryTableRange.Cells[targetRow, upperConfidenceColumn] as ExcelRange;
                                if (targetCell != null)
                                {
                                    // Set the formula using INDEX - Upper Confidence is at row 11 in stats range
                                    string formula = $"=INDEX({statsRangeName}, 11, 1)";
                                    targetCell.Formula = formula;

                                    Console.WriteLine($"   ✅ Successfully set formula: {formula}");
                                }
                                else
                                {
                                    Console.WriteLine($"   ❌ Error: Could not get target cell [{targetRow}, {upperConfidenceColumn}]");
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"   ❌ Error setting formula for Impurity {impNum}: {ex.Message}");
                            }
                            finally
                            {
                                // Always release the target cell
                                if (targetCell != null)
                                {
                                    WorksheetUtilities.ReleaseComObject(targetCell);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"   ❌ Error processing {summaryTableName}: {ex.Message}");
                    }
                    finally
                    {
                        // Always release the summary table range
                        if (summaryTableRange != null)
                        {
                            WorksheetUtilities.ReleaseComObject(summaryTableRange);
                        }
                    }

                    Console.WriteLine($"✅ Completed {summaryTableName}");
                }

                Console.WriteLine($"🎉 Upper Confidence Interval formula updates complete!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in UpdateImpurityUpperConfidenceFormulas: {ex.Message}");
            }
        }

        // Update the HandleImpurity method to call this new function
        private static void HandleImpurity(Worksheet sheet, int numReplicates, int impurityNumSamples)
        {
            Console.WriteLine($"\n=== Processing Impurity section with {numReplicates} replicates and {impurityNumSamples} impurities ===");

            if (impurityNumSamples <= 0)
            {
                if (WorksheetUtilities.NamedRangeExist(sheet, "ImpurityAndImpuritySummary"))
                {
                    WorksheetUtilities.DeleteNamedRangeRows(sheet, "ImpurityAndImpuritySummary");
                    WorksheetUtilities.DeleteNamedRange(sheet, "ImpurityAndImpuritySummary");
                }
                return;
            }

            // Check if required named ranges exist
            if (!WorksheetUtilities.NamedRangeExist(sheet, "PrepNumsImpurityRawData1"))
            {
                Console.WriteLine("❌ ERROR: PrepNumsImpurityRawData1 named range not found! Skipping Impurity section.");
                return;
            }

            // Expand or contract the tables based on the number of replicates
            if (numReplicates > DefaultNumReplicates)
            {
                int numRowsToInsert = numReplicates - DefaultNumReplicates;
                WorksheetUtilities.InsertRowsIntoNamedRange(numRowsToInsert, sheet, "PrepNumsImpurityRawData1", true, XlDirection.xlDown, XlPasteType.xlPasteFormulas);
            }
            else if (numReplicates < DefaultNumReplicates)
            {
                int numRowsToRemove = DefaultNumReplicates - numReplicates;
                if (DefaultNumReplicates - numRowsToRemove < MinNumReplicates) numRowsToRemove = DefaultNumReplicates - MinNumReplicates;
                WorksheetUtilities.DeleteRowsFromNamedRange(numRowsToRemove, sheet, "PrepNumsImpurityRawData1", XlDirection.xlDown);
            }

            // Re-number the preps
            if (numReplicates < MinNumReplicates) numReplicates = MinNumReplicates;
            List<string> prepNumbers = new List<string>(0);
            for (int i = 1; i <= numReplicates; i++) prepNumbers.Add(i.ToString());
            WorksheetUtilities.SetNamedRangeValues(sheet, "PrepNumsImpurityRawData1", prepNumbers);

            // Create additional impurity columns
            int numDataTables = impurityNumSamples - DefaultNumDataTables;
            for (int i = 1; i <= numDataTables; i++)
            {
                int colOffset = 2;
                int namedRangeNum = i + 1;

                try
                {
                    if (WorksheetUtilities.NamedRangeExist(sheet, "ImpurityRawDataTable1"))
                    {
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ImpurityRawDataTable1", "ImpurityRawDataTable" + namedRangeNum, 1, (colOffset * i) + 1, XlPasteType.xlPasteAll);
                        WorksheetUtilities.SetNamedRangeValue(sheet, "ImpurityRawDataTable" + namedRangeNum, "Impurity " + namedRangeNum, 1, 1);
                    }

                    if (WorksheetUtilities.NamedRangeExist(sheet, "ImpurityResults1"))
                    {
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ImpurityResults1", "ImpurityResults" + namedRangeNum, 1, (colOffset * i) + 1, XlPasteType.xlPasteAll);
                    }

                    // Force garbage collection to prevent COM issues
                    System.GC.Collect();
                    System.GC.WaitForPendingFinalizers();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Error creating impurity column {namedRangeNum}: {ex.Message}");
                }
            }

            // Expand summary table
            int DefaultimpurityNumSampless = 1;
            if (impurityNumSamples > DefaultimpurityNumSampless && WorksheetUtilities.NamedRangeExist(sheet, "ImpuritySummary"))
            {
                int numRowsToInsert = impurityNumSamples - DefaultimpurityNumSampless;
                WorksheetUtilities.InsertRowsIntoNamedRange(numRowsToInsert, sheet, "ImpuritySummary", true, XlDirection.xlDown, XlPasteType.xlPasteAll);
            }

            Console.WriteLine("✅ Impurity section processing complete!");
        }

        private static string UpdateColumnReference(string formula, char targetColumn, int prepNumber)
        {
            // Replace D18→E18, D19→E19, etc. for the specific prep
            // Pattern: =FIXED(D18, C45) → =FIXED(E18, C45) for impurity 2
            int targetRow = 18 + (prepNumber - 1); // D18, D19, D20...
            string oldRef = $"D{targetRow}";
            string newRef = $"{targetColumn}{targetRow}";

            return formula.Replace(oldRef, newRef);
        }

        private static void CreateShiftedNamedRangeForImpurity(Worksheet sheet, ExcelRange originalRange, string newRangeName, int rowOffset, int colOffset)
        {
            try
            {
                // Calculate new range position
                int newStartRow = originalRange.Row + rowOffset;
                int newStartCol = originalRange.Column + colOffset;
                int newEndRow = newStartRow + originalRange.Rows.Count - 1;
                int newEndCol = newStartCol + originalRange.Columns.Count - 1;

                // Create the new range reference
                var startCell = sheet.Cells[newStartRow, newStartCol] as ExcelRange;
                var endCell = sheet.Cells[newEndRow, newEndCol] as ExcelRange;

                if (startCell != null && endCell != null)
                {
                    string refersToLocal = $"='{sheet.Name}'!" +
                                         startCell.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing) + ":" +
                                         endCell.get_AddressLocal(true, true, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

                    sheet.Names.Add(newRangeName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                   Type.Missing, Type.Missing, refersToLocal, Type.Missing, Type.Missing, Type.Missing);

                    Console.WriteLine($"✅ Created named range: {newRangeName}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error creating named range {newRangeName}: {ex.Message}");
            }
        }

        private static void HandleWaterContent(Worksheet sheet, int numReplicates, int wcNumSamples,
            string wcOperator1,
            decimal wcValue1,
            string wcOperator2,
            decimal wcValue2,
            string wcOperator3,
            decimal wcValue3,
            string wcOperator4,
            decimal wcValue4)
        {
            SetWaterContentAcceptanceCriteria(sheet, wcOperator1, wcValue1, wcOperator2, wcValue2, wcOperator3, wcValue3, wcOperator4, wcValue4);

            Console.WriteLine($"\n=== Processing Water Content section with {numReplicates} replicates and {wcNumSamples} samples ===");

            // Step 1: Row Expansion based on numReplicates (FIRST - already working)
            if (numReplicates > 6) // Default is 6, expand if more
            {
                int numRowsToInsert = numReplicates - 6;
                Console.WriteLine($"Expanding Water Content by {numRowsToInsert} rows...");

                // Only expand the prep column - other columns expand automatically
                WorksheetUtilities.InsertRowsIntoNamedRange(numRowsToInsert, sheet, "WaterContentPrep1", true, XlDirection.xlDown, XlPasteType.xlPasteFormulas);
            }
            else if (numReplicates < 6) // Contract if less than 6
            {
                int numRowsToRemove = 6 - numReplicates;
                if (numRowsToRemove > 4) numRowsToRemove = 4; // Keep minimum 2 rows
                Console.WriteLine($"Contracting Water Content by {numRowsToRemove} rows...");

                WorksheetUtilities.DeleteRowsFromNamedRange(numRowsToRemove, sheet, "WaterContentPrep1", XlDirection.xlDown);
            }

            // Step 2: Update preparation numbers
            if (numReplicates != 6) // Only if we changed the row count
            {
                Console.WriteLine($"Updating Water Content preparation numbers 1-{numReplicates}...");

                List<string> prepNumbers = new List<string>();
                for (int i = 1; i <= numReplicates; i++)
                {
                    prepNumbers.Add(i.ToString());
                }

                WorksheetUtilities.SetNamedRangeValues(sheet, "WaterContentPrep1", prepNumbers);
            }

            // Step 3: Column Copying based on wcNumSamples (SECOND - after row expansion)
            if (wcNumSamples > 1)
            {
                Console.WriteLine($"Creating {wcNumSamples - 1} additional Water Content sample columns...");

                for (int sampleNum = 2; sampleNum <= wcNumSamples; sampleNum++)
                {
                    int colOffset = sampleNum - 1; // Sample 2 = 1 col right (D), Sample 3 = 2 cols right (E)

                    Console.WriteLine($"Creating Water Content Sample {sampleNum} column at column offset {colOffset}...");

                    try
                    {
                        // Step 3.1: Copy headers first (WaterContentHeader1 → WaterContentHeader2, etc.)
                        string destHeaderName = $"WaterContentHeader{sampleNum}";
                        Console.WriteLine($"Copying headers: WaterContentHeader1 → {destHeaderName}");
                        // Use colOffset + 1 like HandleAssay does
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "WaterContentHeader1", destHeaderName, 1, colOffset + 1, XlPasteType.xlPasteAll);

                        // Step 3.2: Copy data second (WaterContentPercent1 → WaterContentPercent2, etc.)
                        string destDataName = $"WaterContentPercent{sampleNum}";
                        Console.WriteLine($"Copying data: WaterContentPercent1 → {destDataName}");
                        // Use colOffset + 1 like HandleAssay does
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "WaterContentPercent1", destDataName, 1, colOffset + 1, XlPasteType.xlPasteAll);

                        var sourceColumn = sheet.Columns[3] as ExcelRange; // Column C
                        var destColumn = sheet.Columns[3 + colOffset + 1] as ExcelRange; // Target column
                        if (sourceColumn != null && destColumn != null) destColumn.ColumnWidth = sourceColumn.ColumnWidth;

                        Console.WriteLine($"✅ Successfully created Water Content Sample {sampleNum} column");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"❌ Error creating Water Content Sample {sampleNum} column: {ex.Message}");
                    }
                }
            }
            else
            {
                Console.WriteLine("No additional Water Content columns needed (wcNumSamples = 1)");
            }

            // Step 4: Statistics Column Copying (THIRD - after everything else)

            if (wcNumSamples > 1)
            {
                Console.WriteLine($"Creating {wcNumSamples - 1} additional Water Content statistics columns...");

                for (int sampleNum = 2; sampleNum <= wcNumSamples; sampleNum++)
                {
                    int colOffset = sampleNum - 1; // Sample 2 = 1 col right, Sample 3 = 2 cols right

                    Console.WriteLine($"Creating Water Content Statistics Sample {sampleNum} column at column offset {colOffset}...");

                    try
                    {
                        // Copy statistics column (WaterContentStats1 → WaterContentStats2, etc.)
                        string destStatsName = $"WaterContentStats{sampleNum}";
                        Console.WriteLine($"Copying statistics: WaterContentStats1 → {destStatsName}");

                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "WaterContentStats1", destStatsName, 1, colOffset + 1, XlPasteType.xlPasteAll);

                        // Copy column width for statistics
                        var sourceStatsColumn = sheet.Columns[3] as ExcelRange; // Column C (original stats)
                        var destStatsColumn = sheet.Columns[3 + colOffset + 1] as ExcelRange; // Target column
                        if (sourceStatsColumn != null && destStatsColumn != null)
                            destStatsColumn.ColumnWidth = sourceStatsColumn.ColumnWidth;

                        Console.WriteLine($"✅ Successfully created Water Content Statistics Sample {sampleNum} column");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"❌ Error creating Water Content Statistics Sample {sampleNum} column: {ex.Message}");
                    }
                }
            }
            else
            {
                Console.WriteLine("No additional Water Content statistics columns needed (wcNumSamples = 1)");
            }

            Console.WriteLine("✅ Water Content section processing complete!");
        }

        private static void SetWaterContentAcceptanceCriteria(
            Worksheet sheet,
            string wcOperator1,
            decimal wcValue1,
            string wcOperator2,
            decimal wcValue2,
            string wcOperator3,
            decimal wcValue3,
            string wcOperator4,
            decimal wcValue4)
        {
            try
            {
                Console.WriteLine($"🔧 Setting Water Content Acceptance Criteria...");

                var range = WorksheetUtilities.GetNamedRange(sheet, "WaterContentAcceptanceCriteriaRange");
                if (range != null)
                {
                    // Row 1: ≤ 1.57 ≤ 2.01
                    (range.Cells[1, 1] as ExcelRange).Value2 = wcOperator1;      // "≤"
                    (range.Cells[1, 2] as ExcelRange).Value2 = wcValue1;   // "1.57"
                    (range.Cells[1, 3] as ExcelRange).Value2 = wcOperator2;      // "≤"
                    (range.Cells[1, 4] as ExcelRange).Value2 = wcValue2;   // "2.01"

                    // Row 2: ≥ 1.57 ≤ 1.3
                    (range.Cells[2, 1] as ExcelRange).Value2 = wcOperator3;      // "≥"
                    (range.Cells[2, 2] as ExcelRange).Value2 = wcValue3;   // "1.57"
                    (range.Cells[2, 3] as ExcelRange).Value2 = wcOperator4;      // "≤"
                    (range.Cells[2, 4] as ExcelRange).Value2 = wcValue4;   // "1.3"

                    Console.WriteLine($"✅ Water Content Acceptance Criteria updated:");
                    Console.WriteLine($"   Row 1: {wcOperator1} {wcValue1} {wcOperator2} {wcValue2}");
                    Console.WriteLine($"   Row 2: {wcOperator3} {wcValue3} {wcOperator4} {wcValue4}");

                    WorksheetUtilities.ReleaseComObject(range);
                }
                else
                {
                    Console.WriteLine($"⚠️ Named range 'WaterContentAcceptanceCriteriaRange' not found");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error setting Water Content Acceptance Criteria: {ex.Message}");
            }
        }

        private static void HandleWaterContentSummary(Worksheet sheet, int numReplicates, int wcNumSamples)
        {
            Console.WriteLine($"\n=== Processing Water Content Summary section with {numReplicates} replicates and {wcNumSamples} samples ===");

            // Step 1: Row Expansion based on numReplicates (FIRST)
            if (numReplicates > 6) // Default is 6, expand if more
            {
                int numRowsToInsert = numReplicates - 6;
                Console.WriteLine($"Expanding Water Content Summary by {numRowsToInsert} rows...");

                // Only expand the prep column - other columns expand automatically
                WorksheetUtilities.InsertRowsIntoNamedRange(numRowsToInsert, sheet, "WcSummaryPrep1", true, XlDirection.xlDown, XlPasteType.xlPasteFormulas);
            }
            else if (numReplicates < 6) // Contract if less than 6
            {
                int numRowsToRemove = 6 - numReplicates;
                if (numRowsToRemove > 4) numRowsToRemove = 4; // Keep minimum 2 rows
                Console.WriteLine($"Contracting Water Content Summary by {numRowsToRemove} rows...");

                WorksheetUtilities.DeleteRowsFromNamedRange(numRowsToRemove, sheet, "WcSummaryPrep1", XlDirection.xlDown);
            }

            // Step 2: Update preparation numbers
            if (numReplicates != 6) // Only if we changed the row count
            {
                Console.WriteLine($"Updating Water Content Summary preparation numbers 1-{numReplicates}...");

                List<string> prepNumbers = new List<string>();
                for (int i = 1; i <= numReplicates; i++)
                {
                    prepNumbers.Add(i.ToString());
                }

                WorksheetUtilities.SetNamedRangeValues(sheet, "WcSummaryPrep1", prepNumbers);
            }

            // Step 3: Column Copying based on wcNumSamples
            if (wcNumSamples > 1)
            {
                Console.WriteLine($"Creating {wcNumSamples - 1} additional Water Content Summary sample columns...");

                for (int sampleNum = 2; sampleNum <= wcNumSamples; sampleNum++)
                {
                    int colOffset = sampleNum - 1; // Sample 2 = 1 col right (D), Sample 3 = 2 cols right (E)

                    Console.WriteLine($"Creating Water Content Summary Sample {sampleNum} column at column offset {colOffset}...");

                    try
                    {
                        // Step 3.1: Copy headers first (WcSummaryHeader1 → WcSummaryHeader2, etc.)
                        string destHeaderName = $"WcSummaryHeader{sampleNum}";
                        Console.WriteLine($"Copying headers: WcSummaryHeader1 → {destHeaderName}");
                        // Use colOffset + 1 like HandleAssay does
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "WcSummaryHeader1", destHeaderName, 1, colOffset + 1, XlPasteType.xlPasteAll);

                        // Step 3.2: Copy data second (WcSummaryPercent1 → WcSummaryPercent2, etc.)
                        string destDataName = $"WcSummaryPercent{sampleNum}";
                        Console.WriteLine($"Copying data: WcSummaryPercent1 → {destDataName}");
                        // Use colOffset + 1 like HandleAssay does
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "WcSummaryPercent1", destDataName, 1, colOffset + 1, XlPasteType.xlPasteAll);

                        // Step 3.3: Copy column width (simple one-liner)
                        var sourceColumn = sheet.Columns[3] as ExcelRange; // Column C (or whatever the source column is)
                        var destColumn = sheet.Columns[3 + colOffset + 1] as ExcelRange; // Target column
                        if (sourceColumn != null && destColumn != null) destColumn.ColumnWidth = sourceColumn.ColumnWidth;

                        Console.WriteLine($"✅ Successfully created Water Content Summary Sample {sampleNum} column");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"❌ Error creating Water Content Summary Sample {sampleNum} column: {ex.Message}");
                    }
                }
            }

            // Step 4: Statistics Column Copying (THIRD - NEW)
            if (wcNumSamples > 1)
            {
                Console.WriteLine($"Creating {wcNumSamples - 1} additional Water Content Summary statistics columns...");

                for (int sampleNum = 2; sampleNum <= wcNumSamples; sampleNum++)
                {
                    int colOffset = sampleNum - 1; // Sample 2 = 1 col right, Sample 3 = 2 cols right

                    Console.WriteLine($"Creating Water Content Summary Statistics Sample {sampleNum} column at column offset {colOffset}...");

                    try
                    {
                        // Copy statistics column (WcSummaryStats1 → WcSummaryStats2, etc.)
                        string destStatsName = $"WcSummaryStats{sampleNum}";
                        Console.WriteLine($"Copying statistics: WcSummaryStats1 → {destStatsName}");

                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "WcSummaryStats1", destStatsName, 1, colOffset + 1, XlPasteType.xlPasteAll);

                        // Copy column width for statistics
                        var sourceStatsColumn = sheet.Columns[3] as ExcelRange; // Column C (original stats)
                        var destStatsColumn = sheet.Columns[3 + colOffset + 1] as ExcelRange; // Target column
                        if (sourceStatsColumn != null && destStatsColumn != null)
                            destStatsColumn.ColumnWidth = sourceStatsColumn.ColumnWidth;

                        Console.WriteLine($"✅ Successfully created Water Content Summary Statistics Sample {sampleNum} column");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"❌ Error creating Water Content Summary Statistics Sample {sampleNum} column: {ex.Message}");
                    }
                }
            }
            else
            {
                Console.WriteLine("No additional Water Content Summary statistics columns needed (wcNumSamples = 1)");
            }

            Console.WriteLine("✅ Water Content Summary section processing complete!");
        }
    }
}