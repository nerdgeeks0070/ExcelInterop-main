using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using log4net.Core;
using Microsoft.Office.Interop.Excel;

namespace Spreadsheet.Handler
{
    public static class AlternateAutomatedSampling
    {
        private static Application _app;

        private const string TempDirectoryName = "ABD_TempFiles";

        private const int defaultCompSet = 1;
        private const int defaultTimepoints = 5;

        public static string UpdateAlternateAutomatedSamplingSheet(string sourcePath, int compSet, int timepoints,
            Dictionary<string, string> acceptanceCriteria)
        {
            string returnPath = "";
            try
            {
                returnPath = UpdateAlternateAutoSamplingSheet(sourcePath, compSet, timepoints, acceptanceCriteria);
            }
            catch (Exception ex)
            {
                Logger.LogMessage("An error occurred in the call to AlternateAutomatedSampling.UpdateAlternateAutomatedSamplingSheet. Message and stack trace are:\r\n" + ex.Message + "\r\n" + ex.StackTrace, Level.Error);

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
                            Logger.LogMessage("An error occurred in the call to AlternateAutomatedSampling.UpdateAlternateAutomatedSamplingSheet. Failed to save current workbook changes and to get path.", Level.Error);
                        }

                        _app.Workbooks.Close();
                    }
                    _app = null;
                }
                catch
                {
                    Logger.LogMessage("An error occurred in the call to AlternateAutomatedSampling.UpdateAlternateAutomatedSamplingSheet. Application failed to close workbooks. Message and stack trace are:\r\n" + ex.Message + "\r\n" + ex.StackTrace, Level.Error);
                }
                finally
                {
                    WorksheetUtilities.ReleaseExcelApp();
                }
            }
            return returnPath;
        }

        private static string UpdateAlternateAutoSamplingSheet(string sourcePath, int compSet, int timepoints, Dictionary<string, string> acceptanceCriteria = null)
        {
            if (!File.Exists(sourcePath))
            {
                Logger.LogMessage("Error in call to AlternateAutomatedSampling.UpdateAlternateAutomatedSamplingSheet. Invalid source file path specified.", Level.Error);
                return "";
            }

            // Generate an random temp path to save new workbook
            string savePath = WorksheetUtilities.CopyWorkbook(sourcePath, TempDirectoryName, "Alternate Automated Sampling Results.xls");
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

                if (timepoints > defaultTimepoints)
                {

                    int toInsert = timepoints - defaultTimepoints;

                    for (int i = 1; i < compSet; i++)
                    {
                        WorksheetUtilities.InsertRowsIntoNamedRange(toInsert, sheet, "Sampler_Timepoints_Table" + i + "1", true, XlDirection.xlDown, XlPasteType.xlPasteFormulas);
                        WorksheetUtilities.InsertRowsIntoNamedRange(toInsert, sheet, "Sampler_Timepoints_Table" + i + "2", true, XlDirection.xlDown, XlPasteType.xlPasteFormulas);
                        WorksheetUtilities.InsertRowsIntoNamedRange(toInsert, sheet, "Sampler_Timepoints_Table" + i + "3", true, XlDirection.xlDown, XlPasteType.xlPasteFormulas);
                    }

                    for (int i = 1; i <= compSet; i++)
                    {
                        WorksheetUtilities.ResizeNamedRange(sheet, "Sampler_Timepoints_Table11", toInsert, 0);
                        WorksheetUtilities.ResizeNamedRange(sheet, "Sampler_Timepoints_Table12", toInsert, 0);
                        WorksheetUtilities.ResizeNamedRange(sheet, "Sampler_Timepoints_Table13", toInsert, 0);
                    }
                }

                if (timepoints < defaultTimepoints && timepoints != 0)
                {

                    var rowsToDelete = Math.Abs(timepoints - defaultTimepoints);

                    WorksheetUtilities.DeleteRowsFromNamedRange(rowsToDelete, sheet, "Sampler_Timepoints_Table11", XlDirection.xlDown);
                    WorksheetUtilities.DeleteRowsFromNamedRange(rowsToDelete, sheet, "Sampler_Timepoints_Table12", XlDirection.xlDown);
                    WorksheetUtilities.DeleteRowsFromNamedRange(rowsToDelete, sheet, "Sampler_Timepoints_Table13", XlDirection.xlDown);
                }

                if (compSet > defaultCompSet)
                {
                    int rowsToInsert = compSet - defaultCompSet;
                    AppendRowsCopyOnlyFormulasAndRenumber(sheet, "Sample_Table", rowsToInsert, 1, 1, 2);
                    int headerColIndex = 1;

                    for (int i = 2; i <= compSet; i++)
                    {
                        WorksheetUtilities.InsertRowsIntoNamedRange(timepoints * 2 + 14, sheet, "SampleInfoData", false, XlDirection.xlUp, XlPasteType.xlPasteAll);
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ComparisonSet" + (i - 1), "ComparisonSet" + i, timepoints * 2 + 14, 1, XlPasteType.xlPasteAll);
                        WorksheetUtilities.SetNamedRangeValue(sheet, "ComparisonSet" + i, ("Set-" + i).Trim(), 1, 1);

                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "SetManualDisso" + (i - 1), "SetManualDisso" + i, timepoints * 2 + 14, 1, XlPasteType.xlPasteAll);

                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "SetAutoDisso" + (i - 1), "SetAutoDisso" + i, timepoints * 2 + 14, 1, XlPasteType.xlPasteAll);

                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "Auto_Sampler_Timepoints" + (i - 1) + "1", "Auto_Sampler_Timepoints" + i + 1, timepoints * 2 + 14, 1, XlPasteType.xlPasteAll);

                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "Sampler1_Mean" + (i - 1), "Sampler1_Mean" + i, timepoints * 2 + 14, 1, XlPasteType.xlPasteAll);
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "Sampler2_Mean" + (i - 1), "Sampler2_Mean" + i, timepoints * 2 + 14, 1, XlPasteType.xlPasteAll);

                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "Sample_Name" + (i - 1), "Sample_Name" + i, timepoints * 2 + 14, 1, XlPasteType.xlPasteAll);

                        WorksheetUtilities.InsertRowsIntoNamedRange(timepoints + 1, sheet, "SummaryData", false, XlDirection.xlUp, XlPasteType.xlPasteAll);

                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "Summary_Table" + (i - 1), "Summary_Table" + i, timepoints + 1 + 1, 1, XlPasteType.xlPasteAll);

                        SetSamplingSummaryFormula(sheet, i);

                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "SummaryTimepoints" + (i - 1), "SummaryTimepoints" + i, timepoints + 1 + 1, 1, XlPasteType.xlPasteAll);

                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "SummaryManualMeans" + (i - 1), "SummaryManualMeans" + i, timepoints + 1 + 1, 1, XlPasteType.xlPasteAll);
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "SummaryAutoMeans" + (i - 1), "SummaryAutoMeans" + i, timepoints + 1 + 1, 1, XlPasteType.xlPasteAll);

                        LinkNamedRanges(sheet, "Auto_Sampler_Timepoints" + i + 1, "SummaryTimepoints" + i);
                        LinkNamedRanges(sheet, "Sampler1_Mean" + i, "SummaryManualMeans" + i);
                        LinkNamedRanges(sheet, "Sampler2_Mean" + i, "SummaryAutoMeans" + i);
                        SetSamplerHeaderFormula(sheet, "Sample_Name" + i, "I", 2, i);
                        SetSamplerHeaderFormula(sheet, "Sample_Name" + i, "J", timepoints + 7, i);
                        SetSummaryCellFormula(sheet, "Summary_Table" + i, i);

                        if (((i-1) % 4) == 0)
                        {
                            int newHeaderColIndex = headerColIndex + 1;
                            InsertSummaryHeaderBeforeSummary(sheet, timepoints, "Summary_Table" + i, "Sampling_Summary_Header" + headerColIndex, "Sampling_Summary_Header1" + newHeaderColIndex);
                            headerColIndex++;
                        }
                    }
                }

                try
                {
                    _app.Goto(sheet.Cells[1, 1], true);
                }
                catch
                {
                    Logger.LogMessage("Scroll of sheet failed in AlternateAutomatedSampling.UpdateAlternateAutomatedSamplingSheet", Level.Error);
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

        private static void SetSamplerHeaderFormula(Worksheet sheet, string tableName, string samplerCell, int cellIndex, int setIndex)
        {
            Name nm = sheet.Names.Item(tableName, Type.Missing, Type.Missing) as Name;
            if (nm == null) return;

            Range rng = nm.RefersToRange;
            if (rng == null) return;

            int headerRow = rng.Row + cellIndex;
            int targetRow = 12 + (setIndex - 1);
            Range cell = sheet.Cells[headerRow, 3];
            cell.Formula = $"=CONCATENATE(K{targetRow}, \", \", {samplerCell}{targetRow})";

            WorksheetUtilities.ReleaseComObject(cell);
            WorksheetUtilities.ReleaseComObject(rng);
            WorksheetUtilities.ReleaseComObject(nm);
        }

        private static void SetSummaryCellFormula(Worksheet sheet, string tableName, int setIndex)
        {
            Name nm = sheet.Names.Item(tableName, Type.Missing, Type.Missing) as Name;
            if (nm == null) return;

            Range rng = nm.RefersToRange;
            if (rng == null) return;

            int headerRow = rng.Row - 1;
            int targetRow = 12 + (setIndex - 1);
            Range cell = sheet.Cells[headerRow, 2];
            cell.Formula = $"=CONCATENATE(K{targetRow}, \", \", I{targetRow}, \". \", J{targetRow})";

            WorksheetUtilities.ReleaseComObject(cell);
            WorksheetUtilities.ReleaseComObject(rng);
            WorksheetUtilities.ReleaseComObject(nm);
        }

        private static void InsertSummaryHeaderBeforeSummary(
            Worksheet sheet,
            int timePoints,
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

            int insertRow = ssRange.Row -1;
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

        private static void SetSamplingSummaryFormula(Worksheet worksheet, int i)
        {
            string samplingSummaryRangeName = $"Summary_Table{i}";

            Range samplingSummaryRange = worksheet.Range[samplingSummaryRangeName];
            Range firstCell = samplingSummaryRange.Cells[1, 2] as Range;

            if (firstCell != null)
            {
                //firstCell.Formula = $"=CONCATENATE({strengthName},\" (Batch#\", {batchNumName}, \")\")";
                firstCell.Formula = $"=Sample_Name{i}";
                WorksheetUtilities.ReleaseComObject(firstCell);
            }

            WorksheetUtilities.ReleaseComObject(samplingSummaryRange);
        }

        private static void CopySummaryRangeAndLink(
        Worksheet sheet,
        string srcName,
        string destName,
        int rowIndex,
        int colIndex,
        string linkSourceName = null)
        {

            var srcNameObj = sheet.Names.Item(srcName, Type.Missing, Type.Missing) as Name;
            if (srcNameObj == null) return;

            Range srcRange = srcNameObj.RefersToRange;
            if (srcRange == null) return;

            Range destStart = sheet.Cells[rowIndex, colIndex];
            Range destEnd = sheet.Cells[rowIndex + srcRange.Rows.Count - 1,
                                        colIndex + srcRange.Columns.Count - 1];
            Range destBlock = sheet.Range[destStart, destEnd];

            srcRange.Copy(destBlock);


            if (!string.IsNullOrEmpty(destName))
            {
                string refersToLocal = "='" + sheet.Name + "'!" +
                    destBlock.get_AddressLocal(true, true, XlReferenceStyle.xlA1);
                sheet.Names.Add(destName, refersToLocal);
            }


            if (!string.IsNullOrEmpty(linkSourceName))
            {
                Range linkSourceRange = sheet.Range[linkSourceName];
                int rowCount = Math.Min(destBlock.Rows.Count, linkSourceRange.Rows.Count);

                for (int i = 1; i <= rowCount; i++)
                {
                    Range srcCell = linkSourceRange.Cells[i, 1] as Range;
                    Range destCell = destBlock.Cells[i, 1] as Range;

                    if (srcCell != null && destCell != null)
                    {
                        string srcAddress = srcCell.get_Address(false, false, XlReferenceStyle.xlA1);
                        destCell.Formula = $"=IF({srcAddress}=\"\",\"\",{srcAddress})";

                        WorksheetUtilities.ReleaseComObject(srcCell);
                        WorksheetUtilities.ReleaseComObject(destCell);
                    }
                }

                WorksheetUtilities.ReleaseComObject(linkSourceRange);
            }

            WorksheetUtilities.ReleaseComObject(srcRange);
            WorksheetUtilities.ReleaseComObject(destBlock);
        }
    }

}
