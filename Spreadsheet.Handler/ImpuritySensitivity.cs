using log4net.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;


namespace Spreadsheet.Handler
{
    public class ImpuritySensitivity
    {
        private static Application _app;

        private const int DefaulttbNumOfPeaks = 1;
        private const int DefaulttbNumOfInjections = 6;
        private const int DefaultNumImpurities = 1;
        private const int MaxNumImpurities = 8;
        private const string TempDirectoryName = "ABD_TempFiles";

        // This method is mere;y provided for backwards compatibility within the DEV environment (for old documents).
        public static string UpdateImpSensitivitySheet(string sourcePath, int tbNumOfPeaks, bool isChkRL_SN, bool isChkRL_RSD, bool isChkDL_SN,
            Dictionary<string, int> numParams = null,
            Dictionary<string, string> acceptanceCriteria = null)
        {
            string returnPath = "";
            try
            {
                returnPath = UpdateImpSensitivitySheet2(sourcePath, tbNumOfPeaks, isChkRL_SN, isChkRL_RSD, isChkDL_SN, numParams, acceptanceCriteria);
            }
            catch (Exception ex)
            {
                Logger.LogMessage("An error occurred in the call to ImpuritySensitivity.UpdateImpSensitivitySheet. Message and stack trace are:\r\n" + ex.Message + "\r\n" + ex.StackTrace, Level.Error);

                try
                {
                    if (_app.Workbooks.Count > 0)
                    {
                        try
                        {
                            _app.Workbooks[0].Save();
                            returnPath = _app.Workbooks[0].FullName;
                        }
                        catch (Exception exce)
                        {
                            Logger.LogMessage("An error occurred in the call to ImpuritySensitivity.UpdateImpSensitivitySheet. Failed to save current workbook changes and to get path.", Level.Error);
                        }

                        _app.Workbooks.Close();
                    }
                    _app = null;
                }
                catch
                {
                    Logger.LogMessage("An error occurred in the call to ImpuritySensitivity.UpdateImpSensitivitySheet. Application failed to close workbooks. Message and stack trace are:\r\n" + ex.Message + "\r\n" + ex.StackTrace, Level.Error);
                }
                finally
                {
                    WorksheetUtilities.ReleaseExcelApp();
                }
            }
            return returnPath;
        }

        private static string UpdateImpSensitivitySheet2(string sourcePath, int tbNumOfPeaks, bool isChkRL_SN, bool isChkRL_RSD, bool isChkDL_SN,
            Dictionary<string, int> numParams = null,
            Dictionary<string, string> acceptanceCriteria = null)
        {
            if (!File.Exists(sourcePath))
            {
                Logger.LogMessage("Error in call to ImpuritySensitivity.UpdateImpSensitivitySheet. Invalid source file path specified.", Level.Error);
                return "";
            }

            // Generate an random temp path to save new workbook
            string savePath = WorksheetUtilities.CopyWorkbook(sourcePath, TempDirectoryName, "Impurity Sensitivity Results.xls");
            if (String.IsNullOrEmpty(savePath)) return "";

            // Try to open the file
            _app = WorksheetUtilities.GetExcelApp();
            _app.Workbooks.Open(savePath, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            Workbook book = _app.Workbooks[1];
            Worksheet sheet = book.Worksheets[1] as Worksheet;

            if (sheet != null)
            {
                bool wasProtected = WorksheetUtilities.SetSheetProtection(sheet, null, false);
                //As of 02-23, always delete the hidden section of the spreadsheet.
                //WorksheetUtilities.DeleteNamedRangeRows(sheet, "ToDelete");

                int offset = 0;
                int numInjectionsRl = 6;
                int numInjectionsDl = 6;
                var isCalRlDl = false;
                var isValTypeNda = false;
                var cmbRlRsd = "";
                var cmbRlSn = "";
                var cmbDlSn = "";
                var cmbValidation = "";
                var cmbProduct = "";
                var cmbQuantitativeType = "";

                if (numParams != null && numParams.Count > 0)
                {
                    if (numParams.ContainsKey("txtNumReps"))
                        numInjectionsRl = numParams["txtNumReps"];

                    if (numParams.ContainsKey("txtNumRepsDL"))
                        numInjectionsDl = numParams["txtNumRepsDL"];
                }


                if (acceptanceCriteria != null && acceptanceCriteria.Count > 0)
                {
                    if (acceptanceCriteria.ContainsKey("cmbValidation"))
                    {
                        cmbValidation = acceptanceCriteria["cmbValidation"];
                        isValTypeNda = cmbValidation == "NDA";
                        sheet.Cells[2, 4] = cmbValidation;
                    }

                    if (acceptanceCriteria.ContainsKey("cmbProduct"))
                    {
                        cmbProduct = acceptanceCriteria["cmbProduct"];
                        sheet.Cells[2, 6] = cmbProduct;
                    }

                    if (acceptanceCriteria.ContainsKey("cmbQuantitativeType"))
                    {
                        cmbQuantitativeType = acceptanceCriteria["cmbQuantitativeType"];
                        sheet.Cells[2, 8] = cmbQuantitativeType;
                    }

                    if (acceptanceCriteria.ContainsKey("cmbRLDL"))
                    {
                        isCalRlDl = acceptanceCriteria["cmbRLDL"] == "Yes";
                    }

                    if (acceptanceCriteria.ContainsKey("cmbRL_RSD"))
                    {
                        cmbRlRsd = acceptanceCriteria["cmbRL_RSD"];
                        sheet.Cells[4, 3] = cmbRlRsd;
                    }
                    if (acceptanceCriteria.ContainsKey("TBRL_RSD"))
                    {
                        sheet.Cells[4, 4] = acceptanceCriteria["TBRL_RSD"];
                    }
                    if (acceptanceCriteria.ContainsKey("cmbRL_SN"))
                    {
                        cmbRlSn = acceptanceCriteria["cmbRL_SN"];
                        sheet.Cells[6, 3] = cmbRlSn;
                    }
                    if (acceptanceCriteria.ContainsKey("TBRL_SN"))
                    {
                        sheet.Cells[6, 4] = acceptanceCriteria["TBRL_SN"];
                    }
                    if (acceptanceCriteria.ContainsKey("cmbDL_SN"))
                    {
                        cmbDlSn = acceptanceCriteria["cmbDL_SN"];

                        sheet.Cells[8, 3] = cmbDlSn;
                    }
                    if (acceptanceCriteria.ContainsKey("TBDL_SN"))
                    {
                        sheet.Cells[8, 4] = acceptanceCriteria["TBDL_SN"];
                    }
                }



                if (tbNumOfPeaks != DefaulttbNumOfPeaks)
                {
                    if (tbNumOfPeaks > DefaulttbNumOfPeaks)
                    {
                        offset = tbNumOfPeaks - DefaulttbNumOfPeaks;
                    }
                    else
                    {
                        offset = -(DefaulttbNumOfPeaks - tbNumOfPeaks);
                    }
                }
                AdjustInjectionRowsByNamedRange(sheet, "SampleNumsRawData", numInjectionsRl);
                AdjustInjectionRowsByNamedRange(sheet, "SampleNumsRawData2", numInjectionsRl);
                AdjustInjectionRowsByNamedRange(sheet, "SampleRawDataDL", numInjectionsDl);

                if (tbNumOfPeaks > DefaulttbNumOfPeaks)
                {
                    int numImpuritiesToInsert = tbNumOfPeaks - DefaulttbNumOfPeaks;
                    WorksheetUtilities.InsertColumnsIntoNamedRange(numImpuritiesToInsert - MaxNumImpurities, sheet, "ImpurityResults", XlDirection.xlToRight);

                    for (int i = 1; i <= numImpuritiesToInsert; i++)
                    {
                        // Copy the named ranges as needed for each impurity
                        int namedRangeNum = i + 1;
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ImpurityResults" + i, "ImpurityResults" + namedRangeNum, 1, 3, XlPasteType.xlPasteAll);
                        WorksheetUtilities.SetNamedRangeValue(sheet, "ImpurityResults" + namedRangeNum, "Impurity " + namedRangeNum, 2, 2);
                        //Added as 12-2022
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "SignalToNoiseResults" + i, "SignalToNoiseResults" + namedRangeNum, 1, 4, XlPasteType.xlPasteAll);
                        WorksheetUtilities.SetNamedRangeValue(sheet, "SignalToNoiseResults" + namedRangeNum, "S/N", 2, 2);
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "SignalToNoiseResultsDL" + i, "SignalToNoiseResultsDL" + namedRangeNum, 1, 4, XlPasteType.xlPasteAll);
                        WorksheetUtilities.SetNamedRangeValue(sheet, "SignalToNoiseResultsDL" + namedRangeNum, "S/N", 2, 2);

                    }

                    AddSummaryTable1RowsWithFormulaUpdate(sheet, tbNumOfPeaks);
                    if (isCalRlDl)
                    {
                        AddSummaryTable2RowsWithFormulaUpdate(sheet, tbNumOfPeaks);
                    }
                    UpdateImpurityResultsSummaryFormulas(sheet, tbNumOfPeaks, cmbRlRsd);
                    if (isChkRL_SN)
                    {
                        UpdateSignalToNoiseRlFormulas(sheet, tbNumOfPeaks, numInjectionsRl, cmbRlSn);
                        AddSummaryTable3RowsWithFormulaUpdate(sheet, tbNumOfPeaks);

                    }
                    if (isChkDL_SN)
                    {
                        UpdateSignalToNoiseDlFormulas(sheet, tbNumOfPeaks, numInjectionsDl, cmbDlSn);
                        AddSummaryTable4RowsWithFormulaUpdate(sheet, tbNumOfPeaks);
                    }
                }

                ExtendSummaryTable1Injections(sheet, tbNumOfPeaks, numInjectionsRl);

                if (isChkRL_SN)
                    ExtendSummaryTable3Injections(sheet, tbNumOfPeaks, numInjectionsRl);
                if (isChkDL_SN)
                    ExtendSummaryTable4Injections(sheet, tbNumOfPeaks, numInjectionsDl);

                if (!((isChkRL_SN || isChkRL_RSD) && isCalRlDl))
                {
                    WorksheetUtilities.DeleteNamedRangeRows(sheet, "CalculationForRL");
                    WorksheetUtilities.DeleteNamedRangeRows(sheet, "SummaryTable2");
                }

                if (!isChkRL_SN)
                {
                    WorksheetUtilities.DeleteNamedRangeRows(sheet, "Signal_to_Noise_Ratio__RL");
                    WorksheetUtilities.DeleteNamedRangeRows(sheet, "SummaryTable3");
                    WorksheetUtilities.DeleteNamedRangeRows(sheet, "AC_RL_SN");
                }
                else
                {
                    ApplySignalToNoiseRlStyling(sheet, numInjectionsRl, tbNumOfPeaks);
                    ApplySummaryTableWithMinMaxStyling(sheet, "SummaryTable3");
                }

                if (!isChkDL_SN)
                {
                    WorksheetUtilities.DeleteNamedRangeRows(sheet, "Signal_to_Noise_Ratio__DL");
                    WorksheetUtilities.DeleteNamedRangeRows(sheet, "SummaryTable4");
                    WorksheetUtilities.DeleteNamedRangeRows(sheet, "AC_DL_SN");
                }
                else
                {
                    ApplySignalToNoiseDlStyling(sheet, numInjectionsDl, tbNumOfPeaks);
                    ApplySummaryTableWithMinMaxStyling(sheet, "SummaryTable4");
                }

                if (!isValTypeNda)
                    WorksheetUtilities.DeleteNamedRangeRows(sheet, "SummaryTable1");
                else
                    ApplySummaryTableStyling(sheet, "SummaryTable1");

                if (!isChkRL_RSD)
                    WorksheetUtilities.DeleteNamedRangeRows(sheet, "AC_RL_RSD");

                ApplyImpurityResultsStyling(sheet, numInjectionsRl, tbNumOfPeaks);

                try
                {
                    _app.Goto(sheet.Cells[1, 1], true);
                }
                catch
                {
                    Logger.LogMessage("Scroll of sheet failed in ImpuritySensitivity.UpdateImpSensitivitySheet!", Level.Error);
                }

                if (wasProtected) WorksheetUtilities.SetSheetProtection(sheet, null, true);

                while (Marshal.ReleaseComObject(sheet) >= 0) { }
            }

            _app.Workbooks[1].Save();

            while (Marshal.ReleaseComObject(book) >= 0) { }
            _app.Workbooks.Close();

            //while (Marshal.ReleaseComObject(_app) >= 0) { }
            _app = null;
            WorksheetUtilities.ReleaseExcelApp();

            return savePath;
        }

        private static void ReleaseComObject(object obj)
        {
            if (obj != null && Marshal.IsComObject(obj))
            {
                while (Marshal.ReleaseComObject(obj) > 0) { }
            }
        }

        private static string GetExcelColumnLetter(int colNumber)
        {
            string colLetter = "";
            while (colNumber > 0)
            {
                int mod = (colNumber - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                colNumber = (colNumber - mod - 1) / 26;
            }
            return colLetter;
        }

        private static void AddSummaryTable1RowsWithFormulaUpdate(_Worksheet sheet, int tbNumOfPeaks)
        {
            if (tbNumOfPeaks <= 1) return;

            var templateRange = sheet.Names.Item("SummaryTable1Row1").RefersToRange;
            int startRow = templateRange.Row;
            int startCol = templateRange.Column;
            int colCount = templateRange.Columns.Count;

            //var injTemplateRange = sheet.Names.Item("SummaryTable1Inj2").RefersToRange;
            //int colCount = injTemplateRange.Column;

            int insertRow = startRow + 1;

            for (int i = 2; i <= tbNumOfPeaks; i++)
            {
                sheet.Rows[insertRow].Insert();


                for (int col = 1; col <= colCount; col++)
                {
                    var srcCell = templateRange.Cells[1, col];
                    var tgtCell = sheet.Cells[insertRow, startCol + col - 1];

                    string formula = srcCell.Formula;

                    if (!string.IsNullOrEmpty(formula))
                    {

                        formula = UpdateFormulaReferenceColumn(formula, 4, (i - 1) * 2); // Column D = 4
                        tgtCell.Formula = formula;
                    }
                }

                string newRowName = $"SummaryTable1Row{i}";
                Range start = sheet.Cells[insertRow, startCol];
                Range end = sheet.Cells[insertRow, startCol + colCount - 1];
                Range newRange = sheet.Range[start, end];

                string refersToLocal = $"='{sheet.Name}'!{newRange.get_AddressLocal(true, true, XlReferenceStyle.xlA1)}";
                sheet.Names.Add(newRowName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, refersToLocal, Type.Missing, Type.Missing, Type.Missing);

                insertRow++;
            }
        }

        private static void AddSummaryTable2RowsWithFormulaUpdate(_Worksheet sheet, int tbNumOfPeaks)
        {
            if (tbNumOfPeaks <= 1) return;

            var templateRange = sheet.Names.Item("SummaryTable2Row1").RefersToRange;
            int startRow = templateRange.Row;
            int startCol = templateRange.Column;
            int colCount = templateRange.Columns.Count;


            for (int i = 2; i <= tbNumOfPeaks; i++)
            {
                int insertRow = startRow + i - 1;

                sheet.Rows[insertRow].Insert();


                for (int col = 1; col <= colCount; col++)
                {
                    var srcCell = templateRange.Cells[1, col];
                    var tgtCell = sheet.Cells[insertRow, startCol + col - 1];

                    string formula = srcCell.Formula;

                    if (!string.IsNullOrEmpty(formula))
                    {

                        formula = UpdateFormulaReferenceColumn(formula, 4, (i - 1) * 2); // Column D = 4
                        tgtCell.Formula = formula;
                    }
                }
                //var rowRange = sheet.Names.Item($"{baseRowName}{i}").RefersToRange;


                string newRowName = $"SummaryTable2Row{i}";
                Range start = sheet.Cells[insertRow, startCol];
                Range end = sheet.Cells[insertRow, startCol + colCount - 1];
                Range newRange = sheet.Range[start, end];
                newRange.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlLineStyleNone;

                string refersToLocal = $"='{sheet.Name}'!{newRange.get_AddressLocal(true, true, XlReferenceStyle.xlA1)}";
                sheet.Names.Add(newRowName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, refersToLocal, Type.Missing, Type.Missing, Type.Missing);

                insertRow++;
            }

            var lastRowRange = sheet.Names.Item($"SummaryTable2Row{tbNumOfPeaks}").RefersToRange;
            var lastRowIdx = lastRowRange.Row;
            var lastRowFullRange = sheet.Range[sheet.Cells[lastRowIdx, 2], sheet.Cells[lastRowIdx, 9]];
            lastRowFullRange.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
        }

        private static void AddSummaryTable3RowsWithFormulaUpdate(_Worksheet sheet, int tbNumOfPeaks)
        {
            if (tbNumOfPeaks <= 1) return;

            var templateRange = sheet.Names.Item("SummaryTable3Row1").RefersToRange;
            int startRow = templateRange.Row;
            int startCol = templateRange.Column;
            int colCount = templateRange.Columns.Count;

            int insertRow = startRow + 1;

            for (int i = 2; i <= tbNumOfPeaks; i++)
            {
                sheet.Rows[insertRow].Insert();


                for (int col = 1; col <= colCount; col++)
                {
                    var srcCell = templateRange.Cells[1, col];
                    var tgtCell = sheet.Cells[insertRow, startCol + col - 1];

                    string formula = srcCell.Formula;

                    if (!string.IsNullOrEmpty(formula))
                    {

                        formula = UpdateFormulaReferenceColumn(formula, 4, (i - 1) * 3);
                        tgtCell.Formula = formula;
                    }
                }

                string newRowName = $"SummaryTable3Row{i}";
                Range start = sheet.Cells[insertRow, startCol];
                Range end = sheet.Cells[insertRow, startCol + colCount - 1];
                Range newRange = sheet.Range[start, end];

                string refersToLocal = $"='{sheet.Name}'!{newRange.get_AddressLocal(true, true, XlReferenceStyle.xlA1)}";
                sheet.Names.Add(newRowName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, refersToLocal, Type.Missing, Type.Missing, Type.Missing);

                insertRow++;
            }
        }

        private static void AddSummaryTable4RowsWithFormulaUpdate(_Worksheet sheet, int tbNumOfPeaks)
        {
            if (tbNumOfPeaks <= 1) return;

            var templateRange = sheet.Names.Item("SummaryTable4Row1").RefersToRange;
            int startRow = templateRange.Row;
            int startCol = templateRange.Column;
            int colCount = templateRange.Columns.Count;

            int insertRow = startRow + 1;

            for (int i = 2; i <= tbNumOfPeaks; i++)
            {
                sheet.Rows[insertRow].Insert();


                for (int col = 1; col <= colCount; col++)
                {
                    var srcCell = templateRange.Cells[1, col];
                    var tgtCell = sheet.Cells[insertRow, startCol + col - 1];

                    string formula = srcCell.Formula;

                    if (!string.IsNullOrEmpty(formula))
                    {

                        formula = UpdateFormulaReferenceColumn(formula, 4, (i - 1) * 3); // Column D = 4
                        tgtCell.Formula = formula;
                    }
                }

                string newRowName = $"SummaryTable4Row{i}";
                Range start = sheet.Cells[insertRow, startCol];
                Range end = sheet.Cells[insertRow, startCol + colCount - 1];
                Range newRange = sheet.Range[start, end];

                string refersToLocal = $"='{sheet.Name}'!{newRange.get_AddressLocal(true, true, XlReferenceStyle.xlA1)}";
                sheet.Names.Add(newRowName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, refersToLocal, Type.Missing, Type.Missing, Type.Missing);

                insertRow++;
            }
        }

        private static void ExtendSummaryTable1Injections(_Worksheet sheet, int tbNumOfPeaks, int numInjectionsRL)
        {
            var summaryTableRange = sheet.Names.Item("SummaryTable1").RefersToRange;
            const string injPrefix = "SummaryTable1Inj";
            int topRow = summaryTableRange.Row;
            int labelRow = topRow + 1;
            int formulaRow = topRow + 2;
            int startCol = summaryTableRange.Column;
            int currentColCount = summaryTableRange.Columns.Count;

            int baseInjectionOffset = 2;
            int currentInjectionCount = currentColCount - baseInjectionOffset;
            int rowCount = summaryTableRange.Rows.Count;
            for (int injNum = currentInjectionCount; injNum > 1; injNum--)
            {
                if (injNum > numInjectionsRL)
                {
                    string injName = $"{injPrefix}{injNum}";
                    try
                    {
                        var injRange = sheet.Names.Item(injName).RefersToRange;
                        int injCol = injRange.Column;
                        int injRow = injRange.Row;
                        for (int row = injRow; row <= injRow + tbNumOfPeaks + 1; row++)
                        {
                            sheet.Cells[row, injCol].Delete(XlDeleteShiftDirection.xlShiftToLeft);
                        }
                        sheet.Names.Item(injName).Delete();
                    }
                    catch { }
                }
            }

            currentColCount = sheet.Names.Item("SummaryTable1Row1").RefersToRange.Columns.Count;
            currentInjectionCount = currentColCount - baseInjectionOffset;
            if (numInjectionsRL <= currentInjectionCount) return;

            int colsToAdd = numInjectionsRL - currentInjectionCount;

            for (int i = 1; i <= colsToAdd; i++)
            {
                int injectionNumber = currentInjectionCount + i;
                int newColIndex = startCol + baseInjectionOffset + injectionNumber - 1;

                Range insertColRange = sheet.Range[
                    sheet.Cells[topRow, newColIndex],
                    sheet.Cells[topRow + 2, newColIndex]
                ];
                insertColRange.Insert(Type.Missing, XlInsertShiftDirection.xlShiftToRight);

                sheet.Cells[labelRow, newColIndex].Value2 = $"Inj{injectionNumber}";

                int targetDRow = 12 + injectionNumber;
                string formula = $"=D{targetDRow}";
                sheet.Cells[formulaRow, newColIndex].Formula = formula;

                var sourceCell = sheet.Cells[formulaRow, newColIndex - 1];
                var targetCell = sheet.Cells[formulaRow, newColIndex];
                CopyCellFormatting(sourceCell, targetCell);

                for (int rowOffset = 1; ; rowOffset++)
                {
                    string rowName = $"SummaryTable1Row{rowOffset + 1}";
                    try
                    {
                        var rowRange = sheet.Names.Item(rowName).RefersToRange;
                        int absRow = rowRange.Row;

                        string colLetter = GetExcelColumnLetter(4 + (rowOffset * 2));
                        var cell = sheet.Cells[absRow, newColIndex];
                        cell.Formula = $"={colLetter}{targetDRow}";
                        CopyCellFormatting(sourceCell, cell);

                        int newColCount = newColIndex - startCol + 1;
                        Range extendedRange = sheet.Range[
                            sheet.Cells[absRow, startCol],
                            sheet.Cells[absRow, startCol + newColCount - 1]
                        ];
                        string refersTo = $"='{sheet.Name}'!{extendedRange.get_AddressLocal(true, true)}";
                        sheet.Names.Item(rowName).Delete();
                        sheet.Names.Add(rowName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                        Type.Missing, Type.Missing, refersTo, Type.Missing, Type.Missing, Type.Missing);
                    }
                    catch { break; }
                }

                try
                {
                    Range injNamedRange = sheet.Range[
                        sheet.Cells[topRow, newColIndex],
                        sheet.Cells[topRow + 2, newColIndex]
                    ];
                    string injRef = $"='{sheet.Name}'!{injNamedRange.get_AddressLocal(true, true)}";
                    string newInjName = $"SummaryTable1Inj{injectionNumber}";
                    sheet.Names.Add(newInjName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing, injRef, Type.Missing, Type.Missing, Type.Missing);
                }
                catch { }
            }
        }

        private static void ExtendSummaryTable1InjectionsNew(_Worksheet sheet, int numInjectionsRL)
        {
            var summaryTableRange = sheet.Names.Item("SummaryTable1").RefersToRange;
            int topRow = summaryTableRange.Row;
            int startCol = summaryTableRange.Column;
            int totalRows = summaryTableRange.Rows.Count;
            int currentColCount = summaryTableRange.Columns.Count;

            int labelRow = topRow + 1;
            int formulaRow = topRow + 2;

            int baseInjectionOffset = 2;

            int currentInjectionCount = currentColCount - baseInjectionOffset;

            if (currentInjectionCount > 1 && currentInjectionCount > numInjectionsRL)
            {
                for (int injNum = currentInjectionCount; injNum > numInjectionsRL; injNum--)
                {
                    int localColInRange = baseInjectionOffset + injNum;
                    try
                    {
                        Range colInsideTable = summaryTableRange.Columns[localColInRange];
                        colInsideTable.Delete(XlDeleteShiftDirection.xlShiftToLeft);
                    }
                    catch { }

                    string injName = $"SummaryTable1Inj{injNum}";
                    try { sheet.Names.Item(injName).Delete(); } catch { }
                }

                summaryTableRange = sheet.Names.Item("SummaryTable1").RefersToRange;
                currentColCount = summaryTableRange.Columns.Count;
                currentInjectionCount = currentColCount - baseInjectionOffset;
            }

            //if (numInjectionsRL < currentInjectionCount)
            //{

            //    try
            //    {
            //        for (int inj = 1; inj <= currentInjectionCount; inj++)
            //        {
            //            int absCol = startCol + baseInjectionOffset + inj - 1;
            //            Range injNamedRange = sheet.Range[
            //                sheet.Cells[topRow, absCol],
            //                sheet.Cells[topRow + (totalRows - 1), absCol]
            //            ];
            //            string injRef = $"='{sheet.Name}'!{injNamedRange.get_AddressLocal(true, true)}";
            //            string injName = $"SummaryTable1Inj{inj}";
            //            try { sheet.Names.Item(injName).Delete(); } catch { }
            //            sheet.Names.Add(injName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            //                            Type.Missing, Type.Missing, injRef, Type.Missing, Type.Missing, Type.Missing);
            //        }
            //    }
            //    catch { }
            //    return;
            //}

            //int tableWidth = baseInjectionOffset + numInjectionsRL;
            //int tableRightCol = startCol + summaryTableRange.Columns.Count - 1;

            //for(int col = startCol + tableWidth; col <= tableRightCol; col++)
            //{
            //    Range clearRange = sheet.Range[
            //        sheet.Cells[topRow, col],
            //        sheet.Cells[topRow + summaryTableRange.Rows.Count - 1, col]
            //        ];
            //    clearRange.ClearFormats();
            //}

            //int colsToAdd = numInjectionsRL - currentInjectionCount;

            //for (int i = 1; i <= colsToAdd; i++)
            //{
            //    int injectionNumber = currentInjectionCount + i;

            //    int localColToInsert = baseInjectionOffset + injectionNumber;
            //    int absInsertCol = startCol + localColToInsert - 1;

            //    Range insertColSlice = sheet.Range[
            //        sheet.Cells[topRow, absInsertCol],
            //        sheet.Cells[topRow + (totalRows - 1), absInsertCol]
            //    ];
            //    insertColSlice.Insert(XlInsertShiftDirection.xlShiftToRight);

            //    int absCopyCol = absInsertCol - 1;
            //    var srcFmtTop = sheet.Cells[topRow, absCopyCol];
            //    var srcFmtMid = sheet.Cells[formulaRow, absCopyCol];


            //    sheet.Cells[labelRow, absInsertCol].Value2 = $"Inj{injectionNumber}";

            //    int targetDRow = 12 + injectionNumber;
            //    sheet.Cells[formulaRow, absInsertCol].Formula = $"=D{targetDRow}";

            //    CopyCellFormatting(srcFmtTop, sheet.Cells[labelRow, absInsertCol]);
            //    CopyCellFormatting(srcFmtMid, sheet.Cells[formulaRow, absInsertCol]);

            //    for (int rowOffset = 1; ; rowOffset++)
            //    {
            //        string rowName = $"SummaryTable1Row{rowOffset + 1}";
            //        Range rowRange;
            //        try
            //        {
            //            rowRange = sheet.Names.Item(rowName).RefersToRange;
            //        }
            //        catch { break; }

            //        string colLetter = GetExcelColumnLetter(4 + (rowOffset * 2));

            //        var cell = sheet.Cells[rowRange.Row, absInsertCol];
            //        cell.Formula = $"={colLetter}{targetDRow}";

            //        CopyCellFormatting(sheet.Cells[rowRange.Row, absInsertCol - 1], cell);

            //        int newColCountForRow = rowRange.Columns.Count + 1;
            //        Range extended = sheet.Range[
            //            sheet.Cells[rowRange.Row, startCol],
            //            sheet.Cells[rowRange.Row, startCol + newColCountForRow - 1]
            //        ];
            //        string refersTo = $"='{sheet.Name}'!{extended.get_AddressLocal(true, true)}";
            //        sheet.Names.Item(rowName).Delete();
            //        sheet.Names.Add(rowName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            //                        Type.Missing, Type.Missing, refersTo, Type.Missing, Type.Missing, Type.Missing);
            //    }

            //    try
            //    {
            //        Range injNamedRange = sheet.Range[
            //            sheet.Cells[topRow, absInsertCol],
            //            sheet.Cells[topRow + (totalRows - 1), absInsertCol]
            //        ];
            //        string injRef = $"='{sheet.Name}'!{injNamedRange.get_AddressLocal(true, true)}";
            //        string newInjName = $"SummaryTable1Inj{injectionNumber}";
            //        sheet.Names.Add(newInjName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            //                        Type.Missing, Type.Missing, injRef, Type.Missing, Type.Missing, Type.Missing);
            //    }
            //    catch { }
            //}

            //try
            //{
            //    summaryTableRange = sheet.Names.Item("SummaryTable1").RefersToRange;
            //    totalRows = summaryTableRange.Rows.Count;
            //    for (int inj = 1; inj <= numInjectionsRL; inj++)
            //    {
            //        int absCol = startCol + baseInjectionOffset + inj - 1;
            //        Range injNamedRange = sheet.Range[
            //            sheet.Cells[topRow, absCol],
            //            sheet.Cells[topRow + (totalRows - 1), absCol]
            //        ];
            //        string injRef = $"='{sheet.Name}'!{injNamedRange.get_AddressLocal(true, true)}";
            //        string injName = $"SummaryTable1Inj{inj}";
            //        try { sheet.Names.Item(injName).Delete(); } catch { }
            //        sheet.Names.Add(injName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            //                        Type.Missing, Type.Missing, injRef, Type.Missing, Type.Missing, Type.Missing);
            //    }
            //}
            //catch { }
        }
        private static void CopyCellFormatting(Range source, Range target)
        {
            target.Borders.LineStyle = source.Borders.LineStyle;
            target.Font.Bold = source.Font.Bold;
            target.Font.Size = source.Font.Size;
            target.HorizontalAlignment = source.HorizontalAlignment;
            target.VerticalAlignment = source.VerticalAlignment;
            target.Interior.Color = source.Interior.Color;
            target.NumberFormat = source.NumberFormat;
        }

        private static void ExtendSummaryTable3Injections(_Worksheet sheet, int tbNumOfPeaks, int numInjectionsRL)
        {
            const string tableRowPrefix = "SummaryTable3Row";
            const string injPrefix = "SummaryTable3Inj";
            const string tableRow1Name = "SummaryTable3Row1";
            const int baseInjection = 6;
            int signalStartRow = sheet.Names.Item("SignalToNoiseResults1").RefersToRange.Row;

            var row1 = sheet.Names.Item(tableRow1Name).RefersToRange;
            int baseRow = row1.Row;
            int startCol = row1.Column;

            for (int injNum = baseInjection; injNum > 1; injNum--)
            {
                if (injNum > numInjectionsRL)
                {
                    string injName = $"{injPrefix}{injNum}";
                    try
                    {
                        var injRange = sheet.Names.Item(injName).RefersToRange;
                        int injCol = injRange.Column;
                        int injRow = injRange.Row;
                        for (int row = injRow; row <= injRow + tbNumOfPeaks + 1; row++)
                        {
                            sheet.Cells[row, injCol].Delete(XlDeleteShiftDirection.xlShiftToLeft);
                        }
                        sheet.Names.Item(injName).Delete();
                    }
                    catch { }
                }
            }

            for (int injNum = 1; injNum <= Math.Min(numInjectionsRL, baseInjection); injNum++)
            {
                string injName = $"{injPrefix}{injNum}";
                try
                {
                    var injRange = sheet.Names.Item(injName).RefersToRange;
                    int injCol = injRange.Column;

                    for (int rowNum = 1; ; rowNum++)
                    {
                        string rowName = $"{tableRowPrefix}{rowNum}";
                        Range rowRange;
                        try { rowRange = sheet.Names.Item(rowName).RefersToRange; } catch { break; }

                        int absRow = rowRange.Row;
                        int signalRow = signalStartRow + 1 + injNum;
                        int scaleRow = baseRow - 2;
                        string signalColLetter = GetExcelColumnLetter(4 + (rowNum - 1) * 3);
                        string injColLetter = GetExcelColumnLetter(injCol);

                        var cell = sheet.Cells[absRow, injCol];
                        string formula = $"=IF({signalColLetter}{signalRow}=\"\",\"\",FIXED({signalColLetter}{signalRow}, {injColLetter}{scaleRow}))";
                        cell.Formula = formula;
                    }
                }
                catch { continue; }
            }

            row1 = sheet.Names.Item(tableRow1Name).RefersToRange;
            int totalCols = row1.Columns.Count;
            if (numInjectionsRL <= baseInjection) return;

            int colsToAdd = numInjectionsRL - baseInjection;
            int rowCount = 0;
            for (int i = 1; ; i++)
            {
                try { var _ = sheet.Names.Item($"{tableRowPrefix}{i}").RefersToRange; rowCount++; } catch { break; }
            }

            int insertBeforeCol = FindColumnIndex(sheet, baseRow - 1, "Minimum");

            for (int i = 1; i <= colsToAdd; i++)
            {
                int newInjNum = baseInjection + i;
                int insertAtCol = insertBeforeCol + i - 1;

                Range insertRange = sheet.Range[
                    sheet.Cells[baseRow - 1, insertAtCol],
                    sheet.Cells[baseRow - 1 + rowCount, insertAtCol]
                ];
                insertRange.Insert(XlInsertShiftDirection.xlShiftToRight);

                sheet.Cells[baseRow - 1, insertAtCol].Value2 = $"Inj{newInjNum}";
                var targetZeroCell = sheet.Cells[baseRow - 2, insertAtCol];
                targetZeroCell.Value2 = 0;
                var styleSrcAbove = sheet.Cells[baseRow - 2, insertBeforeCol - 1];
                CopyCellFormatting(styleSrcAbove, targetZeroCell);

                var styleSrc = sheet.Cells[baseRow, insertBeforeCol - 1];

                for (int rowNum = 1; rowNum <= rowCount; rowNum++)
                {
                    string rowName = $"{tableRowPrefix}{rowNum}";
                    var rowRange = sheet.Names.Item(rowName).RefersToRange;

                    int newColCount = rowRange.Columns.Count + 1;
                    Range extended = sheet.Range[
                        sheet.Cells[rowRange.Row, startCol],
                        sheet.Cells[rowRange.Row, startCol + newColCount - 1]
                    ];

                    string refersTo = $"='{sheet.Name}'!{extended.get_AddressLocal(true, true)}";
                    sheet.Names.Item(rowName).Delete();
                    sheet.Names.Add(rowName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing, refersTo, Type.Missing, Type.Missing, Type.Missing);

                    int injOffset = newInjNum - 1;
                    int signalRow = signalStartRow + 1 + newInjNum;
                    int scaleRow = baseRow - 2;

                    string injColLetter = GetExcelColumnLetter(insertAtCol);
                    string signalColLetter = GetExcelColumnLetter(4 + ((rowNum - 1) * 3));

                    var newCell = sheet.Cells[rowRange.Row, insertAtCol];
                    newCell.Formula = $"=IF({signalColLetter}{signalRow}=\"\",\"\",FIXED({signalColLetter}{signalRow}, {injColLetter}{scaleRow}))";
                    CopyCellFormatting(styleSrc, newCell);
                }

                Range injNamedRange = sheet.Range[
                    sheet.Cells[baseRow - 1, insertAtCol],
                    sheet.Cells[baseRow - 1 + rowCount, insertAtCol]
                ];
                string injRef = $"='{sheet.Name}'!{injNamedRange.get_AddressLocal(true, true)}";
                string newInjName = $"{injPrefix}{newInjNum}";
                sheet.Names.Add(newInjName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, injRef, Type.Missing, Type.Missing, Type.Missing);
            }
            int minCol = FindColumnIndex(sheet, baseRow - 1, "Minimum");
            int maxCol = FindColumnIndex(sheet, baseRow - 1, "Maximum");
            var styeSrcFinal = sheet.Cells[baseRow - 2, minCol - 1];

            var minZeroCell = sheet.Cells[baseRow - 2, minCol];
            minZeroCell.Value2 = 0;
            CopyCellFormatting(styeSrcFinal, minZeroCell);

            var maxZeroCell = sheet.Cells[baseRow - 2, maxCol];
            maxZeroCell.Value2 = 0;
            CopyCellFormatting(styeSrcFinal, maxZeroCell);

            for (int rowNum = 1; rowNum <= rowCount; rowNum++)
            {
                string rowName = $"{tableRowPrefix}{rowNum}";
                var rowRange = sheet.Names.Item(rowName).RefersToRange;

                int rowIdx = rowRange.Row;

                string minColLetter = GetExcelColumnLetter(minCol);
                string maxColLetter = GetExcelColumnLetter(maxCol);
                string signalColLetter = GetExcelColumnLetter(4 + ((rowNum - 1) * 3));
                int signalRow = signalStartRow + 1;
                int snMinRow = signalRow + numInjectionsRL + 2;
                int snMaxRow = signalRow + numInjectionsRL + 3;

                var minCell = sheet.Cells[rowIdx, minCol];
                var maxCell = sheet.Cells[rowIdx, maxCol];

                minCell.Formula = $"=IF({signalColLetter}{snMinRow}=\"\",\"\",FIXED({signalColLetter}{snMinRow},{minColLetter}{baseRow - 2}))";
                maxCell.Formula = $"=IF({signalColLetter}{snMaxRow}=\"\",\"\",FIXED({signalColLetter}{snMaxRow},{maxColLetter}{baseRow - 2}))";

                CopyCellFormatting(sheet.Cells[rowIdx, minCol - 1], minCell);
                CopyCellFormatting(sheet.Cells[rowIdx, maxCol - 1], maxCell);
            }
        }

        private static void ExtendSummaryTable4Injections(_Worksheet sheet, int tbNumOfPeaks, int numInjectionsDl)
        {
            const string tableRowPrefix = "SummaryTable4Row";
            const string injPrefix = "SummaryTable4Inj";
            const string tableRow1Name = "SummaryTable4Row1";
            const int baseInjection = 6;
            int signalStartRow = sheet.Names.Item("SignalToNoiseResultsDL1").RefersToRange.Row;

            var row1 = sheet.Names.Item(tableRow1Name).RefersToRange;
            int baseRow = row1.Row;
            int startCol = row1.Column;

            for (int injNum = baseInjection; injNum > 1; injNum--)
            {
                if (injNum > numInjectionsDl)
                {
                    string injName = $"{injPrefix}{injNum}";
                    try
                    {
                        var injRange = sheet.Names.Item(injName).RefersToRange;
                        int injCol = injRange.Column;
                        int injRow = injRange.Row;
                        for (int row = injRow; row <= injRow + tbNumOfPeaks + 1; row++)
                        {
                            sheet.Cells[row, injCol].Delete(XlDeleteShiftDirection.xlShiftToLeft);
                        }
                        sheet.Names.Item(injName).Delete();
                    }
                    catch { }
                }
            }

            for (int injNum = 1; injNum <= Math.Min(numInjectionsDl, baseInjection); injNum++)
            {
                string injName = $"{injPrefix}{injNum}";
                try
                {
                    var injRange = sheet.Names.Item(injName).RefersToRange;
                    int injCol = injRange.Column;

                    for (int rowNum = 1; ; rowNum++)
                    {
                        string rowName = $"{tableRowPrefix}{rowNum}";
                        Range rowRange;
                        try { rowRange = sheet.Names.Item(rowName).RefersToRange; } catch { break; }

                        int absRow = rowRange.Row;
                        int signalRow = signalStartRow + 1 + injNum;
                        int scaleRow = baseRow - 2;
                        string signalColLetter = GetExcelColumnLetter(4 + (rowNum - 1) * 3);
                        string injColLetter = GetExcelColumnLetter(injCol);

                        var cell = sheet.Cells[absRow, injCol];
                        string formula = $"=IF({signalColLetter}{signalRow}=\"\",\"\",FIXED({signalColLetter}{signalRow}, {injColLetter}{scaleRow}))";
                        cell.Formula = formula;
                    }
                }
                catch { continue; }
            }

            row1 = sheet.Names.Item(tableRow1Name).RefersToRange;
            int totalCols = row1.Columns.Count;
            if (numInjectionsDl <= baseInjection) return;

            int colsToAdd = numInjectionsDl - baseInjection;
            int rowCount = 0;
            for (int i = 1; ; i++)
            {
                try { var _ = sheet.Names.Item($"{tableRowPrefix}{i}").RefersToRange; rowCount++; } catch { break; }
            }

            int insertBeforeCol = FindColumnIndex(sheet, baseRow - 1, "Minimum");

            for (int i = 1; i <= colsToAdd; i++)
            {
                int newInjNum = baseInjection + i;
                int insertAtCol = insertBeforeCol + i - 1;

                Range insertRange = sheet.Range[
                    sheet.Cells[baseRow - 1, insertAtCol],
                    sheet.Cells[baseRow - 1 + rowCount, insertAtCol]
                ];
                insertRange.Insert(XlInsertShiftDirection.xlShiftToRight);

                sheet.Cells[baseRow - 1, insertAtCol].Value2 = $"Inj{newInjNum}";
                var targetZeroCell = sheet.Cells[baseRow - 2, insertAtCol];
                targetZeroCell.Value2 = 0;
                var styleSrcAbove = sheet.Cells[baseRow - 2, insertBeforeCol - 1];
                CopyCellFormatting(styleSrcAbove, targetZeroCell);

                var styleSrc = sheet.Cells[baseRow, insertBeforeCol - 1];

                for (int rowNum = 1; rowNum <= rowCount; rowNum++)
                {
                    string rowName = $"{tableRowPrefix}{rowNum}";
                    var rowRange = sheet.Names.Item(rowName).RefersToRange;

                    int newColCount = rowRange.Columns.Count + 1;
                    Range extended = sheet.Range[
                        sheet.Cells[rowRange.Row, startCol],
                        sheet.Cells[rowRange.Row, startCol + newColCount - 1]
                    ];

                    string refersTo = $"='{sheet.Name}'!{extended.get_AddressLocal(true, true)}";
                    sheet.Names.Item(rowName).Delete();
                    sheet.Names.Add(rowName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing, refersTo, Type.Missing, Type.Missing, Type.Missing);

                    int injOffset = newInjNum - 1;
                    int signalRow = signalStartRow + 1 + newInjNum;
                    int scaleRow = baseRow - 2;
                    string injColLetter = GetExcelColumnLetter(insertAtCol);
                    string signalColLetter = GetExcelColumnLetter(4 + ((rowNum - 1) * 3));

                    var newCell = sheet.Cells[rowRange.Row, insertAtCol];
                    newCell.Formula = $"=IF({signalColLetter}{signalRow}=\"\",\"\",FIXED({signalColLetter}{signalRow}, {injColLetter}{scaleRow}))";
                    CopyCellFormatting(styleSrc, newCell);
                }

                Range injNamedRange = sheet.Range[
                    sheet.Cells[baseRow - 1, insertAtCol],
                    sheet.Cells[baseRow - 1 + rowCount, insertAtCol]
                ];
                string injRef = $"='{sheet.Name}'!{injNamedRange.get_AddressLocal(true, true)}";
                string newInjName = $"{injPrefix}{newInjNum}";
                sheet.Names.Add(newInjName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, injRef, Type.Missing, Type.Missing, Type.Missing);
            }
            int minCol = FindColumnIndex(sheet, baseRow - 1, "Minimum");
            int maxCol = FindColumnIndex(sheet, baseRow - 1, "Maximum");
            var styeSrcFinal = sheet.Cells[baseRow - 2, minCol - 1];

            var minZeroCell = sheet.Cells[baseRow - 2, minCol];
            minZeroCell.Value2 = 0;
            CopyCellFormatting(styeSrcFinal, minZeroCell);

            var maxZeroCell = sheet.Cells[baseRow - 2, maxCol];
            maxZeroCell.Value2 = 0;
            CopyCellFormatting(styeSrcFinal, maxZeroCell);

            for (int rowNum = 1; rowNum <= rowCount; rowNum++)
            {
                string rowName = $"{tableRowPrefix}{rowNum}";
                var rowRange = sheet.Names.Item(rowName).RefersToRange;

                int rowIdx = rowRange.Row;

                string minColLetter = GetExcelColumnLetter(minCol);
                string maxColLetter = GetExcelColumnLetter(maxCol);
                string signalColLetter = GetExcelColumnLetter(4 + ((rowNum - 1) * 3));
                int signalRow = signalStartRow + 1;
                int snMinRow = signalRow + numInjectionsDl + 2;
                int snMaxRow = signalRow + numInjectionsDl + 3;

                var minCell = sheet.Cells[rowIdx, minCol];
                var maxCell = sheet.Cells[rowIdx, maxCol];

                minCell.Formula = $"=IF({signalColLetter}{snMinRow}=\"\",\"\",FIXED({signalColLetter}{snMinRow},{minColLetter}{baseRow - 2}))";
                maxCell.Formula = $"=IF({signalColLetter}{snMaxRow}=\"\",\"\",FIXED({signalColLetter}{snMaxRow},{maxColLetter}{baseRow - 2}))";

                CopyCellFormatting(sheet.Cells[rowIdx, minCol - 1], minCell);
                CopyCellFormatting(sheet.Cells[rowIdx, maxCol - 1], maxCell);
            }
        }

        private static void UpdateSignalToNoiseRlFormulas(_Worksheet sheet, int tbNumPeaks, int numInjectionsRL, string cmbRlSn)
        {
            const int columnsPerBlock = 3;
            //int baseCol = 3;
            int resultsColOffset = 2;
            const int headerRows = 2;
            const int summaryRows = 3;

            if (cmbRlSn == "≥") cmbRlSn = ">=";
            if (cmbRlSn == "≤") cmbRlSn = "<=";

            for (int peakIndex = 1; peakIndex <= tbNumPeaks; peakIndex++)
            {
                string rangeName = $"SignalToNoiseResults{peakIndex}";
                try
                {
                    var range = sheet.Names.Item(rangeName).RefersToRange;
                    int startRow = range.Row;
                    int startCol = range.Column;
                    int totalRows = range.Rows.Count;

                    int resultCol = startCol + resultsColOffset;
                    int snDataCol = startCol + 1;
                    int resultStartRow = startRow + headerRows;

                    int acRlSnRow = sheet.Names.Item("AC_RL_SN").RefersToRange.Row;
                    string thresholdRef = $"$D${acRlSnRow}";

                    int existingInjectionCount = 0;
                    for (int i = 0; i < totalRows - headerRows - summaryRows; i++)
                    {
                        var cell = sheet.Cells[resultStartRow + i, resultCol];
                        if (cell.Formula != null && cell.Formula.ToString().Trim() != "")
                        {
                            cell.Formula = cell.Formula.ToString().Replace(">=", cmbRlSn);
                            existingInjectionCount++;
                        }

                    }

                    int rowsToAdd = numInjectionsRL - existingInjectionCount;
                    if (rowsToAdd <= 0) continue;

                    int signalStartRow = resultStartRow + existingInjectionCount;

                    for (int i = 0; i < rowsToAdd; i++)
                    {
                        int injRow = signalStartRow + i;
                        string snDataCell = GetExcelColumnLetter(snDataCol) + injRow;
                        string formula = $"=IF(AND({thresholdRef} <> \"\",{snDataCell} <> \"\"),IF(ROUND({snDataCell},0){cmbRlSn}{thresholdRef},\"Pass\",\"Fail\"),\"\")";

                        var targetCell = sheet.Cells[injRow, resultCol];
                        targetCell.Formula = formula;

                        var srcCell = sheet.Cells[signalStartRow - 1, resultCol];
                        CopyCellFormatting(srcCell, targetCell);
                    }


                    int updatedRowCount = headerRows + numInjectionsRL + summaryRows;
                    Range newRange = sheet.Range[
                        sheet.Cells[startRow, startCol],
                        sheet.Cells[startRow + updatedRowCount - 1, startCol + columnsPerBlock - 1]
                    ];
                    string refersTo = $"='{sheet.Name}'!{newRange.get_AddressLocal(true, true)}";
                    sheet.Names.Item(rangeName).Delete();
                    sheet.Names.Add(rangeName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing, refersTo, Type.Missing, Type.Missing, Type.Missing);

                    int meanRow = startRow + headerRows + numInjectionsRL;
                    int minRow = meanRow + 1;
                    int maxRow = minRow + 1;
                    int firstInjRow = startRow + headerRows;
                    int lastInjRow = firstInjRow + numInjectionsRL - 1;
                    string rangeLetter = GetExcelColumnLetter(snDataCol);

                    sheet.Cells[meanRow, snDataCol].Formula = $"=IF(SUM({rangeLetter}{firstInjRow}:{rangeLetter}{lastInjRow})<>0,ROUND(AVERAGE({rangeLetter}{firstInjRow}:{rangeLetter}{lastInjRow}),0),\"\")";
                    sheet.Cells[minRow, snDataCol].Formula = $"=IF(SUM({rangeLetter}{firstInjRow}:{rangeLetter}{lastInjRow})<>0,ROUND(MIN({rangeLetter}{firstInjRow}:{rangeLetter}{lastInjRow}),0),\"\")";
                    sheet.Cells[maxRow, snDataCol].Formula = $"=IF(SUM({rangeLetter}{firstInjRow}:{rangeLetter}{lastInjRow})<>0,ROUND(MAX({rangeLetter}{firstInjRow}:{rangeLetter}{lastInjRow}),0),\"\")";

                }
                catch
                {

                }
            }

        }

        private static void UpdateImpurityResultsSummaryFormulas(_Worksheet sheet, int tbNumPeaks, string cmbRlRsd)
        {
            var sampleRange = sheet.Names.Item("SampleNumsRawData").RefersToRange;
            int startRow = sampleRange.Row;
            int injectionCount = sampleRange.Rows.Count;
            int summaryStartRow = startRow + injectionCount;

            for (int i = 1; i <= tbNumPeaks; i++)
            {
                string baseRangeName = $"ImpurityResults{i}";
                try
                {
                    var baseRange = sheet.Names.Item(baseRangeName).RefersToRange;
                    int baseCol = baseRange.Column;
                    string idColLetter = GetExcelColumnLetter(baseCol);
                    string respColLetter = GetExcelColumnLetter(baseCol + 1);

                    int decimalPrecisionRow = summaryStartRow;
                    string decimalPrecisionCell = $"{idColLetter}{decimalPrecisionRow}";

                    // N
                    sheet.Cells[summaryStartRow, baseCol].Formula = "=IF(ISERROR(IF(LEN($D$3)-FIND(\".\",$D$3)>0,(LEN($D$3)-FIND(\".\",$D$3)))),0,IF(LEN($D$3)-FIND(\".\",$D$3)>0,(LEN($D$3)-FIND(\".\",$D$3))))";

                    sheet.Cells[summaryStartRow, baseCol + 1].Formula =
                        $"=IF(SUM({respColLetter}{startRow}:{respColLetter}{summaryStartRow - 1})<>0,COUNT({respColLetter}{startRow}:{respColLetter}{summaryStartRow - 1}),\"\")";

                    // Mean
                    sheet.Cells[summaryStartRow + 1, baseCol + 1].Formula =
                        $"=IF(SUM({respColLetter}{startRow}:{respColLetter}{summaryStartRow - 1})<>0,ROUND(AVERAGE({respColLetter}{startRow}:{respColLetter}{summaryStartRow - 1}), 0),\"\")";

                    // StdDev
                    sheet.Cells[summaryStartRow + 2, baseCol + 1].Formula =
                        $"=IF(SUM({respColLetter}{startRow}:{respColLetter}{summaryStartRow - 1})<>0,ROUND(STDEV({respColLetter}{startRow}:{respColLetter}{summaryStartRow - 1}), 2),\"\")";

                    // %RSD
                    sheet.Cells[summaryStartRow + 3, baseCol].Formula =
                        $"=IF(AND($D$3<>\"\",{respColLetter}{summaryStartRow + 1}<>\"\"),IF(ROUND((STDEV({respColLetter}{startRow}:{respColLetter}{summaryStartRow - 1})/AVERAGE({respColLetter}{startRow}:{respColLetter}{summaryStartRow - 1}))*100,{decimalPrecisionCell}){cmbRlRsd}VALUE($D$3),\"Pass\",\"Fail\"),\"\")";

                    sheet.Cells[summaryStartRow + 3, baseCol + 1].Formula =
                        $"=IF(SUM({respColLetter}{startRow}:{respColLetter}{summaryStartRow - 1})<>0,FIXED((STDEV({respColLetter}{startRow}:{respColLetter}{summaryStartRow - 1})/AVERAGE({respColLetter}{startRow}:{respColLetter}{summaryStartRow - 1}))*100,{decimalPrecisionCell}),\"\")";
                }
                catch { continue; }
            }

        }

        private static void UpdateSignalToNoiseDlFormulas(_Worksheet sheet, int tbNumPeaks, int numInjectionsDL, string cmbDlSn)
        {
            const int columnsPerBlock = 3;
            int resultsColOffset = 2;
            const int headerRows = 2;
            const int summaryRows = 3;
            if (cmbDlSn == "≥") cmbDlSn = ">=";
            if (cmbDlSn == "≤") cmbDlSn = "<=";

            for (int peakIndex = 1; peakIndex <= tbNumPeaks; peakIndex++)
            {
                string rangeName = $"SignalToNoiseResultsDL{peakIndex}";
                try
                {
                    var range = sheet.Names.Item(rangeName).RefersToRange;
                    int startRow = range.Row;
                    int startCol = range.Column;
                    int totalRows = range.Rows.Count;

                    int resultCol = startCol + resultsColOffset;
                    int snDataCol = startCol + 1;
                    int resultStartRow = startRow + headerRows;

                    int existingInjectionCount = 0;
                    for (int i = 0; i < totalRows - headerRows - summaryRows; i++)
                    {
                        var cell = sheet.Cells[resultStartRow + i, resultCol];
                        if (cell.Formula != null && cell.Formula.ToString().Trim() != "")
                        {
                            cell.Formula = cell.Formula.ToString().Replace(">=", cmbDlSn);
                            existingInjectionCount++;
                        }
                    }

                    int rowsToAdd = numInjectionsDL - existingInjectionCount;
                    if (rowsToAdd <= 0) continue;

                    int signalStartRow = resultStartRow + existingInjectionCount;
                    int acDlSnRow = sheet.Names.Item("AC_DL_SN").RefersToRange.Row;

                    string thresholdRef = $"$D${acDlSnRow}";

                    for (int i = 0; i < rowsToAdd; i++)
                    {
                        int injRow = signalStartRow + i;
                        string snDataCell = GetExcelColumnLetter(snDataCol) + injRow;
                        string formula = $"=IF(AND({thresholdRef} <> \"\",{snDataCell} <> \"\"),IF(ROUND({snDataCell},0){cmbDlSn}{thresholdRef},\"Pass\",\"Fail\"),\"\")";

                        var targetCell = sheet.Cells[injRow, resultCol];
                        targetCell.Formula = formula;

                        var srcCell = sheet.Cells[signalStartRow - 1, resultCol];
                        CopyCellFormatting(srcCell, targetCell);
                    }

                    int updatedRowCount = headerRows + numInjectionsDL + summaryRows;
                    Range newRange = sheet.Range[
                        sheet.Cells[startRow, startCol],
                        sheet.Cells[startRow + updatedRowCount - 1, startCol + columnsPerBlock - 1]
                    ];
                    string refersTo = $"='{sheet.Name}'!{newRange.get_AddressLocal(true, true)}";
                    sheet.Names.Item(rangeName).Delete();
                    sheet.Names.Add(rangeName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing, refersTo, Type.Missing, Type.Missing, Type.Missing);

                    int meanRow = startRow + headerRows + numInjectionsDL;
                    int minRow = meanRow + 1;
                    int maxRow = minRow + 1;
                    int firstInjRow = startRow + headerRows;
                    int lastInjRow = firstInjRow + numInjectionsDL - 1;
                    string rangeLetter = GetExcelColumnLetter(snDataCol);

                    sheet.Cells[meanRow, snDataCol].Formula = $"=IF(SUM({rangeLetter}{firstInjRow}:{rangeLetter}{lastInjRow})<>0,ROUND(AVERAGE({rangeLetter}{firstInjRow}:{rangeLetter}{lastInjRow}),0),\"\")";
                    sheet.Cells[minRow, snDataCol].Formula = $"=IF(SUM({rangeLetter}{firstInjRow}:{rangeLetter}{lastInjRow})<>0,ROUND(MIN({rangeLetter}{firstInjRow}:{rangeLetter}{lastInjRow}),0),\"\")";
                    sheet.Cells[maxRow, snDataCol].Formula = $"=IF(SUM({rangeLetter}{firstInjRow}:{rangeLetter}{lastInjRow})<>0,ROUND(MAX({rangeLetter}{firstInjRow}:{rangeLetter}{lastInjRow}),0),\"\")";

                }
                catch
                {

                }
            }

        }

        private static int FindColumnIndex(_Worksheet sheet, int headerRow, string labelText)
        {
            for (int col = 1; col <= sheet.UsedRange.Columns.Count; col++)
            {
                var val = sheet.Cells[headerRow, col].Value2?.ToString().Trim();
                if (val == labelText) return col;
            }
            return sheet.UsedRange.Columns.Count;
        }

        private static void AdjustInjectionRowsByNamedRange(_Worksheet sheet, string rangeName, int numInjections)
        {
            var namedRange = sheet.Names.Item(rangeName, Type.Missing, Type.Missing);
            var sampleRange = namedRange.RefersToRange;

            int baseRow = sampleRange.Row;
            int baseCount = sampleRange.Rows.Count;
            int baseCol = sampleRange.Column;

            int desiredCount = numInjections;

            if (desiredCount > baseCount)
            {
                for (int i = baseCount + 1; i <= desiredCount; i++)
                {
                    int insertAt = baseRow + i - 1;
                    sheet.Rows[insertAt].Insert();
                    sheet.Cells[insertAt, baseCol].Value2 = i;
                }
            }
            else if (desiredCount < baseCount)
            {
                for (int i = baseCount; i > desiredCount; i--)
                {
                    int deleteAt = baseRow + i - 1;
                    sheet.Rows[deleteAt].Delete();
                }
            }

            var newBottomRow = baseRow + desiredCount - 1;
            Range updatedRange = sheet.Range[sheet.Cells[baseRow, baseCol], sheet.Cells[newBottomRow, baseCol]];
            string refersTo = $"='{sheet.Name}'!{updatedRange.get_AddressLocal(true, true)}";

            sheet.Names.Add(rangeName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, refersTo, Type.Missing, Type.Missing, Type.Missing);

            //CopyCellFormatting(sampleRange, updatedRange);
        }

        private static string UpdateFormulaReferenceColumn(string formula, int baseColIndex, int offset)
        {
            string originalCol = GetExcelColumnLetter(baseColIndex);
            string newCol = GetExcelColumnLetter(baseColIndex + offset);

            var regEx = new System.Text.RegularExpressions.Regex($@"\b{originalCol}(\d+)\b");
            return regEx.Replace(formula, match => $"{newCol}{match.Groups[1].Value}");
        }

        private static void ApplyImpurityResultsStyling(_Worksheet sheet, int numInjectionsRL, int tbNumOfPeaks)
        {
            try
            {
                var sampleRange = sheet.Names.Item("SampleNumsRawData").RefersToRange;
                int startRow = sampleRange.Row;
                int lastRow = startRow + numInjectionsRL - 1;
                int startCol = sampleRange.Column;

                var lastImpurityRange = sheet.Names.Item($"ImpurityResults{tbNumOfPeaks}").RefersToRange;
                int lastCol = lastImpurityRange.Column + lastImpurityRange.Columns.Count - 1;


                for (int row = startRow; row < lastRow; row++)
                {
                    for (int col = startCol; col <= lastCol; col++)
                    {
                        sheet.Cells[row, col].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlLineStyleNone;
                    }
                }

                Range bottomRange = sheet.Range[
                    sheet.Cells[lastRow, sampleRange.Column],
                    sheet.Cells[lastRow, lastCol]
                ];

                bottomRange.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                bottomRange.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;


                for (int row = startRow; row <= lastRow + 4; row++)
                {
                    var cell = sheet.Cells[row, lastCol];
                    cell.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    cell.Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThin;
                }
            }
            catch (Exception ex)
            {
                Logger.LogMessage($"[Styling] Failed in ApplyImpurityResultsStyling: {ex.Message}", Level.Error);
            }
        }

        private static void ApplySignalToNoiseRlStyling(_Worksheet sheet, int numInjectionsRL, int tbNumOfPeaks)
        {
            try
            {
                var sampleRange = sheet.Names.Item("SampleNumsRawData2").RefersToRange;
                int startRow = sampleRange.Row;
                int lastRow = startRow + numInjectionsRL - 1;
                int startCol = sampleRange.Column;

                var lastImpurityRange = sheet.Names.Item($"SignalToNoiseResults{tbNumOfPeaks}").RefersToRange;
                int lastCol = lastImpurityRange.Column + lastImpurityRange.Columns.Count - 1;


                for (int row = startRow; row < lastRow; row++)
                {
                    for (int col = startCol; col <= lastCol; col++)
                    {
                        sheet.Cells[row, col].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlLineStyleNone;
                    }
                }

                Range bottomRange = sheet.Range[
                    sheet.Cells[lastRow, sampleRange.Column],
                    sheet.Cells[lastRow, lastCol]
                ];

                bottomRange.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                bottomRange.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;


                for (int row = startRow; row <= lastRow + 4; row++)
                {
                    var cell = sheet.Cells[row, lastCol];
                    cell.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    cell.Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThin;
                }
            }
            catch (Exception ex)
            {
                Logger.LogMessage($"Styling Failed in ApplySignalToNoiseRlStyling: {ex.Message}", Level.Error);
            }
        }

        private static void ApplySignalToNoiseDlStyling(_Worksheet sheet, int numInjectionsDL, int tbNumOfPeaks)
        {
            try
            {
                var sampleRange = sheet.Names.Item("SampleRawDataDL").RefersToRange;
                int startRow = sampleRange.Row;
                int lastRow = startRow + numInjectionsDL - 1;
                int startCol = sampleRange.Column;

                var lastImpurityRange = sheet.Names.Item($"SignalToNoiseResultsDL{tbNumOfPeaks}").RefersToRange;
                int lastCol = lastImpurityRange.Column + lastImpurityRange.Columns.Count - 1;


                for (int row = startRow; row < lastRow; row++)
                {
                    for (int col = startCol; col <= lastCol; col++)
                    {
                        sheet.Cells[row, col].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlLineStyleNone;
                    }
                }

                Range bottomRange = sheet.Range[
                    sheet.Cells[lastRow, sampleRange.Column],
                    sheet.Cells[lastRow, lastCol]
                ];

                bottomRange.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                bottomRange.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;


                for (int row = startRow; row <= lastRow + 4; row++)
                {
                    var cell = sheet.Cells[row, lastCol];
                    cell.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    cell.Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThin;
                }
            }
            catch (Exception ex)
            {
                Logger.LogMessage($"Styling Failed in ApplySignalToNoiseDlStyling: {ex.Message}", Level.Error);
            }
        }

        private static void ApplySummaryTableStyling(_Worksheet sheet, string summaryTablename)
        {
            try
            {
                string baseRowName = $"{summaryTablename}Row";
                int rowIndex = 1;

                while (true)
                {
                    try
                    {
                        var _ = sheet.Names.Item($"{baseRowName}{rowIndex}").RefersToRange;
                        rowIndex++;
                    }
                    catch
                    {
                        break;
                    }
                }

                if (rowIndex == 1) return;
                int lastRowIndex = rowIndex - 1;

                var summryTableRange = sheet.Names.Item($"{summaryTablename}").RefersToRange;
                int tableStartCol = summryTableRange.Column;

                int injCount = 0;
                for (int i = 1; ; i++)
                {
                    try
                    {
                        var _ = sheet.Names.Item($"{summaryTablename}Inj{i}").RefersToRange;
                        injCount++;
                    }
                    catch
                    {
                        break;
                    }
                }

                int tableEndCol = tableStartCol + 1 + injCount;

                for (int i = 1; i <= lastRowIndex; i++)
                {
                    try
                    {
                        var rowRange = sheet.Names.Item($"{baseRowName}{i}").RefersToRange;
                        var rowIdx = rowRange.Row;
                        var fullRowRange = sheet.Range[sheet.Cells[rowIdx, tableStartCol], sheet.Cells[rowIdx, tableEndCol]];
                        fullRowRange.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlLineStyleNone;
                        fullRowRange.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlLineStyleNone;
                    }
                    catch { }
                }

                var firstRowRange = sheet.Names.Item($"{baseRowName}1").RefersToRange;
                var firstRowIdx = firstRowRange.Row;
                var fullRowFullRange = sheet.Range[sheet.Cells[firstRowIdx, tableStartCol + 1], sheet.Cells[firstRowIdx, tableEndCol]];

                fullRowFullRange.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                fullRowFullRange.Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlMedium;

                var lastRowRange = sheet.Names.Item($"{baseRowName}{lastRowIndex}").RefersToRange;
                var lastRowIdx = lastRowRange.Row;
                var lastRowFullRange = sheet.Range[sheet.Cells[lastRowIdx, tableStartCol + 1], sheet.Cells[lastRowIdx, tableEndCol]];
                lastRowFullRange.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;

                int headerRow = firstRowRange.Row - 1;

                var injHeaderRow = sheet.Range[sheet.Cells[headerRow, tableStartCol + 1], sheet.Cells[headerRow, tableEndCol]];
                injHeaderRow.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                injHeaderRow.Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlThin;
            }
            catch (Exception ex)
            {
                Logger.LogMessage($"Styling Error in ApplySummaryTable1Styling: {ex.Message}", Level.Error);
            }
        }

        private static void ApplySummaryTableWithMinMaxStyling(_Worksheet sheet, string summaryTablename)
        {
            try
            {
                string baseRowName = $"{summaryTablename}Row";
                int rowIndex = 1;

                while (true)
                {
                    try
                    {
                        var _ = sheet.Names.Item($"{baseRowName}{rowIndex}").RefersToRange;
                        rowIndex++;
                    }
                    catch
                    {
                        break;
                    }
                }

                if (rowIndex == 1) return;
                int lastRowIndex = rowIndex - 1;

                var summryTableRange = sheet.Names.Item($"{summaryTablename}Row1").RefersToRange;
                int tableStartCol = summryTableRange.Column;

                int injCount = 0;
                for (int i = 1; ; i++)
                {
                    try
                    {
                        var _ = sheet.Names.Item($"{summaryTablename}Inj{i}").RefersToRange;
                        injCount++;
                    }
                    catch
                    {
                        break;
                    }
                }

                int tableEndCol = tableStartCol + injCount + 3;

                for (int i = 1; i <= lastRowIndex; i++)
                {
                    try
                    {
                        var rowRange = sheet.Names.Item($"{baseRowName}{i}").RefersToRange;
                        rowRange.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlLineStyleNone;
                    }
                    catch { }
                }

                var lastRowRange = sheet.Names.Item($"{baseRowName}{lastRowIndex}").RefersToRange;
                var lastRowIdx = lastRowRange.Row;
                var lastRowFullRange = sheet.Range[sheet.Cells[lastRowIdx, tableStartCol + 1], sheet.Cells[lastRowIdx, tableEndCol]];
                lastRowFullRange.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
            }
            catch (Exception ex)
            {
                Logger.LogMessage($"Styling Error in ApplySummaryTable1Styling: {ex.Message}", Level.Error);
            }
        }
    } // end class

}
