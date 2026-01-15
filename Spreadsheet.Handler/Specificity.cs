using log4net.Core;
using Microsoft.Office.Interop.Excel;
using Internal.Framework.Collections;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;

namespace Spreadsheet.Handler
{
    public static class Specificity
    {
        private static Application _app;

        private const string TempDirectoryName = "ABD_TempFiles";

        private const int DefaulttbNumOfPeaks = 2;
        private const int DefaulttbNumOfSamples = 1;


        /// <summary>
        /// Method to be called via Scripts - GenerateExperiment
        /// </summary>
        /// <param name="sourcePath"></param>
        /// <returns></returns>
        public static string UpdateSpecificitySheet(string sourcePath, int numPeaks, int numSamples, int numSolPeakPurity, int numSolDissol, string cmbProtocolType, string cmbProductType, string cmbTestType)
        {
            string returnPath = "";
            try
            {
                returnPath = UpdateSpecificitySheet2(sourcePath, numPeaks, numSamples, numSolPeakPurity, numSolDissol, cmbProtocolType, cmbProductType, cmbTestType);
            }
            catch (Exception ex)
            {
                Logger.LogMessage("An error occurred in the call to Specificity.Specificity. Message and stack trace are:\r\n" + ex.Message + "\r\n" + ex.StackTrace, Level.Error);

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
                            Logger.LogMessage("An error occurred in the call to Specificity.UpdateSpecificitySheet. Failed to save current workbook changes and to get path.", Level.Error);
                        }

                        _app.Workbooks.Close();
                    }
                    _app = null;
                }
                catch
                {
                    Logger.LogMessage("An error occurred in the call to Specificity.UpdateSpecificitySheet. Application failed to close workbooks. Message and stack trace are:\r\n" + ex.Message + "\r\n" + ex.StackTrace, Level.Error);
                }
                finally
                {
                    WorksheetUtilities.ReleaseExcelApp();
                }
            }
            return returnPath;
        }


        //-------------------------
        //-----PRIVATE METHODS-----
        //-------------------------

        /// <summary>
        /// Method with the logic for calling / updating Excel spreadsheet for Specificity.
        /// </summary>
        /// <param name="sourcePath"></param>
        /// <param name=""></param>
        /// <returns></returns>
        private static string UpdateSpecificitySheet2(string sourcePath, int numPeaks, int numSamples, int numSolPeakPurity, int numSolDissol, string cmbProtocolType, string cmbProductType, string cmbTestType)
        {
            if (!File.Exists(sourcePath))
            {
                Logger.LogMessage("Error in call to Specificity.UpdateSpecificitySheet. Invalid source file path specified.", Level.Error);
                return "";
            }

            // Generate an random temp path to save new workbook
            string savePath = WorksheetUtilities.CopyWorkbook(sourcePath, TempDirectoryName, "Specificity Results.xls");
            if (String.IsNullOrEmpty(savePath)) return "";

            // Try to open the file
            _app = WorksheetUtilities.GetExcelApp();
            _app.Workbooks.Open(savePath, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            Workbook book = _app.Workbooks[1];
            // Worksheet sheet = book.Worksheets[1] as Worksheet;
            book.Application.DisplayAlerts = false;

            bool isProdTypeDrugSubstance = false;
            bool isProdTypeSdd = false;
            bool isDrugProduct = false;
            isProdTypeDrugSubstance = cmbProductType == "Drug Substance";
            isProdTypeSdd = cmbProductType == "SDD";
            isDrugProduct = cmbProductType == "Drug Product";
            var keepSheetName = "";

            if ((isProdTypeDrugSubstance || isProdTypeSdd || isDrugProduct) && !((cmbTestType == "Water content" || cmbTestType == "Dissolution" || cmbTestType == "Volatiles")))
            {
                keepSheetName = "Specificity-updated";

                Worksheet sheet = book.Sheets[keepSheetName] as Worksheet;
                //Worksheet sheet = book.Worksheets[1] as Worksheet;
                EnrichSpecificityUpadted(sheet, numPeaks, numSamples, numSolPeakPurity, isProdTypeDrugSubstance, isProdTypeSdd, isDrugProduct, cmbProtocolType, cmbProductType, cmbTestType);
            }
            else if (cmbTestType == "Dissolution")
            {
                keepSheetName = "Specificity-Dissolution";
                Worksheet sheet = book.Sheets[keepSheetName] as Worksheet;
                EnrichSpecificityDissolution(sheet, numPeaks, numSamples, numSolPeakPurity, cmbProtocolType, cmbProductType, cmbTestType);

            }
            else if (cmbTestType == "Water content")
            {
                keepSheetName = "Specificity-Water Content";
                Worksheet sheet = book.Sheets[keepSheetName] as Worksheet;
                EnrichSpecificityWaterContent(sheet, cmbProtocolType, cmbProductType, cmbTestType);
            }
            else if (cmbTestType == "Volatiles")
            {
                keepSheetName = "Specificity-Volatiles";

                Worksheet sheet = book.Sheets[keepSheetName] as Worksheet;

                EnrichSpecificityVolatiles(sheet, numPeaks, numSamples, numSolPeakPurity, isProdTypeDrugSubstance, isProdTypeSdd, isDrugProduct, cmbProtocolType, cmbProductType, cmbTestType);
            }

            if (!string.IsNullOrEmpty(keepSheetName))
            {
                foreach (Worksheet ws in book.Sheets)
                {
                    if ((ws.Name != keepSheetName))
                    {
                        ws.Delete();
                    }
                }
            }

            book.Application.DisplayAlerts = true;

            _app.Workbooks[1].Save();

            WorksheetUtilities.ReleaseComObject(book);
            _app.Workbooks.Close();

            //WorksheetUtilities.ReleaseComObject(_app);
            _app = null;
            WorksheetUtilities.ReleaseExcelApp();

            // Return the path
            return savePath;
        }

        public static void EnrichSpecificityUpadted(Worksheet sheet, int numPeaks, int numSamples, int numSolPeakPurity, bool isProdTypeDrugSubstance, bool isProdTypeSdd, bool isDrugProduct, string cmbProtocolType, string cmbProductType, string cmbTestType)
        {

            if (sheet != null)
            {
                bool wasProtected = WorksheetUtilities.SetSheetProtection(sheet, null, false);

                if (!string.IsNullOrEmpty(cmbProtocolType))
                {
                    sheet.Cells[6, 4] = cmbProtocolType;
                }
                if (!string.IsNullOrEmpty(cmbProductType))
                {
                    sheet.Cells[7, 4] = cmbProductType;
                }

                if (!string.IsNullOrEmpty(cmbTestType))
                {
                    sheet.Cells[8, 4] = cmbTestType;
                }


                if (numPeaks > DefaulttbNumOfPeaks)
                {
                    int rowsToInsert = numPeaks - DefaulttbNumOfPeaks;
                    int indexToInsertFrom = WorksheetUtilities.GetNamedRangeRowCount(sheet, "Default_Peak_Of_Interest_Table");

                    WorksheetUtilities.InsertRowsIntoNamedRangeFromRow(rowsToInsert, sheet, "Default_Peak_Of_Interest_Table", false, XlDirection.xlDown, XlPasteType.xlPasteAll, indexToInsertFrom);
                }

                if (isProdTypeDrugSubstance)
                {
                    if (numPeaks > DefaulttbNumOfPeaks)
                    {
                        int rowsToInsert = numPeaks - DefaulttbNumOfPeaks;
                        int indexToInsertFrom = WorksheetUtilities.GetNamedRangeRowCount(sheet, "Drug_Substance_Summary1");

                        WorksheetUtilities.InsertRowsIntoNamedRangeFromRow(rowsToInsert, sheet, "Drug_Substance_Summary1", false, XlDirection.xlDown, XlPasteType.xlPasteAll, indexToInsertFrom);
                    }

                    if (numSamples > DefaulttbNumOfSamples)
                    {
                        int rowsToInsert = numSamples - DefaulttbNumOfSamples;
                        int indexToInsertFrom = WorksheetUtilities.GetNamedRangeRowCount(sheet, "Drug_Substance_Sample");
                        AppendRowsCopyOnlyFormulasAndRenumber(sheet, "Drug_Substance_Sample", rowsToInsert, 1, 1, 3);
                        EnsureDrugSubstanceSummaries(sheet, "Drug_Substance_Summary1", "Drug_Substance_Sample", numSamples);
                    }

                    if (numSamples > DefaulttbNumOfSamples)
                    {
                        //EnsureDrugSubstanceSummaries(sheet, "Drug_Substance_Summary1", numSamples);
                    }

                    if (numSolPeakPurity <= 0)
                    {
                        WorksheetUtilities.DeleteNamedRangeRows(sheet, "Drug_Substance_Peak");
                    }
                    else
                    {
                        if (numSolPeakPurity > 1)
                        {
                            int rowsToInsert = numSolPeakPurity - 1;

                            AppendRowsCopyOnlyFormulasAndRenumber(sheet, "Drug_Substance_Peak_Sample", rowsToInsert, 1, 1, 3);
                            EnsurePeakSummaryRowsWithFixedArgs(sheet, numSolPeakPurity, "Drug_Substance_Peak_Summary1");
                        }
                    }
                }
                else
                {
                    WorksheetUtilities.DeleteNamedRangeRows(sheet, "Drug_Substance_Section");
                }

                if (isProdTypeSdd)
                {
                    if (numPeaks > DefaulttbNumOfPeaks)
                    {
                        int rowsToInsert = numPeaks - DefaulttbNumOfPeaks;
                        int indexToInsertFrom = WorksheetUtilities.GetNamedRangeRowCount(sheet, "SDD_Summary1");

                        WorksheetUtilities.InsertRowsIntoNamedRangeFromRow(rowsToInsert, sheet, "SDD_Summary1", false, XlDirection.xlDown, XlPasteType.xlPasteAll, indexToInsertFrom);
                    }

                    if (numSamples > DefaulttbNumOfSamples)
                    {
                        int rowsToInsert = numSamples - DefaulttbNumOfSamples;
                        int indexToInsertFrom = WorksheetUtilities.GetNamedRangeRowCount(sheet, "SDD_Sample");
                        AppendRowsCopyOnlyFormulasAndRenumber(sheet, "SDD_Sample", rowsToInsert, 1, 1, 3, isProdTypeSdd);
                        EnsureDrugSubstanceSummaries(sheet, "SDD_Summary1", "SDD_Sample", numSamples);

                    }

                    if (numSamples > DefaulttbNumOfSamples)
                    {
                        //EnsureDrugSubstanceSummaries(sheet, "Drug_Substance_Summary1", numSamples);
                    }

                    if (numSolPeakPurity <= 0)
                    {
                        WorksheetUtilities.DeleteNamedRangeRows(sheet, "SDD_Peak_Purity_Section");
                    }
                    else
                    {
                        if (numSolPeakPurity > 1)
                        {
                            int rowsToInsert = numSolPeakPurity - 1;

                            AppendRowsCopyOnlyFormulasAndRenumber(sheet, "SDD_Peak_Sample", rowsToInsert, 1, 1, 3, isProdTypeSdd);
                            EnsurePeakSummaryRowsWithFixedArgs(sheet, numSolPeakPurity, "SDD_Peak_Summary1");
                        }
                    }
                }
                else
                {
                    WorksheetUtilities.DeleteNamedRangeRows(sheet, "SDD_Section");
                }

                if (isDrugProduct && !(cmbTestType == "Water content" || cmbTestType == "Dissolution" || cmbTestType == "Volatiles"))
                {
                    if (numPeaks > DefaulttbNumOfPeaks)
                    {
                        int rowsToInsert = numPeaks - DefaulttbNumOfPeaks;
                        int indexToInsertFrom = WorksheetUtilities.GetNamedRangeRowCount(sheet, "Drug_Product_Summary1");

                        WorksheetUtilities.InsertRowsIntoNamedRangeFromRow(rowsToInsert, sheet, "Drug_Product_Summary1", false, XlDirection.xlDown, XlPasteType.xlPasteAll, indexToInsertFrom);
                    }

                    if (numSamples > DefaulttbNumOfSamples)
                    {
                        int rowsToInsert = numSamples - DefaulttbNumOfSamples;
                        int indexToInsertFrom = WorksheetUtilities.GetNamedRangeRowCount(sheet, "Drug_Product_Sample");
                        AppendRowsCopyOnlyFormulasAndRenumber(sheet, "Drug_Product_Sample", rowsToInsert, 1, 1, 3);
                        EnsureDrugSubstanceSummaries(sheet, "Drug_Product_Summary1", "Drug_Product_Sample", numSamples);

                    }

                    if (numSamples > DefaulttbNumOfSamples)
                    {
                        //EnsureDrugSubstanceSummaries(sheet, "Drug_Substance_Summary1", numSamples);
                    }

                    if (numSolPeakPurity <= 0)
                    {
                        WorksheetUtilities.DeleteNamedRangeRows(sheet, "Drug_Product_Peak_Purity");
                    }
                    else
                    {
                        if (numSolPeakPurity > 1)
                        {
                            int rowsToInsert = numSolPeakPurity - 1;

                            AppendRowsCopyOnlyFormulasAndRenumber(sheet, "Drug_Product_Peak_Sample", rowsToInsert, 1, 1, 3);
                            EnsurePeakSummaryRowsWithFixedArgs(sheet, numSolPeakPurity, "Drug_Product_Peak_Summary1");
                        }
                    }
                }
                else
                {
                    WorksheetUtilities.DeleteNamedRangeRows(sheet, "Drug_Product_Section");
                }


                //if (!isDrugProduct || isDrugProduct && (cmbTestType == "Water content" || cmbTestType == "Dissolution" || cmbTestType == "Volatiles"))

                try
                {
                    _app.Goto(sheet.Cells[1, 1], true);
                }
                catch
                {
                    Logger.LogMessage("Scroll of sheet failed in Specificity.Specificity!", Level.Error);
                }

                if (wasProtected) WorksheetUtilities.SetSheetProtection(sheet, null, false);

                WorksheetUtilities.ReleaseComObject(sheet);
            }
        }

        public static void EnrichSpecificityDissolution(Worksheet sheet, int numPeaks, int numSamples, int numSolPeakPurity, string cmbProtocolType, string cmbProductType, string cmbTestType)
        {
            if (sheet != null)
            {
                bool wasProtected = WorksheetUtilities.SetSheetProtection(sheet, null, false);

                if (!string.IsNullOrEmpty(cmbProtocolType))
                {
                    sheet.Cells[4, 3] = cmbProtocolType;
                }
                if (!string.IsNullOrEmpty(cmbProductType))
                {
                    sheet.Cells[5, 3] = cmbProductType;
                }

                if (!string.IsNullOrEmpty(cmbTestType))
                {
                    sheet.Cells[6, 3] = cmbTestType;
                }

                if (numSamples > DefaulttbNumOfSamples)
                {
                    int rowsToInsert = numSamples - DefaulttbNumOfSamples;
                    int indexToInsertFrom = WorksheetUtilities.GetNamedRangeRowCount(sheet, "Dissolution_Sample");
                    AppendRowsCopyOnlyFormulasAndRenumber(sheet, "Dissolution_Sample", rowsToInsert, 1, 1, 3);

                    int indexToInsertSummaryFrom = WorksheetUtilities.GetNamedRangeRowCount(sheet, "Dissolution_Summary1");
                    WorksheetUtilities.InsertRowsIntoNamedRangeFromRow(rowsToInsert, sheet, "Dissolution_Summary1", false, XlDirection.xlDown, XlPasteType.xlPasteAll, indexToInsertSummaryFrom + 1);
                }

                try
                {
                    _app.Goto(sheet.Cells[1, 1], true);
                }
                catch
                {
                    Logger.LogMessage("Scroll of sheet failed in Specificity.Dissolution!", Level.Error);
                }

                if (wasProtected) WorksheetUtilities.SetSheetProtection(sheet, null, false);

                WorksheetUtilities.ReleaseComObject(sheet);
            }
        }

        public static void EnrichSpecificityWaterContent(Worksheet sheet, string cmbProtocolType, string cmbProductType, string cmbTestType)
        {
            if (sheet != null)
            {
                bool wasProtected = WorksheetUtilities.SetSheetProtection(sheet, null, false);

                if (!string.IsNullOrEmpty(cmbProtocolType))
                {
                    sheet.Cells[4, 3] = cmbProtocolType;
                }
                if (!string.IsNullOrEmpty(cmbProductType))
                {
                    sheet.Cells[5, 3] = cmbProductType;
                }

                if (!string.IsNullOrEmpty(cmbTestType))
                {
                    sheet.Cells[6, 3] = cmbTestType;
                }

                try
                {
                    _app.Goto(sheet.Cells[1, 1], true);
                }
                catch
                {
                    Logger.LogMessage("Scroll of sheet failed in Specificity.WaterContent!", Level.Error);
                }

                if (wasProtected) WorksheetUtilities.SetSheetProtection(sheet, null, false);

                WorksheetUtilities.ReleaseComObject(sheet);
            }
        }

        public static void EnrichSpecificityVolatiles(Worksheet sheet, int numPeaks, int numSamples, int numSolPeakPurity, bool isProdTypeDrugSubstance, bool isProdTypeSdd, bool isDrugProduct, string cmbProtocolType, string cmbProductType, string cmbTestType)
        {

            if (sheet != null)
            {
                bool wasProtected = WorksheetUtilities.SetSheetProtection(sheet, null, false);

                if (!string.IsNullOrEmpty(cmbProtocolType))
                {
                    sheet.Cells[6, 4] = cmbProtocolType;
                }
                if (!string.IsNullOrEmpty(cmbProductType))
                {
                    sheet.Cells[7, 4] = cmbProductType;
                }

                if (!string.IsNullOrEmpty(cmbTestType))
                {
                    sheet.Cells[8, 4] = cmbTestType;
                }

                if (numPeaks > DefaulttbNumOfPeaks)
                {
                    int rowsToInsert = numPeaks - DefaulttbNumOfPeaks;
                    int indexToInsertFrom = WorksheetUtilities.GetNamedRangeRowCount(sheet, "Default_Peak_Of_Interest_Table");

                    WorksheetUtilities.InsertRowsIntoNamedRangeFromRow(rowsToInsert, sheet, "Default_Peak_Of_Interest_Table", false, XlDirection.xlDown, XlPasteType.xlPasteAll, indexToInsertFrom);
                }

                if (isProdTypeDrugSubstance)
                {
                    if (numPeaks > DefaulttbNumOfPeaks)
                    {
                        int rowsToInsert = numPeaks - DefaulttbNumOfPeaks;
                        int indexToInsertFrom = WorksheetUtilities.GetNamedRangeRowCount(sheet, "Drug_Substance_Summary1");

                        WorksheetUtilities.InsertRowsIntoNamedRangeFromRow(rowsToInsert, sheet, "Drug_Substance_Summary1", false, XlDirection.xlDown, XlPasteType.xlPasteAll, indexToInsertFrom);
                    }

                    if (numSamples > DefaulttbNumOfSamples)
                    {
                        int rowsToInsert = numSamples - DefaulttbNumOfSamples;
                        int indexToInsertFrom = WorksheetUtilities.GetNamedRangeRowCount(sheet, "Drug_Substance_Sample");
                        AppendRowsCopyOnlyFormulasAndRenumber(sheet, "Drug_Substance_Sample", rowsToInsert, 1, 1, 3);
                        EnsureDrugSubstanceSummaries(sheet, "Drug_Substance_Summary1", "Drug_Substance_Sample", numSamples);

                    }
                }
                else
                {
                    WorksheetUtilities.DeleteNamedRangeRows(sheet, "Drug_Substance_Section");
                }

                if (isProdTypeSdd)
                {
                    if (numPeaks > DefaulttbNumOfPeaks)
                    {
                        int rowsToInsert = numPeaks - DefaulttbNumOfPeaks;
                        int indexToInsertFrom = WorksheetUtilities.GetNamedRangeRowCount(sheet, "SDD_Summary1");

                        WorksheetUtilities.InsertRowsIntoNamedRangeFromRow(rowsToInsert, sheet, "SDD_Summary1", false, XlDirection.xlDown, XlPasteType.xlPasteAll, indexToInsertFrom);
                    }

                    if (numSamples > DefaulttbNumOfSamples)
                    {
                        int rowsToInsert = numSamples - DefaulttbNumOfSamples;
                        int indexToInsertFrom = WorksheetUtilities.GetNamedRangeRowCount(sheet, "SDD_Sample");
                        AppendRowsCopyOnlyFormulasAndRenumber(sheet, "SDD_Sample", rowsToInsert, 1, 1, 3, isProdTypeSdd);
                        EnsureDrugSubstanceSummaries(sheet, "SDD_Summary1", "SDD_Sample", numSamples);
                    }
                }
                else
                {
                    WorksheetUtilities.DeleteNamedRangeRows(sheet, "SDD_Section");
                }

                if (isDrugProduct && cmbTestType == "Volatiles")
                {
                    if (numPeaks > DefaulttbNumOfPeaks)
                    {
                        int rowsToInsert = numPeaks - DefaulttbNumOfPeaks;
                        int indexToInsertFrom = WorksheetUtilities.GetNamedRangeRowCount(sheet, "Drug_Product_Summary1");

                        WorksheetUtilities.InsertRowsIntoNamedRangeFromRow(rowsToInsert, sheet, "Drug_Product_Summary1", false, XlDirection.xlDown, XlPasteType.xlPasteAll, indexToInsertFrom);
                    }

                    if (numSamples > DefaulttbNumOfSamples)
                    {
                        int rowsToInsert = numSamples - DefaulttbNumOfSamples;
                        int indexToInsertFrom = WorksheetUtilities.GetNamedRangeRowCount(sheet, "Drug_Product_Sample");
                        AppendRowsCopyOnlyFormulasAndRenumber(sheet, "Drug_Product_Sample", rowsToInsert, 1, 1, 3);
                        EnsureDrugSubstanceSummaries(sheet, "Drug_Product_Summary1", "Drug_Product_Sample", numSamples);

                    }
                }
                else
                {
                    WorksheetUtilities.DeleteNamedRangeRows(sheet, "Drug_Product_Section");
                }

                try
                {
                    _app.Goto(sheet.Cells[1, 1], true);
                }
                catch
                {
                    Logger.LogMessage("Scroll of sheet failed in Specificity.Volatiles!", Level.Error);
                }

                if (wasProtected) WorksheetUtilities.SetSheetProtection(sheet, null, false);

                WorksheetUtilities.ReleaseComObject(sheet);
            }
        }


        public static void AppendRowsCopyOnlyFormulasAndRenumber(
        Worksheet sheet,
        string namedRange,
        int rowsToAdd,
        int headerRows = 1,
        int templateRowIndex = 1,
        int sampleColIndex = 1,
        bool isProdTypeSdd = false)
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

                    if (isProdTypeSdd)
                    {
                        Range sddCell = rng.Cells[headerRows + i, sampleColIndex + 2] as Range;
                        sddCell.Value = "SDD";
                    }
                }
            }
            finally
            {

            }
        }

        public static void EnsureDrugSubstanceSummaries(
            Worksheet sheet,
            string baseSummaryName,
            string baseSampleName,
            int numSamples)
        {
            if (sheet == null) throw new System.ArgumentNullException(nameof(sheet));
            if (string.IsNullOrWhiteSpace(baseSummaryName)) throw new System.ArgumentException("Base name required.", nameof(baseSummaryName));
            if (numSamples < 1) return;

            var m = System.Text.RegularExpressions.Regex.Match(baseSummaryName, @"^(.*?)(\d+)$");
            if (!m.Success) throw new System.ArgumentException("Base name must end with a number, e.g. Drug_Substance_Summary1");
            string prefix = m.Groups[1].Value;
            int baseIndex = int.Parse(m.Groups[2].Value);

            var nmBase = sheet.Names.Item(baseSummaryName, Type.Missing, Type.Missing) as Microsoft.Office.Interop.Excel.Name
                         ?? throw new System.InvalidOperationException($"Named range '{baseSummaryName}' not found on sheet '{sheet.Name}'.");
            Range baseRange = nmBase.RefersToRange
                           ?? throw new System.InvalidOperationException($"'{baseSummaryName}' has no RefersToRange.");

            var nmSample = sheet.Names.Item(baseSampleName, Type.Missing, Type.Missing) as Microsoft.Office.Interop.Excel.Name
                           ?? throw new System.InvalidOperationException($"Named range '{baseSampleName}' not found.");
            Range sampleRange = nmSample.RefersToRange
                              ?? throw new System.InvalidOperationException($"'{baseSampleName}' has no RefersToRange.");

            int blockHeight = baseRange.Rows.Count;
            int blockWidth = baseRange.Columns.Count;
            int leftCol = baseRange.Column;

            int relHeaderRow = -1, relHeaderCol = -1;
            string headerBaseFormula = null;

            var rxRef = new System.Text.RegularExpressions.Regex(@"\$?([A-Z]+)\$?(\d+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            string headerColLetter = null;

            for (int r = 1; r <= blockHeight && headerBaseFormula == null; r++)
            {
                for (int c = 1; c <= blockWidth && headerBaseFormula == null; c++)
                {
                    var cell = baseRange.Cells[r, c] as Microsoft.Office.Interop.Excel.Range;
                    try
                    {
                        if (cell != null && (bool)(cell.HasFormula ?? false))
                        {
                            string f = cell.Formula;
                            var mm = rxRef.Match(f);
                            if (mm.Success)
                            {
                                relHeaderRow = r;
                                relHeaderCol = c;
                                headerBaseFormula = f;
                                headerColLetter = mm.Groups[1].Value.ToUpper();
                            }
                        }
                    }
                    finally { if (cell != null) while (System.Runtime.InteropServices.Marshal.ReleaseComObject(cell) > 0) { } }
                }
            }

            int existingMax = 0;
            foreach (Microsoft.Office.Interop.Excel.Name n in sheet.Names)
            {
                var mm = System.Text.RegularExpressions.Regex.Match(n.Name, $"^{System.Text.RegularExpressions.Regex.Escape(prefix)}(\\d+)$");
                if (mm.Success)
                {
                    int idx = int.Parse(mm.Groups[1].Value);
                    if (idx > existingMax) existingMax = idx;
                }
            }
            if (existingMax == 0) existingMax = baseIndex;
            if (existingMax >= numSamples) return;

            Range lastBlock = (existingMax == baseIndex)
                ? baseRange
                : (sheet.Names.Item(prefix + existingMax, Type.Missing, Type.Missing) as Microsoft.Office.Interop.Excel.Name)?.RefersToRange
                  ?? throw new System.InvalidOperationException($"Could not resolve last block '{prefix}{existingMax}'.");

            for (int next = existingMax + 1; next <= numSamples; next++)
            {
                int insertRow = lastBlock.Row + lastBlock.Rows.Count;
                var insertBand = sheet.Range[
                    sheet.Rows[insertRow, Type.Missing],
                    sheet.Rows[insertRow + blockHeight - 1, Type.Missing]
                ];
                insertBand.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown, Type.Missing);
                while (System.Runtime.InteropServices.Marshal.ReleaseComObject(insertBand) > 0) { }

                var destTopLeft = sheet.Cells[lastBlock.Row + lastBlock.Rows.Count, leftCol] as Microsoft.Office.Interop.Excel.Range;
                var destBlock = sheet.Range[
                    destTopLeft,
                    sheet.Cells[destTopLeft.Row + blockHeight - 1, destTopLeft.Column + blockWidth - 1]
                ];

                baseRange.Copy(Type.Missing);
                destBlock.PasteSpecial(
                    Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll,
                    Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone,
                    false, false
                );
                sheet.Application.CutCopyMode = 0;

                if (headerBaseFormula != null && relHeaderRow > 0 && relHeaderCol > 0 && !string.IsNullOrEmpty(headerColLetter))
                {
                    int sampleFirstDataRow = sampleRange.Row + 1;
                    int targetRow = sampleFirstDataRow + (next - 1);
                    var destHeaderCell = destBlock.Cells[relHeaderRow, relHeaderCol] as Microsoft.Office.Interop.Excel.Range;
                    try
                    {
                        var rxFirstRef = new System.Text.RegularExpressions.Regex(@"\$?[A-Z]+\$?\d+");
                        string newFormula = rxFirstRef.Replace(headerBaseFormula, $"{headerColLetter}{targetRow}", 1);
                        destHeaderCell.Formula = newFormula;
                    }
                    finally { if (destHeaderCell != null) while (System.Runtime.InteropServices.Marshal.ReleaseComObject(destHeaderCell) > 0) { } }
                }

                string newName = prefix + next;
                try
                {
                    var existing = sheet.Names.Item(newName, Type.Missing, Type.Missing) as Microsoft.Office.Interop.Excel.Name;
                    if (existing != null)
                    {
                        existing.Delete();
                        while (System.Runtime.InteropServices.Marshal.ReleaseComObject(existing) > 0) { }
                    }
                }
                catch { }

                sheet.Names.Add(newName, destBlock);

                if (!object.ReferenceEquals(lastBlock, baseRange)) while (System.Runtime.InteropServices.Marshal.ReleaseComObject(lastBlock) > 0) { }
                lastBlock = destBlock;

                if (destTopLeft != null) while (System.Runtime.InteropServices.Marshal.ReleaseComObject(destTopLeft) > 0) { }
            }

            if (!object.ReferenceEquals(lastBlock, baseRange) && lastBlock != null)
                while (System.Runtime.InteropServices.Marshal.ReleaseComObject(lastBlock) > 0) { }
        }

        private static void TryDeleteSheetName(Worksheet sheet, string name)
        {
            try
            {
                var n = sheet.Names.Item(name, Type.Missing, Type.Missing) as Name;
                if (n != null)
                {
                    n.Delete();
                    Marshal.ReleaseComObject(n);
                }
            }
            catch { }
        }

        public static void EnsurePeakSummaryRowsWithFixedArgs(
        Worksheet sheet,
        int numPeaks,
        string namedRange)
        {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (numPeaks < 1) numPeaks = 1;

            Name nm = null;
            Range rng = null;
            Range headerRow = null;
            Range templateRow = null;

            try
            {
                var decimalRange = sheet.Names.Item($"{namedRange}_Decimals");
                int decimalsRow = decimalRange.RefersToRange.Row;

                nm = sheet.Names.Item(namedRange, Type.Missing, Type.Missing) as Name
                     ?? throw new InvalidOperationException($"Named range '{namedRange}' not found.");
                rng = nm.RefersToRange ?? throw new InvalidOperationException($"Named range '{namedRange}' has no RefersToRange.");

                const int headerRows = 2;
                int totalCols = rng.Columns.Count;
                int totalRows = rng.Rows.Count;
                if (totalRows <= headerRows)
                    throw new InvalidOperationException("Summary range has no data row under the header.");

                headerRow = rng.Rows[2, Type.Missing] as Range;
                int angleColInRange = -1, thresholdColInRange = -1;

                for (int c = 1; c <= totalCols; c++)
                {
                    var hCell = headerRow.Cells[1, c] as Range;
                    string text = (hCell?.Text ?? "").ToString().Trim();
                    if (angleColInRange == -1 && text.Equals("Purity Angle", StringComparison.OrdinalIgnoreCase))
                        angleColInRange = c;
                    if (thresholdColInRange == -1 && text.Equals("Purity threshold", StringComparison.OrdinalIgnoreCase))
                        thresholdColInRange = c;
                    if (hCell != null) Marshal.ReleaseComObject(hCell);
                }
                if (angleColInRange == -1 || thresholdColInRange == -1)
                    throw new InvalidOperationException("Could not find 'Purity Angle' and/or 'Purity threshold' columns in the header.");


                int firstDataRowInRange = headerRows + 1;
                templateRow = rng.Rows[firstDataRowInRange, Type.Missing] as Range;


                string[] tplR1C1 = new string[totalCols];
                object[] tplVals = new object[totalCols];
                for (int c = 1; c <= totalCols; c++)
                {
                    var cell = templateRow.Cells[1, c] as Range;
                    bool hasFormula = false;
                    try { hasFormula = (bool)(cell?.HasFormula ?? false); } catch { }
                    tplR1C1[c - 1] = hasFormula ? (string)cell.FormulaR1C1 : null;
                    tplVals[c - 1] = cell?.Value2;
                    if (cell != null) Marshal.ReleaseComObject(cell);
                }


                int currentDataRows = totalRows - headerRows;
                int rowsToAdd = Math.Max(0, numPeaks - currentDataRows);
                if (rowsToAdd > 0)
                {
                    int insertAt = rng.Row + totalRows;
                    var insertBlock = sheet.Range[sheet.Rows[insertAt, Type.Missing],
                                                  sheet.Rows[insertAt + rowsToAdd - 1, Type.Missing]];
                    insertBlock.EntireRow.Insert(XlInsertShiftDirection.xlShiftDown);
                    Marshal.ReleaseComObject(insertBlock);

                    Range startCell = sheet.Cells[rng.Row, rng.Column] as Range;
                    Range endCell = sheet.Cells[rng.Row + totalRows + rowsToAdd - 1, rng.Column + totalCols - 1] as Range;
                    string refersTo = $"='{sheet.Name}'!{sheet.Range[startCell, endCell].get_Address(true, true, XlReferenceStyle.xlA1)}";
                    nm.RefersTo = refersTo;
                    Marshal.ReleaseComObject(startCell);
                    Marshal.ReleaseComObject(endCell);


                    rng = nm.RefersToRange;
                }


                for (int i = 1; i <= numPeaks; i++)
                {
                    int rowInRange = headerRows + i;
                    Range destRow = rng.Rows[rowInRange, Type.Missing] as Range;


                    templateRow.Copy(Type.Missing);
                    destRow.PasteSpecial(XlPasteType.xlPasteFormats,
                                         XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);


                    for (int c = 1; c <= totalCols; c++)
                    {
                        Range destCell = destRow.Cells[1, c] as Range;

                        if (c == angleColInRange)
                        {
                            int sheetRow = destRow.Row;
                            string jAddr = sheet.Cells[sheetRow, 10].Address(false, false); // col 10 = J
                            string fAddr = sheet.Cells[decimalsRow, 6].Address(false, false);
                            destCell.Formula = $"=FIXED({jAddr}, {fAddr})";
                        }
                        else if (c == thresholdColInRange)
                        {
                            int sheetRow = destRow.Row;
                            string kAddr = sheet.Cells[sheetRow, 11].Address(false, false); // col 11 = K
                            string gAddr = sheet.Cells[decimalsRow, 7].Address(false, false);
                            destCell.Formula = $"=FIXED({kAddr}, {gAddr})";
                        }
                        else
                        {

                            if (!string.IsNullOrEmpty(tplR1C1[c - 1]))
                                destCell.FormulaR1C1 = tplR1C1[c - 1];
                            else
                                destCell.Value2 = tplVals[c - 1];
                        }

                        Marshal.ReleaseComObject(destCell);
                    }

                    Marshal.ReleaseComObject(destRow);
                }

                sheet.Application.CutCopyMode = 0;
            }
            finally
            {
                if (templateRow != null) while (Marshal.ReleaseComObject(templateRow) > 0) { }
                if (headerRow != null) while (Marshal.ReleaseComObject(headerRow) > 0) { }
                if (rng != null) while (Marshal.ReleaseComObject(rng) > 0) { }
                if (nm != null) while (Marshal.ReleaseComObject(nm) > 0) { }
            }
        }
    }//End of Class
}
