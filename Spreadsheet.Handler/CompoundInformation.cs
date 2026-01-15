using log4net.Core;
using Microsoft.Office.Interop.Excel;
using Internal.Framework.Collections.Extensions;
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace Spreadsheet.Handler
{
    public static class CompoundInformation
    {
        private static Application _app;
        private const string TempDirectoryName = "ABD_TempFiles";


        public static string UpdateCompoundInformationSheet(
            string sourcePath,
            int txtReferenceStandard,
            int txtImpurityMixture,
            int txtDrugSubstance,
            int txtSDD,
            int txtDrugProduct,
            int txtPlacebo,
            int txtPolymer,
            int txtIndividualImpurity,
            int txtVolatiles,
            string cmbProtocolType,
            string cmbProductType,
            string cmbTestType)
        {
            string returnPath = "";
            try
            {
                //returnPath = UpdateCompoundInformationSheet2(
                //    sourcePath,
                //    txtReferenceStandard,
                //    txtImpurityMixture,
                //    txtDrugSubstance,
                //    txtSDD,
                //    txtDrugProduct,
                //    txtPlacebo,
                //    txtPolymer,
                //    txtIndividualImpurity,
                //    txtVolatiles,
                //    cmbProtocolType,
                //    cmbProductType,
                //    cmbTestType);
            }
            catch (Exception ex)
            {
                Logger.LogMessage("Error in CompoundInformation.UpdateCompoundInformationSheet: " +
                    ex.Message + "\r\n" + ex.StackTrace, Level.Error);
            }
            finally
            {
                WorksheetUtilities.ReleaseExcelApp();
            }

            return returnPath;
        }

        public static string UpdateCompoundInformationSheet2(
            string sourcePath,
            int txtReferenceStandard,
            int txtImpurityMixture,
            int txtDrugSubstance,
            int txtSDD,
            int txtDrugProduct,
            int txtPlacebo,
            int txtPolymer,
            int txtIndividualImpurity,
            int txtVolatiles,
            string cmbProtocolType,
            string cmbProductType,
            string cmbTestType,int txtimpurity1)
        {

            if (!File.Exists(sourcePath))
            {
                Logger.LogMessage("Invalid source file path specified.", Level.Error);
                return "";
            }
            // Generate an random temp path to save new workbook
            string savePath = WorksheetUtilities.CopyWorkbook(sourcePath, TempDirectoryName, "Compound Information Testing.xlsx");
            if (String.IsNullOrEmpty(savePath)) return "";
            // for testing sample inserting

            //Application app = new Application();
            //Workbook book1 = app.Workbooks.Open(path);
            //Worksheet sheet1 = (Worksheet)book1.Worksheets[1];
            //app.Visible = true; // just for testing

            //// List all named ranges
            //foreach (Name n in book1.Names)
            //    Console.WriteLine($"{n.Name} → {n.RefersTo}");

            //// Try to replicate named range “Reference_Standard_1”
            //Name range = book1.Names.Cast<Name>()
            //    .FirstOrDefault(n => n.Name == "Reference_Standard_1");

            //if (range != null)
            //{
            //    Range r = range.RefersToRange;
            //    r.Copy();
            //    Range nextRow = sheet1.Rows[r.Row + r.Rows.Count];
            //    nextRow.Insert(XlInsertShiftDirection.xlShiftDown);
            //    sheet1.Rows[r.Row + r.Rows.Count].PasteSpecial(XlPasteType.xlPasteAll);
            //}

            //book1.Save();
            //book1.Close();
            //app.Quit();
            //testing end statement

            Workbook book = null;
            Worksheet sheet = null;

            try
            {
                // ✅ Open workbook
                _app = WorksheetUtilities.GetExcelApp();
                //book = (Workbook)_app.Workbooks.Open(
                //    sourcePath, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing,
                //    Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing,
                //    Type.Missing, Type.Missing, Type.Missing);
                book = (Workbook)_app.Workbooks.Open(
                    savePath, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing);

                book.Application.DisplayAlerts = false;
                Logger.LogMessage($"Workbook '{book.Name}' opened successfully.", Level.Info);

                // ✅ Assume one sheet ("Compound Information")
                sheet = (Worksheet)book.Worksheets[1];
                Logger.LogMessage($"Active sheet: {sheet.Name}", Level.Info);

                // ✅ Write the combo-box values
                sheet.Range["C2"].Value = cmbProtocolType;
                sheet.Range["C3"].Value = cmbProductType;
                sheet.Range["C4"].Value = cmbTestType;

                // ✅ Replicate rows by named ranges
                ReplicateNamedRangeRows(book, "Reference_Standard_1", txtReferenceStandard);
                ReplicateNamedRangeRows(book, "Impurity_Mixture_1", txtImpurityMixture);
                ReplicateNamedRangeRows(book, "Drug_Substance_1", txtDrugSubstance);
                ReplicateNamedRangeRows(book, "SDD_1", txtSDD);
                ReplicateNamedRangeRows(book, "Drug_Product_1", txtDrugProduct);
                ReplicateNamedRangeRows(book, "Placebo_1", txtPlacebo);
                ReplicateNamedRangeRows(book, "Polymer_1", txtPolymer);
                ReplicateNamedRangeRows(book, "Impurity_1", txtIndividualImpurity);
                ReplicateNamedRangeRows(book, "Volatiles_1", txtVolatiles);
               // ReplicateNamedRangeRows(book, "Impurity_1", txtimpurity1);

                // ✅ Save workbook (overwrite same file)
                book.Save();
                string fullPath = book.FullName;
                Logger.LogMessage($"Workbook '{book.Name}' saved successfully at: {fullPath}", Level.Info);

                return fullPath;
            }
            catch (Exception ex)
            {
                Logger.LogMessage("Error in CompoundInformation.UpdateCompoundInformationSheet2: " +
                    ex.Message + "\r\n" + ex.StackTrace, Level.Error);
                return "";
            }
            finally
            {
                try
                {
                    if (book != null)
                    {
                        book.Close(false);
                        Marshal.ReleaseComObject(sheet);
                        Marshal.ReleaseComObject(book);
                    }
                }
                catch { }

                WorksheetUtilities.ReleaseExcelApp();
            }
        }

        public static void ReplicateNamedRangeRows(Workbook book, string rangeName, int count)
        {
          

            if (count <= 1) return; // nothing to replicate if only one

            try
            {
                // ✅ Try to get named range (workbook- or sheet-level)
                Name namedRange = book.Names.Cast<Name>()
                    .FirstOrDefault(n => n.Name.Equals(rangeName, StringComparison.OrdinalIgnoreCase));

                if (namedRange == null || namedRange.RefersToRange == null)
                {
                    // fallback — sheet scoped
                    foreach (Worksheet ws in book.Worksheets)
                    {
                        try
                        {
                            foreach (Name n in ws.Names)
                            {
                                // Excel prefixes sheet-scoped names as 'SheetName'!RangeName internally
                                string cleanName = n.Name;
                                if (cleanName.Contains("!"))
                                    cleanName = cleanName.Substring(cleanName.IndexOf("!") + 1);

                                if (cleanName.Equals(rangeName, StringComparison.OrdinalIgnoreCase))
                                {
                                    namedRange = n;
                                    Logger.LogMessage($"✅ Found sheet-scoped named range '{rangeName}' in sheet '{ws.Name}'", Level.Info);
                                    break;
                                }
                            }
                            if (namedRange != null) break;
                        }
                        catch { }
                    }
                }

                if (namedRange == null || namedRange.RefersToRange == null)
                {
                    Logger.LogMessage($"❌ Named range '{rangeName}' not found.", Level.Info);
                    return;
                }

                Range sourceRange = namedRange.RefersToRange;
                Worksheet sheet = sourceRange.Worksheet;

                int startRow = sourceRange.Row;
                int startCol = sourceRange.Column;
                int rowCount = sourceRange.Rows.Count;
                int colCount = sourceRange.Columns.Count;

                Logger.LogMessage($"✅ Found range '{rangeName}' at {sheet.Name}!R{startRow}C{startCol}, size {rowCount}x{colCount}", Level.Info);

                // ✅ Loop to add rows and copy data for (count - 1)
                for (int i = 1; i < count; i++)
                {
                    int insertRowIndex = startRow + (i * rowCount);

                    Range insertTarget = sheet.Rows[insertRowIndex];
                    insertTarget.Insert(
                        XlInsertShiftDirection.xlShiftDown,
                        XlInsertFormatOrigin.xlFormatFromRightOrBelow);


                    // Define target range of same size as original
                    Range destRange = sheet.Range[
                        sheet.Cells[insertRowIndex, startCol],
                        sheet.Cells[insertRowIndex + rowCount - 1, startCol + colCount - 1]
                    ];

                    // ✅ Copy-paste all (values, formatting, formulas)
                    //sourceRange.Copy(destRange);

                    sourceRange.Copy();
                    destRange.PasteSpecial(XlPasteType.xlPasteAll);
                    destRange.PasteSpecial(XlPasteType.xlPasteFormats);
                    sheet.Application.CutCopyMode = XlCutCopyMode.xlCopy;


                    // 👇 ADD THIS CALL for adding new row in dependent sheet
                    UpdateDependentSheetsForNewRow(book, insertRowIndex);




                    //for adding auto increment if it is impurity_1
                    // ✅ Auto-increment first column only for "Impurity_1"
                    if (rangeName.Equals("Impurity_1", StringComparison.OrdinalIgnoreCase) ||
    rangeName.Equals("Volatiles_1", StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            // Assume the first column of the block holds the serial number
                            Range serialCell = sheet.Cells[insertRowIndex, startCol];

                            // Get the previous cell’s number and increment
                            Range prevSerialCell = sheet.Cells[insertRowIndex - rowCount, startCol];
                            object prevValue = prevSerialCell.Value;

                            int newSerial = 1;
                            if (prevValue != null && int.TryParse(prevValue.ToString(), out int prevNum))
                                newSerial = prevNum + 1;

                            serialCell.Value = newSerial;
                            // ✅ Update TEXTJOIN formula automatically
                            // Locate the TEXTJOIN formula cell by name or by pattern
                            Range usedRange = sheet.UsedRange;

                           
 
  

  
  


                        }
                        catch (Exception ex)
                        {
                            Logger.LogMessage($"⚠️ Error auto-incrementing serial for '{rangeName}': {ex.Message}", Level.Error);
                        }
                    }


                    //end of impurity thing


                    //for adding new rows in the another sheet to check formulas and insert
                   // ✅ Add corresponding rows in dependent sheets
foreach (Worksheet ws in book.Worksheets)
{
    if (ws.Name == "Compound Information") continue;

    Range usedRange = ws.UsedRange;
    int headerEndRow = 0;

    // Step 1: Identify header region properly
    foreach (Range row in ws.UsedRange.Rows)
    {
        bool likelyHeader = false;
        foreach (Range cell in row.Columns)
        {
            string val = cell.Value2?.ToString() ?? "";
            if (val.IndexOf("Validation", StringComparison.OrdinalIgnoreCase) >= 0 ||
                val.IndexOf("Parameter", StringComparison.OrdinalIgnoreCase) >= 0 ||
                val.IndexOf("Compound", StringComparison.OrdinalIgnoreCase) >= 0 ||
                val.IndexOf("Property", StringComparison.OrdinalIgnoreCase) >= 0 ||
                val.IndexOf("Specification", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                likelyHeader = true;
                break;
            }
        }
        if (likelyHeader)
        {
            headerEndRow = row.Row;
        }
        else if (headerEndRow > 0)
        {
            break; // first non-header row found
        }
    }
    if (headerEndRow == 0) headerEndRow = 10;

    // Step 2: Find formula referencing 'Compound Information'
    Range formulaRow = null;

                        //commented ofr testing new code for dependent exact formulas
    //foreach (Range cell in usedRange)
    //{
    //    if (cell.HasFormula && cell.Formula.Contains("'Compound Information'!"))
    //    {
    //        formulaRow = ws.Rows[cell.Row];
    //        break;
    //    }
    //}

    //if (formulaRow == null) continue;

    //// Step 3: Insert one new row after the formula section
    //int depInsertRowIndex = formulaRow.Row + 1;
    //Range depInsertTarget = ws.Rows[depInsertRowIndex];
    //depInsertTarget.Insert(XlInsertShiftDirection.xlShiftDown, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);

    //// Step 4: Copy formulas from previous row
    //Range prevRow = ws.Rows[depInsertRowIndex - 1];
    //Range newRow = ws.Rows[depInsertRowIndex];
    //prevRow.Copy();
    //newRow.PasteSpecial(XlPasteType.xlPasteFormulas);
    //ws.Application.CutCopyMode = XlCutCopyMode.xlCopy;

    //Logger.LogMessage($"🔁 Inserted row below formula section in '{ws.Name}' (avoided headers) at row {depInsertRowIndex}", Level.Info);
    //commented ofr testing new code for dependent exact formulas
}


                }
                // 🔁 After all rows have been inserted
if (rangeName.Equals("Impurity_1", StringComparison.OrdinalIgnoreCase))
{
    try
    {
        Worksheet ws = sheet;

        // Find existing Impurity TEXTJOIN cell in column C (same as you had before)
        Range formulaCells = null;
        try { formulaCells = ws.UsedRange.SpecialCells(XlCellType.xlCellTypeFormulas); } catch { }

        Range formulaCell = null;
        if (formulaCells != null)
        {
            foreach (Range cell in formulaCells)
            {
                string f = cell.Formula?.ToString() ?? "";

                if (cell.Column != 3) continue; // C
                if (f.IndexOf("TEXTJOIN", StringComparison.OrdinalIgnoreCase) < 0) continue;
                if (!Regex.IsMatch(f, @"(\b|\W)J\d+", RegexOptions.IgnoreCase)) continue;

                formulaCell = cell;
                break;
            }
        }

        if (formulaCell != null)
        {
            // Build J list from Impurity_1 rows
            Range impRange = ws.Range["Impurity_1"];
            int firstRow = impRange.Row;
            int totalRows = count; // total Impurity entries

            var refList = new System.Text.StringBuilder();
            for (int r = 0; r < totalRows; r++)
            {
                if (r > 0) refList.Append(",");
                refList.Append($"J{firstRow + r}");
            }

            int lastRow = firstRow + totalRows - 1;
            string newFormula =
                $"=TEXTJOIN(\", \",TRUE,{refList},\"and \" & J{lastRow})";

            formulaCell.Formula = newFormula;
            Logger.LogMessage($"📈 Updated Impurity TEXTJOIN in {ws.Name}!{formulaCell.Address} → {newFormula}", Level.Info);
        }
    }
    catch (Exception ex)
    {
        Logger.LogMessage($"⚠️ Error updating Impurity TEXTJOIN: {ex.Message}", Level.Error);
    }
}
else if (rangeName.Equals("Volatiles_1", StringComparison.OrdinalIgnoreCase))
{
    try
    {
                        Worksheet ws = sheet;




                        // 🧪 4) Put formula ONLY in the new cell (this is what you asked for)

                        //new process


                        // Find existing Impurity TEXTJOIN cell in column C (same as you had before)
                        Range formulaCells = null;
                        try { formulaCells = ws.UsedRange.SpecialCells(XlCellType.xlCellTypeFormulas); } catch { }
                        Range formulaCell = null;
                        if (formulaCells != null)
                        {
                            foreach (Range cell in formulaCells)
                            {
                                string f = cell.Formula?.ToString() ?? "";

                                if (cell.Column != 3) continue; // C
                                if (f.IndexOf("TEXTJOIN", StringComparison.OrdinalIgnoreCase) < 0) continue;
                                if (!Regex.IsMatch(f, @"(\b|\W)C\d+", RegexOptions.IgnoreCase)) continue;

                                formulaCell = cell;
                                break;
                            }
                        }
                        if (formulaCell != null)
                        {
                            // Build J list from Impurity_1 rows
                            Range impRange = ws.Range["Volatiles_1"];
                            int firstRow = impRange.Row;
                            int totalRows = count; // total Impurity entries

                            var refList = new System.Text.StringBuilder();
                            for (int r = 0; r < totalRows-1; r++)
                            {
                                if (r > 0) refList.Append(",");
                                refList.Append($"C{firstRow + r}");
                            }

                            int lastRow = firstRow + totalRows - 1;
                            string newFormula =
                                $"=TEXTJOIN(\", \",TRUE,{refList},\"and \" & C{lastRow})";

                            formulaCell.Formula = newFormula;
                            Logger.LogMessage($"📈 Updated Impurity TEXTJOIN in {ws.Name}!{formulaCell.Address} → {newFormula}", Level.Info);
                        }

                        //new process end
                        //formulaCell.Formula = newFormula;
       // Logger.LogMessage($"📈 Inserted Volatiles TEXTJOIN in {ws.Name}!C{formulaRow} → {newFormula}", Level.Info);
    }
    catch (Exception ex)
    {
        Logger.LogMessage($"⚠️ Error updating Volatiles TEXTJOIN: {ex.Message}", Level.Error);
    }
}




                // ✅ Final table bottom border for Polymer_1 (only once after loop)
                if (rangeName.Equals("Polymer_1", StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        int lastRow = startRow + (count * rowCount) - 1;

                        // Define the columns for left and right tables
                        int leftStartCol = 3;  // C
                        int leftEndCol = 9;    // I
                        int rightStartCol = 10; // J
                        int rightEndCol = 12;   // L

                        // 🟦 Step 1: Remove all horizontal borders between Polymer rows for J:L
                        for (int r = startRow; r < lastRow; r++)
                        {
                            Range rightPolymerRow = sheet.Range[
                                sheet.Cells[r, rightStartCol],
                                sheet.Cells[r, rightEndCol]
                            ];
                            rightPolymerRow.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlLineStyleNone;
                        }
                        // 🟨 Step 2: Ensure left side Polymer (C:I) retains its normal look
                        Range leftPolymerBlock = sheet.Range[
                            sheet.Cells[startRow, leftStartCol],
                            sheet.Cells[lastRow, leftEndCol]
                        ];
                        leftPolymerBlock.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                        leftPolymerBlock.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;

                        // 🟩 Step 3: Add a single clean bottom border across both sides (C:L)
                        Range fullBottomRow = sheet.Range[
                            sheet.Cells[lastRow, leftStartCol],
                            sheet.Cells[lastRow, rightEndCol]
                        ];
                        fullBottomRow.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                        fullBottomRow.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;




                        Logger.LogMessage("📏 Added bottom border up to column I (left table) and L (right table) for Polymer.", Level.Info);
                    }
                    catch (Exception ex)
                    {
                        Logger.LogMessage($"⚠️ Error adding closing border for Polymer: {ex.Message}", Level.Error);
                    }
                
            }

               
        

        Logger.LogMessage($"✅ Replicated named range '{rangeName}' for total {count} blocks.", Level.Info);
            }
            catch (Exception ex)
            {
                Logger.LogMessage($"Error replicating rows for '{rangeName}': {ex.Message}\r\n{ex.StackTrace}", Level.Error);
            }

        }
        // *** RED *** new helper
        private static void UpdateDependentSheetsForNewRow(Workbook book, int mainRowIndex) // *** RED ***
        {
            const string mainSheetName = "Compound Information";                    // *** RED ***
            const string mainSheetPrefix = "'Compound Information'!";               // *** RED ***

            foreach (Worksheet ws in book.Worksheets)                               // *** RED ***
            {                                                                       // *** RED ***
                if (ws.Name == mainSheetName)                                       // *** RED ***
                    continue;                                                       // *** RED ***

                Range usedRange = ws.UsedRange;                                     // *** RED ***
                if (usedRange == null)                                              // *** RED ***
                    continue;                                                       // *** RED ***

                int templateRowIndex = 0;                                           // *** RED ***

                // --- find the template row in this dependent sheet ---           // *** RED ***
                // We look in column B for a formula pointing to row (mainRowIndex - 1) // *** RED ***
                int firstRow = usedRange.Row;                                       // *** RED ***
                int lastRow = usedRange.Row + usedRange.Rows.Count - 1;            // *** RED ***

                for (int r = firstRow; r <= lastRow; r++)                           // *** RED ***
                {                                                                   // *** RED ***
                    Range cellB = (Range)ws.Cells[r, 2]; // column B               // *** RED ***
                    if (!cellB.HasFormula)                                          // *** RED ***
                        continue;                                                   // *** RED ***

                    string f = cellB.Formula?.ToString() ?? "";                     // *** RED ***
                    if (f.IndexOf(mainSheetPrefix, StringComparison.OrdinalIgnoreCase) < 0) // *** RED ***
                        continue;                                                   // *** RED ***

                    // Look for any reference like 'Compound Information'!J40 or !M40 // *** RED ***
                    var match = Regex.Match(                                       // *** RED ***
                        f,                                                          // *** RED ***
                        @"'Compound Information'!\$?[A-Z]+\$?(\d+)",                // *** RED ***
                        RegexOptions.IgnoreCase);                                   // *** RED ***

                    if (!match.Success)                                             // *** RED ***
                        continue;                                                   // *** RED ***

                    if (!int.TryParse(match.Groups[1].Value, out int rowInMain))    // *** RED ***
                        continue;                                                   // *** RED ***

                    // Template row is the one that currently points to mainRowIndex - 1 // *** RED ***
                    if (rowInMain == mainRowIndex - 1)                              // *** RED ***
                    {                                                               // *** RED ***
                        templateRowIndex = r;                                       // *** RED ***
                                                                                    // do NOT break; if multiple rows qualify, we keep the last, // *** RED ***
                                                                                    // which is typically closest to the end of the block       // *** RED ***
                    }                                                               // *** RED ***
                }                                                                   // *** RED ***

                if (templateRowIndex == 0)                                          // *** RED ***
                    continue;                                                       // *** RED ***

                // --- insert a new row just below templateRowIndex ---             // *** RED ***
                Range insertRow = ws.Rows[templateRowIndex + 1];                    // *** RED ***
                insertRow.Insert(                                                   // *** RED ***
                    XlInsertShiftDirection.xlShiftDown,                             // *** RED ***
                    XlInsertFormatOrigin.xlFormatFromLeftOrAbove);                  // *** RED ***

                // After insert, the new blank row is at templateRowIndex + 1       // *** RED ***
                Range newRow = ws.Rows[templateRowIndex + 1];                  // *** RED ***
                Range templateRow = ws.Rows[templateRowIndex];                      // *** RED ***

                // Copy entire row (formulas + formats)                             // *** RED ***
                templateRow.Copy(newRow);                                           // *** RED ***
                ws.Application.CutCopyMode = XlCutCopyMode.xlCopy;                  // *** RED ***

                // --- fix formulas in column B and column M ---                    // *** RED ***
                Range newCellB = (Range)ws.Cells[newRow.Row, 2];   // column B      // *** RED ***
                Range newCellM = (Range)ws.Cells[newRow.Row, 13];  // column M      // *** RED ***

                AdjustDependentFormulaCell(newCellB, mainRowIndex);                 // *** RED ***
                AdjustDependentFormulaCell(newCellM, mainRowIndex);                 // *** RED ***
            }                                                                       // *** RED ***
        }
        // *** RED *** new helper
        private static void AdjustDependentFormulaCell(Range cell, int mainRowIndex) // *** RED ***
        {
            if (cell == null || !cell.HasFormula)                                   // *** RED ***
                return;                                                             // *** RED ***

            string formula = cell.Formula?.ToString() ?? "";                        // *** RED ***
            if (formula.IndexOf("'Compound Information'!", StringComparison.OrdinalIgnoreCase) < 0) // *** RED ***
                return;                                                             // *** RED ***

            string updatedFormula = Regex.Replace(                                  // *** RED ***
                formula,                                                            // *** RED ***
                @"('Compound Information'!\$?[A-Z]+\$?)(\d+)",                      // *** RED ***
                m =>                                                                 // *** RED ***
                {                                                                   // *** RED ***
                    if (!int.TryParse(m.Groups[2].Value, out int oldRow))           // *** RED ***
                        return m.Value;                                             // *** RED ***

                    // Only bump references that pointed to mainRowIndex - 1        // *** RED ***
                    if (oldRow == mainRowIndex - 1)                                 // *** RED ***
                        return m.Groups[1].Value + mainRowIndex.ToString();        // *** RED ***

                    return m.Value;                                                 // *** RED ***
                });                                                                 // *** RED ***

            if (!string.Equals(formula, updatedFormula, StringComparison.Ordinal))  // *** RED ***
                cell.Formula = updatedFormula;                                      // *** RED ***
        }



    }
}
