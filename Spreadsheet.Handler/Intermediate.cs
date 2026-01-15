using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using log4net.Core;
using Microsoft.Office.Interop.Excel;


namespace Spreadsheet.Handler
{
    public class Intermediate
    {
        private static Application _app;

        private const int DefaultNumReps = 3;
        private const int DefaultNumBatches = 3;

        private const string TempDirectoryName = "ABD_TempFiles";

        public static string UpdateIntermediateSheet(string sourcePath, int numReps)
        {
            string returnPath = "";
            try
            {
                returnPath = UpdateIntermediateSheet2(sourcePath, numReps);
            }
            catch (Exception ex)
            {
                Logger.LogMessage("An error occurred in the call to Intermediate.UpdateIntermediateSheet. Message and stack trace are:\r\n" + ex.Message + "\r\n" + ex.StackTrace, Level.Error);

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
                            Logger.LogMessage("An error occurred in the call to Intermediate.UpdateIntermediateSheet. Failed to save current workbook changes and to get path.", Level.Error);
                        }

                        _app.Workbooks.Close();
                    }
                    _app = null;
                }
                catch
                {
                    Logger.LogMessage("An error occurred in the call to Intermediate.UpdateIntermediateSheet. Application failed to close workbooks. Message and stack trace are:\r\n" + ex.Message + "\r\n" + ex.StackTrace, Level.Error);
                }
                finally
                {
                    WorksheetUtilities.ReleaseExcelApp();
                }
            }
            return returnPath;
        }

        private static string UpdateIntermediateSheet2(string sourcePath, int numReps)
        {
            if (!File.Exists(sourcePath))
            {
                Logger.LogMessage("Error in call to Intermediate.UpdateIntermediateSheet. Invalid source file path specified.", Level.Error);
                return "";
            }

            // Generate an random temp path to save new workbook
            string savePath = WorksheetUtilities.CopyWorkbook(sourcePath, TempDirectoryName, "Intermediate Results.xls");
            if (String.IsNullOrEmpty(savePath)) return "";

            // Try to open the file
            _app = WorksheetUtilities.GetExcelApp();
            _app.Workbooks.Open(savePath, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            Workbook book = _app.Workbooks[1];
            Worksheet sheet = book.Worksheets[1] as Worksheet;

            if (sheet != null)
            {
                bool wasProtected = WorksheetUtilities.SetSheetProtection(sheet, null, false);

                if (numReps > DefaultNumReps)
                {
                    int numRowsToInsert = numReps - DefaultNumReps;
                    for (int i = 1; i <= DefaultNumBatches; i++)
                    {
                        WorksheetUtilities.InsertRowsIntoNamedRange(numRowsToInsert, sheet, "RunsBatch" + i, false, XlDirection.xlDown, XlPasteType.xlPasteFormulas);
                        WorksheetUtilities.InsertRowsIntoNamedRange(numRowsToInsert, sheet, "ValidationResultsBatch" + i, true, XlDirection.xlDown, XlPasteType.xlPasteFormulas);
                    }
                }
                else if (numReps < DefaultNumReps)
                {
                    // Only can delete ONE row otherwise the sheet will be corrupted!
                    for (int i = 1; i <= DefaultNumBatches; i++)
                    {
                        WorksheetUtilities.DeleteRowFromNamedRange(sheet, "RunsBatch" + i, 2);
                        WorksheetUtilities.DeleteRowFromNamedRange(sheet, "ValidationResultsBatch" + i, 2);
                    }
                }

                if (numReps > DefaultNumReps || numReps < DefaultNumReps)
                {
                    // Update the prep numberings in the sheet
                    List<string> prepNumbers = new List<string>(0);
                    for (int i = 1; i <= numReps; i++) prepNumbers.Add(i.ToString());
                    for (int i = 1; i <= DefaultNumBatches; i++)
                    {
                        WorksheetUtilities.SetNamedRangeValues(sheet, "PrepNumsBatch" + i, prepNumbers);
                        WorksheetUtilities.SetNamedRangeValues(sheet, "PrepNumsValBatch" + i, prepNumbers);
                    }
                }

                try
                {
                    _app.Goto(sheet.Cells[1, 1], true);
                }
                catch
                {
                    Logger.LogMessage("Scroll of sheet failed in Intermediate.UpdateIntermediateSheet!", Level.Error);
                }

                if (wasProtected) WorksheetUtilities.SetSheetProtection(sheet, null, true);

                WorksheetUtilities.ReleaseComObject(sheet);
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
    }
}
