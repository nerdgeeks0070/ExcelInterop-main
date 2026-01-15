using log4net.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.IO;

namespace Spreadsheet.Handler
{
    public static class AccuracyNew
    {
        private static Application _app;

        private const string TempDirectoryName = "ABD_TempFiles";

        public static string UpdateAccuracySheet(
            string sourcePath,
            // --- General ---
            string strcmbProtocolType, string strcmbProductType, string strcmbTestType)
        {
            string returnPath = "";

            try
            {
                returnPath = UpdateAccuracySheet2(
                    sourcePath,
                    // --- General ---
                    strcmbProtocolType, strcmbProductType, strcmbTestType
                );
            }
            catch (Exception ex)
            {
                Logger.LogMessage("An error occurred in the call to AccuracyNew.UpdateAccuracySheet." +
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

        private static string UpdateAccuracySheet2(
            string sourcePath,
            // --- General ---
            string strcmbProtocolType, string strcmbProductType, string strcmbTestType
        )
        {
            if (!File.Exists(sourcePath))
            {
                Logger.LogMessage("Error in call to AccuracyNew.UpdateAccuracySheet2. Invalid source file path specified.", Level.Error);
                return "";
            }

            string savePath = WorksheetUtilities.CopyWorkbook(sourcePath, TempDirectoryName, "Accuracy Results.xls");
            if (string.IsNullOrEmpty(savePath)) return "";

            _app = WorksheetUtilities.GetExcelApp();
            _app.Workbooks.Open(savePath, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            Workbook book = _app.Workbooks[1];

            //Worksheet sheetAssay = book.Worksheets["Assay"] as Worksheet;
            //Worksheet sheetWater = book.Worksheets["Water Content"] as Worksheet;

            // add code here.

            book.Save();
            WorksheetUtilities.ReleaseComObject(book);
            _app.Workbooks.Close();
            _app = null;
            WorksheetUtilities.ReleaseExcelApp();

            return savePath;
        }
    }
}
