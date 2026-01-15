using Spreadsheet.Handler;
using log4net.Core;
using Microsoft.Office.Interop.Excel;

using System;
using System.Collections.Generic;
using System.IO;

namespace Spreadsheet.Handler
{
    public static class SystemSuitability
    {
        private static Application _app;

        private const string TempDirectoryName = "ABD_TempFiles";

        /// <summary>
        /// Method to be called via Scripts - GenerateExperiment
        /// </summary>
        /// <param name="sourcePath"></param>
        public static string UpdateSystemSuitabilitySheet(
            string sourcePath,
            string strcmbProtocolType,
            string strcmbProductType,
            string strcmbTestType,
            int numBlankInterference,
            int numSensitivity,
            int numRSD,
            int numStandardAgreement,
            int numTailingFactor,
            int numResolutionTest,
            int numTheoreticalPlates,
            int numPeakToValleyRatio,
            int numRetentionFactor,
            int numDetectability,
            int numStdRecovery,
            int numAvgTiterVal,
            int numOther)
        {
            string returnPath = "";
            try
            {
                returnPath = UpdateSystemSuitabilitySheet2(
                                sourcePath,
                                strcmbProtocolType,
                                strcmbProductType,
                                strcmbTestType,
                                numBlankInterference,
                                numSensitivity,
                                numRSD,
                                numStandardAgreement,
                                numTailingFactor,
                                numResolutionTest,
                                numTheoreticalPlates,
                                numPeakToValleyRatio,
                                numRetentionFactor,
                                numDetectability,
                                numStdRecovery,
                                numAvgTiterVal,
                                numOther
                            );
            }
            catch (Exception ex)
            {
                Logger.LogMessage("An error occurred in the call to SystemSuitability.SystemSuitability. Message and stack trace are:\r\n" + ex.Message + "\r\n" + ex.StackTrace, Level.Error);

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
                            Logger.LogMessage("An error occurred in the call to SystemSuitability.UpdateSystemSuitabilitySheet. Failed to save current workbook changes and to get path.", Level.Error);
                        }

                        _app.Workbooks.Close();
                    }
                    _app = null;
                }
                catch
                {
                    Logger.LogMessage("An error occurred in the call to SystemSuitability.UpdateSystemSuitabilitySheet. Application failed to close workbooks. Message and stack trace are:\r\n" + ex.Message + "\r\n" + ex.StackTrace, Level.Error);
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
        /// Method with the logic for calling / updating Excel spreadsheet for SystemSuitability
        /// </summary>
        /// <param name="sourcePath"></param>
        /// <param name="criteriaDict"></param>
        /// <param name=""></param>
        /// <returns></returns>
        private static string UpdateSystemSuitabilitySheet2(
            string sourcePath,
            string strcmbProtocolType,
            string strcmbProductType,
            string strcmbTestType,
            int numBlankInterference,
            int numSensitivity,
            int numRSD,
            int numStandardAgreement,
            int numTailingFactor,
            int numResolutionTest,
            int numTheoreticalPlates,
            int numPeakToValleyRatio,
            int numRetentionFactor,
            int numDetectability,
            int numStdRecovery,
            int numAvgTiterVal,
            int numOther)
        {
            if (!File.Exists(sourcePath))
            {
                Logger.LogMessage("Error in call to SystemSuitability.UpdateSystemSuitabilitySheet. Invalid source file path specified.", Level.Error);
                return "";
            }

            // Generate an random temp path to save new workbook
            string savePath = WorksheetUtilities.CopyWorkbook(sourcePath, TempDirectoryName, "System Suitability Results.xls");
            if (String.IsNullOrEmpty(savePath)) return "";

            // Try to open the file
            _app = WorksheetUtilities.GetExcelApp();
            _app.Workbooks.Open(savePath, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            Workbook book = _app.Workbooks[1];
            Worksheet sheetToDelete = book.Worksheets[strcmbTestType == "Water Content" ? "All" : "Water Content"] as Worksheet;
            WorksheetUtilities.DeleteSheet(sheetToDelete);

            Worksheet sheet = book.Worksheets[1] as Worksheet;
            if (sheet != null)
            {
                bool wasProtected = WorksheetUtilities.SetSheetProtection(sheet, null, false);

                WorksheetUtilities.SetMetadataValues(sheet, strcmbProtocolType, strcmbProductType, strcmbTestType);

                Dictionary<string, int> fieldCounts = new Dictionary<string, int>
                {
                    { "Interference", numBlankInterference },
                    { "RSD", numRSD },
                    { "StandardAgreement", numStandardAgreement },
                    { "TailingSymmetryFactor", numTailingFactor },
                    { "Resolution", numResolutionTest },
                    { "SNRatio", numSensitivity },
                    { "NumberTheoreticalPlates", numTheoreticalPlates },
                    { "PeakToValleyRatio", numPeakToValleyRatio },
                    { "Detectability", numDetectability },
                    { "RetentionCapacityFactor", numRetentionFactor },
                    { "StandardRecovery_KF", numStdRecovery },
                    { "AverageTiterValue_KF", numAvgTiterVal },
                    { "Other", numOther }
                };

                Dictionary<string, int> itemCounts;

                if (strcmbTestType == "Water Content")
                {
                    itemCounts = new Dictionary<string, int>
                    {
                        { "AverageTiterValue",        fieldCounts["AverageTiterValue_KF"] },
                        { "RSD",                      fieldCounts["RSD"] },
                        { "StandardRecovery",         fieldCounts["StandardRecovery_KF"] },
                        { "Other",                    fieldCounts["Other"] }
                    };
                }
                else
                {
                    bool isVolatiles = strcmbTestType == "Volatiles";
                    bool isDissolution = strcmbTestType == "Dissolution";
                    bool isOther = !isVolatiles && !isDissolution;

                    itemCounts = new Dictionary<string, int>
                    {
                        { "Interference_Dissolution", isDissolution ? fieldCounts["Interference"] : 0 },
                        { "Interference_Other",       isOther       ? fieldCounts["Interference"] : 0 },
                        { "Interference_Volatiles",   isVolatiles   ? fieldCounts["Interference"] : 0 },
                        { "RSD_Volatiles",            isVolatiles   ? fieldCounts["RSD"] : 0 },
                        { "RSD_Other",                !isVolatiles  ? fieldCounts["RSD"] : 0 },
                        { "StandardAgreement",        fieldCounts["StandardAgreement"] },
                        { "TailingSymmetryFactor",    fieldCounts["TailingSymmetryFactor"] },
                        { "Resolution",               fieldCounts["Resolution"] },
                        { "SNRatio_Other",            !isVolatiles  ? fieldCounts["SNRatio"] : 0 },
                        { "SNRatio_Volatiles",        isVolatiles   ? fieldCounts["SNRatio"] : 0 },
                        { "NumberTheoreticalPlates",  fieldCounts["NumberTheoreticalPlates"] },
                        { "PeakToValleyRatio",        fieldCounts["PeakToValleyRatio"] },
                        { "Detectability",            fieldCounts["Detectability"] },
                        { "RetentionCapacityFactor",  fieldCounts["RetentionCapacityFactor"] },
                        { "Other",                    fieldCounts["Other"] }
                    };
                }

                ProcessRows(sheet, itemCounts);

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

        private static void ProcessRows(Worksheet sheet, Dictionary<string, int> itemCounts)
        {
            int baseRow = 5;    // First data row
            int colStart = 2;   // Column B (index 2)
            int colEnd = 15;    // Column O (index 15)
            int currRow = baseRow;

            foreach (var kvp in itemCounts)
            {
                string baseRangeName = kvp.Key;
                int copies = kvp.Value;
                Range baseRange = sheet.Range[baseRangeName];

                if (copies <= 0)
                {
                    baseRange.EntireRow.Delete();
                }
                else
                {
                    for (int copyIndex = 1; copyIndex < copies; copyIndex++)
                    {
                        Range insertAt = sheet.Rows[baseRange.Row + copyIndex];
                        baseRange.EntireRow.Copy();
                        insertAt.Insert(XlInsertShiftDirection.xlShiftDown);

                        string newRangeName = $"{baseRangeName}{copyIndex + 1}";
                        Range newRange = sheet.Range[
                            sheet.Cells[baseRange.Row + copyIndex, colStart],
                            sheet.Cells[baseRange.Row + copyIndex, colEnd]
                        ];

                        try { sheet.Names.Item(newRangeName).Delete(); } catch { }

                        sheet.Names.Add(Name: newRangeName, RefersTo: newRange);
                    }

                    currRow += copies;
                }
            }
        }
    }//End of Class
}
