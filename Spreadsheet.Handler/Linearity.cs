using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using log4net.Core;
using Microsoft.Office.Interop.Excel;


namespace Spreadsheet.Handler
{
    public static class Linearity
    {
        private static Application _app;

        private const int DefaultDataPairCount = 7;
        private const int MinNumDataPairs = 2;

        private const string TempDirectoryName = "ABD_TempFiles";

        public static string UpdateLinearitySheet(
            string sourcePath,
            string strcmbProtocolType,
            string strcmbProductType,
            string strcmbTestType,
            int numReps,
            int numPeaks1,
            int numPeaks2,
            int numPeaks3,
            int numLevel1,
            int numLevel2,
            int numLevel3,
            string strcmbRRF,
            string strcmbR,
            string strcmbY,
            decimal rValue,
            decimal yInterceptValue
        )
        {
            string returnPath = "";
            try
            {
                returnPath = UpdateLinearitySheet2(
                                sourcePath,
                                strcmbProtocolType,
                                strcmbProductType,
                                strcmbTestType,
                                numReps,
                                numPeaks1,
                                numPeaks2,
                                numPeaks3,
                                numLevel1,
                                numLevel2,
                                numLevel3,
                                strcmbRRF,
                                strcmbR,
                                strcmbY,
                                rValue,
                                yInterceptValue
                            );
            }
            catch (Exception ex)
            {
                Logger.LogMessage("An error occurred in the call to Linearity.UpdateLinearitySheet. Message and stack trace are:\r\n" + ex.Message + "\r\n" + ex.StackTrace, Level.Error);

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
                            Logger.LogMessage("An error occurred in the call to Linearity.UpdateLinearitySheet. Failed to save current workbook changes and to get path.", Level.Error);
                        }

                        _app.Workbooks.Close();
                    }
                    _app = null;
                }
                catch
                {
                    Logger.LogMessage("An error occurred in the call to Linearity.UpdateLinearitySheet. Application failed to close workbooks. Message and stack trace are:\r\n" + ex.Message + "\r\n" + ex.StackTrace, Level.Error);
                }
                finally
                {
                    WorksheetUtilities.ReleaseExcelApp();
                }
            }

            return returnPath;
        }

        private static string UpdateLinearitySheet2(
            string sourcePath,
            string strcmbProtocolType,
            string strcmbProductType,
            string strcmbTestType,
            int numReps,
            int numPeaks1,
            int numPeaks2,
            int numPeaks3,
            int numLevel1,
            int numLevel2,
            int numLevel3,
            string strcmbRRF,
            string strcmbR,
            string strcmbY,
            decimal rValue,
            decimal yInterceptValue
        )
        {
            if (!File.Exists(sourcePath))
            {
                Logger.LogMessage("Error in call to Linearity.UpdateLinearitySheet. Invalid source file path specified.",Level.Error);
                return "";
            }

            if (numReps <= 0)
            {
                Logger.LogMessage("Error in call to Linearity.UpdateLinearitySheet. Concentration list is empty.", Level.Error);
                return "";
            }

            // Generate an random temp path to save new workbook
            string savePath = WorksheetUtilities.CopyWorkbook(sourcePath, TempDirectoryName, "Linearity Results.xls");
            if (String.IsNullOrEmpty(savePath)) return "";

            // Try to open the file
            _app = WorksheetUtilities.GetExcelApp();
            _app.Workbooks.Open(savePath, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            var peakLevelMap = new Dictionary<int, (int numPeaks, int numLevels)>
            {
                { 1, (numPeaks1, numLevel1) },
                { 2, (numPeaks2, numLevel2) },
                { 3, (numPeaks3, numLevel3) }
            };

            var nameMap = new Dictionary<int, List<string>>
            {
                { 1, "Alpha,Beta,Gamma".Split(',').ToList() },
                { 2, "Delta,Epsilon".Split(',').ToList() },
                { 3, "Zeta,Eta,Theta,Iota,Kappa,Lambda".Split(',').ToList() }
            };

            Workbook book = _app.Workbooks[1];
            Worksheet baseSheet = book.Worksheets[1] as Worksheet;

            if (baseSheet != null)
            {
                bool wasBaseSheetProtected = WorksheetUtilities.SetSheetProtection(baseSheet, null, false);

                WorksheetUtilities.SetMetadataValues(baseSheet, strcmbProtocolType, strcmbProductType, strcmbTestType);

                SetAcceptanceCriteriaValues(baseSheet, strcmbR, strcmbY, rValue, yInterceptValue);

                WorksheetUtilities.SetNamedRangeValue(baseSheet, "NumReps", numReps.ToString(), 1, 1);

                HandlePreps(baseSheet, numReps);

                foreach (var kvp in peakLevelMap)
                {
                    int setNumber = kvp.Key;
                    int numPeaks = kvp.Value.numPeaks;
                    int numLevels = kvp.Value.numLevels;

                    if (numPeaks == 0)
                    {
                        nameMap.Remove(setNumber);
                        continue;
                    }

                    var names = nameMap.ContainsKey(setNumber) ? nameMap[setNumber] : new List<string>();

                    for (int peakIndex = 0; peakIndex < numPeaks; peakIndex++)
                    {
                        baseSheet.Copy(After: book.Sheets[book.Sheets.Count]);
                        Worksheet copiedSheet = (Worksheet)book.Sheets[book.Sheets.Count];

                        bool wasProtected = WorksheetUtilities.SetSheetProtection(copiedSheet, null, false);

                        string peakName = (peakIndex < names.Count) ? names[peakIndex] : $"Peak {peakIndex + 1}";
                        copiedSheet.Name = $"Set {setNumber} {peakName}";

                        WorksheetUtilities.SetNamedRangeValue(copiedSheet, "MainPeakName", peakName, 1, 1);

                        UpdateWorksheet(copiedSheet, numReps, numLevels);

                        WorksheetUtilities.PostProcessSheet(copiedSheet);
                    }
                }

                baseSheet.Delete();

                WorksheetUtilities.ReleaseComObject(baseSheet);
            }

            book.Save();

            WorksheetUtilities.ReleaseComObject(book);
            _app.Workbooks.Close();

            _app = null;
            WorksheetUtilities.ReleaseExcelApp();

            // Return the path
            return savePath;
        }

        private static void SetAcceptanceCriteriaValues(Worksheet sheet, string strcmbR, string strcmbY, decimal rValue, decimal yInterceptValue)
        {
            WorksheetUtilities.SetNamedRangeValue(sheet, "AcceptanceCriteriaRange", strcmbR, 1, 1);
            WorksheetUtilities.SetNamedRangeValue(sheet, "AcceptanceCriteriaRange", strcmbY, 1, 2);
            WorksheetUtilities.SetNamedRangeValue(sheet, "AcceptanceCriteriaRange", rValue.ToString(), 2, 1);
            WorksheetUtilities.SetNamedRangeValue(sheet, "AcceptanceCriteriaRange", yInterceptValue.ToString(), 2, 2);
        }

        private static void HandlePreps(Worksheet sheet, int numReps)
        {
            if (numReps > 2)
            {
                WorksheetUtilities.InsertRowsIntoNamedRange(numReps - 2, sheet, "PrepsDetails", true, XlDirection.xlUp, XlPasteType.xlPasteAll);

                // set row numbers
                List<string> list = new List<string>(0);
                for (int i = 1; i <= numReps; i++)
                {
                    list.Add(i.ToString());
                }

                WorksheetUtilities.SetNamedRangeValues(sheet, "Preps", list);
            }
            else
            {
                WorksheetUtilities.DeleteRowsFromNamedRange(1, sheet, "PrepsDetails", XlDirection.xlDown);
            }
        }

        private static void UpdateWorksheet(Worksheet sheet, int numReps, int numLevels)
        {
            HandleLevels(sheet, numLevels);

            int numPoints = numReps * numLevels;

            if (numPoints > DefaultDataPairCount)
            {
                int numRowsToInsert = numPoints - DefaultDataPairCount;

                //For xyTable, a new format of the PeakName column was added, so two new methods had to be added (02/23)
                WorksheetUtilities.InsertRowsIntoNamedRangeFromRow(numRowsToInsert, sheet, "XY_Table", true, XlDirection.xlDown, XlPasteType.xlPasteAll, 3);
                WorksheetUtilities.RefreshFormulasforNamedRange(sheet, "Empower_Peak_Name", 2);

                //Changed the method for inserting rows into the named ranges as there is an issue when inserting rows where formulas breaks
                WorksheetUtilities.InsertRowsIntoNamedRangeFromRow(numRowsToInsert, sheet, "ResultsTable", true, XlDirection.xlDown, XlPasteType.xlPasteAll, 3);
                WorksheetUtilities.InsertRowsIntoNamedRangeFromRow(numRowsToInsert, sheet, "ValidationResultsTable", true, XlDirection.xlDown, XlPasteType.xlPasteAll, 3);
            }
            else if (numPoints < DefaultDataPairCount)
            {
                int numRowsToRemove = DefaultDataPairCount - numPoints;

                // There needs to be at least 2 rows in order to not corrupt the sheet's formulas
                if (DefaultDataPairCount - numRowsToRemove < MinNumDataPairs) numRowsToRemove = DefaultDataPairCount - MinNumDataPairs;

                // Delete the rows (from the bottom of the sheet tables up)
                WorksheetUtilities.DeleteRowsFromNamedRange(numRowsToRemove, sheet, "ValidationResultsTable", XlDirection.xlUp);
                WorksheetUtilities.DeleteRowsFromNamedRange(numRowsToRemove, sheet, "ResultsTable", XlDirection.xlUp);
                WorksheetUtilities.DeleteRowsFromNamedRange(numRowsToRemove, sheet, "XY_Table", XlDirection.xlUp);
            }

            // read value of Units named range. If not null or empty set the units of the charts below.

            //if (!String.IsNullOrEmpty(units))
            //{
            //    WorksheetUtilities.ReplaceInSheet(sheet, "mg/ml", units);
            //    WorksheetUtilities.UpdateChartCategoryAxisTitle(sheet, "Chart 3", "mg/mL", units);
            //    WorksheetUtilities.UpdateChartCategoryAxisTitle(sheet, "Chart 4", "mg/mL", units);
            //}
        }

        private static void HandleLevels(Worksheet sheet, int numLevels)
        {
            if (numLevels > 2)
            {
                WorksheetUtilities.InsertRowsIntoNamedRange(numLevels - 2, sheet, "Levels", true, XlDirection.xlUp, XlPasteType.xlPasteAll);

                // set row numbers
                List<string> list = new List<string>(0);
                for (int i = 1; i <= numLevels; i++)
                {
                    list.Add("Level " + i.ToString());
                }

                WorksheetUtilities.SetNamedRangeValues(sheet, "Levels", list);
            }
            else
            {
                WorksheetUtilities.DeleteRowsFromNamedRange(1, sheet, "Levels", XlDirection.xlDown);
            }
        }
    }// end of class
}
