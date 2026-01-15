using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Spreadsheet.Handler.Objects;
using log4net.Core;
using Microsoft.Office.Interop.Excel;
using Internal.Framework.Properties;


namespace Spreadsheet.Handler
{
    public class DissolvedComponent
    {
        private static Application _app;

        private const int DEFAULT_NUM_TRANSFER_TIMES = 2;
        private const double DEFAULT_CHART_WIDTH = 468.75;

        private const string TEMP_DIRECTORY_NAME = "ABD_TempFiles";
        private const string TEMP_FILE_NAME = "DissolvedComponent.xls";

        private const string BASE_SERIES_NAMED_RANGE = "Series";
        private const string BATH_NAMED_RANGE = "Bath";
        private const string VESSEL_NAMED_RANGE = "Vessel";
        private const string TRANSFER_TIMES_NAMED_RANGE = "TransferTimes";
        //private const string TRANSFER_TIME_UNIT = " min";
        private const string STATISTICS_NAMED_RANGE = "Stats1";
        private const string BASE_SHEET_NAME = "DissolvedComponent";
        
        private static List<DissolvedComponentResult> DissolvedComponentResults { get; set; }
        private static List<String> DistinctTransferTimes { get; set; }
        private static Dictionary<string, List<DissolvedComponentResult>> DissolvedComponentResultSets { get; set; }


        public static string InsertDissolvedComponentSheet(string sourcePath, IPropertySetHost[] rows)
        {
            string returnPath = "";
            
            if (rows == null || rows.Length <=0)
            {
                Logger.LogMessage("An error occurred in the call to DissolvedComponent.InsertDissolvedComponentSheet. No rows passed to method.", Level.Error);
                return returnPath;
            }

            try
            {
                returnPath = InsertDissolvedComponentSheet2(sourcePath, rows);
            }
            catch (Exception ex)
            {
                Logger.LogMessage("An error occurred in the call to DissolvedComponent.InsertDissolvedComponentSheet. Message and stack trace are:\r\n" + ex.Message + "\r\n" + ex.StackTrace, Level.Error);

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
                            Logger.LogMessage("An error occurred in the call to DissolvedComponent.InsertDissolvedComponentSheet. Failed to save current workbook changes and to get path.", Level.Error);
                        }

                        _app.Workbooks.Close();
                    }
                    _app = null;
                }
                catch
                {
                    Logger.LogMessage("An error occurred in the call to DissolvedComponent.InsertDissolvedComponentSheet. Application failed to close workbooks. Message and stack trace are:\r\n" + ex.Message + "\r\n" + ex.StackTrace, Level.Error);
                }
                finally
                {
                    WorksheetUtilities.ReleaseExcelApp();
                }
            }
            return returnPath;
        }


        private static string InsertDissolvedComponentSheet2(string sourcePath, IEnumerable<IPropertySetHost> rows)
        {
            if (!File.Exists(sourcePath))
            {
                Logger.LogMessage("Error in call to DissolvedComponent.InsertDissolvedComponentSheet. Invalid source file path specified.", Level.Error);
                return "";
            }

            SetupComponentResults(rows);
            SetupComponentResultSets();

            if (DissolvedComponentResultSets.Count <= 0) return "";

            // Generate an random temp path to save new workbook
            string savePath = WorksheetUtilities.CopyWorkbook(sourcePath, TEMP_DIRECTORY_NAME, TEMP_FILE_NAME);
            if (String.IsNullOrEmpty(savePath)) return "";

            // Try to open the file
            _app = WorksheetUtilities.GetExcelApp();
            _app.Workbooks.Open(savePath, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            Workbook book = _app.Workbooks[1];
            Worksheet sheet = book.Worksheets[1] as Worksheet;

            if (sheet != null)
            {
                bool wasProtected = WorksheetUtilities.IsSheetPasswordProtected(sheet);
                //bool wasProtected = WorksheetUtilities.SetSheetProtection(sheet, null, false);

                // Add copies of the sheet
                if (DissolvedComponentResultSets.Count > 1)
                {
                    for (int i = 1; i < DissolvedComponentResultSets.Count; i++)
                    {
                        sheet.Copy(Type.Missing, book.Worksheets[i]);
                    }
                }

                // Now populate the sheets
                int j = 1;
                foreach (string key in DissolvedComponentResultSets.Keys)
                {
                    Worksheet resultSheet = book.Worksheets[j] as Worksheet;
                    if (resultSheet == null) continue;

                    // Name the sheet
                    resultSheet.Name = BASE_SHEET_NAME + j;

                    SetResultsIntoSheet(resultSheet, DissolvedComponentResultSets[key]);

                    try
                    {
                        _app.Goto(resultSheet.Cells[1, 1], true);
                    }
                    catch
                    {
                        Logger.LogMessage("Scroll of sheet failed in DissolvedComponent.UpdateDissolvedComponentSheet method!", Level.Error);
                    }

                   WorksheetUtilities.ReleaseComObject(resultSheet);
                    j++;

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

            return savePath;
        }

        /// <summary>
        /// Sets the results into the sheet
        /// </summary>
        /// <param name="sheet">The sheet into which to add the results.</param>
        /// <param name="results">The list of results</param>
        private static void SetResultsIntoSheet(_Worksheet sheet, List<DissolvedComponentResult> results)
        {
            // Get the vessel data together
            results.Sort(CompareByVessel);

            // Set the distinct transfer times
            SetDistinctTransferTimes(results);

            if (DistinctTransferTimes.Count > DEFAULT_NUM_TRANSFER_TIMES)
            {
                int diff = DistinctTransferTimes.Count - DEFAULT_NUM_TRANSFER_TIMES;

                // Expand the data ranges - adding columns to the main data range should expand the series
                WorksheetUtilities.InsertColumnsIntoNamedRange(diff, sheet, "MainDataTable", XlDirection.xlToRight);

                // Copy the statistics cells
                for (int i = 1; i <= diff; i++)
                {
                    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, STATISTICS_NAMED_RANGE, "Stats" + (DEFAULT_NUM_TRANSFER_TIMES + i), 1, DEFAULT_NUM_TRANSFER_TIMES + (i - 1), XlPasteType.xlPasteAll);
                }
            }

            // Set Peak Name
            WorksheetUtilities.SetNamedRangeValue(sheet, "PeakName", results[0].PeakName, 1, 1);

            // Set Resultset ID
            WorksheetUtilities.SetNamedRangeValue(sheet, "PeakName", "Resultset ID:", 2, 0);
            WorksheetUtilities.SetNamedRangeCellFontBold(sheet, "PeakName", 2, 0);
            WorksheetUtilities.SetNamedRangeValue(sheet, "PeakName", results[0].ResultSetId, 2, 1);    

            // Set transfer times
            WorksheetUtilities.SetNamedRangeValues2(sheet, TRANSFER_TIMES_NAMED_RANGE, DistinctTransferTimes);

            string currBathTransferTime = results[0].Bath + ":" + results[0].Vessel;
            int seriesNum = 1;
            List<DissolvedComponentResult> seriesResults = new List<DissolvedComponentResult>(0);
            foreach (DissolvedComponentResult result in results)
            {
                // Set the data into the named ranges - Bath, Vessel, SeriesN
                if (currBathTransferTime.Equals(result.Bath+ ":" + result.Vessel))
                {
                    seriesResults.Add(result);
                }
                else
                {
                    SetSeriesValues(sheet, seriesResults, seriesNum);

                    // Reset for the next set of series values
                    seriesResults = new List<DissolvedComponentResult> {result};

                    currBathTransferTime = result.Bath + ":" + result.Vessel;
                    seriesNum++;
                }
            }

            // Set the last set of series results
            seriesResults.Sort(CompareByResultId);
            SetSeriesValues(sheet, seriesResults, seriesNum);

            // Set the width of the chart to default - will autosize so keep consistent. There should be only ONE chart in the sheet.
            WorksheetUtilities.ResizeChart(sheet, 1, -1, DEFAULT_CHART_WIDTH, WorksheetUtilities.ResizeType.Width);
        }


        private static void SetSeriesValues(_Worksheet sheet, List<DissolvedComponentResult> seriesResults, int seriesNum)
        {
            List<String> vals = new List<string>(0);

            // Sort by result id since this should put the values in the correct order
            seriesResults.Sort(CompareByResultId);
            foreach (var seriesResult in seriesResults)
            {
                vals.Add(seriesResult.RoundedDissolvedAmount);
            }

            // Set the bath value
            WorksheetUtilities.SetNamedRangeValue(sheet, BATH_NAMED_RANGE, seriesResults[0].Bath, seriesNum, 1);

            // Set the vessel value
            WorksheetUtilities.SetNamedRangeValue(sheet, VESSEL_NAMED_RANGE, seriesResults[0].Vessel, seriesNum, 1);

            // Set the values into the current series
            WorksheetUtilities.SetNamedRangeValues2(sheet, BASE_SERIES_NAMED_RANGE + seriesNum, vals);
        }


        private static void SetDistinctTransferTimes(IEnumerable<DissolvedComponentResult> results)
        {
            List<double> distinctTransferTimes = new List<double>(0);

            foreach (DissolvedComponentResult result in results)
            {
                if (!distinctTransferTimes.Contains(result.TransferTime)) distinctTransferTimes.Add(result.TransferTime);
            }

            distinctTransferTimes.Sort();
            DistinctTransferTimes = new List<string>(0);
            foreach (int transferTime in distinctTransferTimes)
            {
                DistinctTransferTimes.Add(transferTime.ToString());
            }
        }


        private static void SetupComponentResultSets()
        {
            DissolvedComponentResultSets = new Dictionary<string, List<DissolvedComponentResult>>(0);
            if (DissolvedComponentResults.Count <= 0) return;

            // Sort the data by result set id first
            DissolvedComponentResults.Sort(CompareByResultSetIdPeakNameKey);

            // Get the first result set id/ peak name key - this separatess the result sets
            string currentResultSetIdPeakNameKey = DissolvedComponentResults[0].ResultSetIdPeakNameKey;

            List<DissolvedComponentResult> currentResults = new List<DissolvedComponentResult>(0);

            // Add the result sets to the dictionary
            foreach (DissolvedComponentResult component in DissolvedComponentResults)
            {
                if (!currentResultSetIdPeakNameKey.Equals(component.ResultSetIdPeakNameKey))
                {
                    DissolvedComponentResultSets.Add(currentResultSetIdPeakNameKey, currentResults);

                    currentResultSetIdPeakNameKey = component.ResultSetIdPeakNameKey;
                    currentResults = new List<DissolvedComponentResult>(0);
                }

                currentResults.Add(component);
            }

            DissolvedComponentResultSets.Add(currentResultSetIdPeakNameKey, currentResults);
        }


        private static void SetupComponentResults(IEnumerable<IPropertySetHost> rows)
        {
            DissolvedComponentResults = new List<DissolvedComponentResult>(0);

            foreach (IPropertySetHost row in rows)
            {
    
                try
                {
                    DissolvedComponentResult component = new DissolvedComponentResult
                    {
                        Bath = row.PropertySets["Dissolution_Results"]["Bath"].Value.ToString(),
                        PeakName = row.PropertySets["Dissolution_Results"]["Name"].Value.ToString(),
                        ResultId = row.PropertySets["Dissolution_Results"]["ResultId"].Value.ToString(),
                        ResultSetId = row.PropertySets["Dissolution_Results"]["ResultSetID"].Value.ToString(),
                        RoundedDissolvedAmount = row.PropertySets["Dissolution_Results"]["Dissolved_Percent"].Value.ToString(),
                        TransferTime = (double)row.PropertySets["Dissolution_Results"]["Transfer_Time"].Value,
                        Vessel = row.PropertySets["Dissolution_Results"]["Vessel"].Value.ToString(),

                    };
                    DissolvedComponentResults.Add(component);
                }
                catch
                {
                    continue;
                }
            }
        }


        private static int CompareByResultSetIdPeakNameKey(DissolvedComponentResult cr1, DissolvedComponentResult cr2)
        {
            if (cr1 == null || String.IsNullOrEmpty(cr1.ResultSetIdPeakNameKey))
            {
                if (cr2 == null || String.IsNullOrEmpty(cr2.ResultSetIdPeakNameKey))
                {
                    // Here, both are null so they are the same
                    return 0;
                }
                // Since cr2 is NOT null is greater. 
                return -1;
            }

            // In this case, if cr2 is null then cr1 is greater
            if (cr2 == null || String.IsNullOrEmpty(cr2.ResultSetIdPeakNameKey))
            {
                return 1;
            }

            // Now the actual values need to be compared
            string resultSetIdPeakNameKey1 = cr1.ResultSetIdPeakNameKey;
            string resultSetIdPeakNameKey2 = cr2.ResultSetIdPeakNameKey;

            return resultSetIdPeakNameKey1.CompareTo(resultSetIdPeakNameKey2);
        }


// ReSharper disable UnusedMember.Local
        private static int CompareByResultSetId(DissolvedComponentResult cr1, DissolvedComponentResult cr2)
// ReSharper restore UnusedMember.Local
        {
            if (cr1 == null || String.IsNullOrEmpty(cr1.ResultSetId))
            {
                if (cr2 == null || String.IsNullOrEmpty(cr2.ResultSetId))
                {
                    // Here, both are null so they are the same
                    return 0;
                }
                // Since cr2 is NOT null is greater. 
                return -1;
            }

            // In this case, if cr2 is null then cr1 is greater
            if (cr2 == null || String.IsNullOrEmpty(cr2.ResultSetId))
            {
                return 1;
            }

            // Now the actual values need to be compared
            long resultSetId1 = long.Parse(cr1.ResultSetId);
            long resultSetId2 = long.Parse(cr2.ResultSetId);

            if (resultSetId1 > resultSetId2)
            {
                return 1;
            }

            if (resultSetId1 == resultSetId2)
            {
                return 0;
            }

            // cr2 MUST be greater than cr1
            return -1;

        }


        private static int CompareByResultId(DissolvedComponentResult cr1, DissolvedComponentResult cr2)
        {
            if (cr1 == null || String.IsNullOrEmpty(cr1.ResultId))
            {
                if (cr2 == null || String.IsNullOrEmpty(cr2.ResultId))
                {
                    // Here, both are null so they are the same
                    return 0;
                }
                // Since cr2 is NOT null is greater. 
                return -1;
            }

            // In this case, if cr2 is null then cr1 is greater
            if (cr2 == null || String.IsNullOrEmpty(cr2.ResultId))
            {
                return 1;
            }

            // Now the actual values need to be compared
            long resultId1 = long.Parse(cr1.ResultId);
            long resultId2 = long.Parse(cr2.ResultId);

            if (resultId1 > resultId2)
            {
                return 1;
            }

            if (resultId1 == resultId2)
            {
                return 0;
            }

            // cr2 MUST be greater than cr1
            return -1;

        }

        private static int CompareByVessel(DissolvedComponentResult cr1, DissolvedComponentResult cr2)
        {
            if (cr1 == null || String.IsNullOrEmpty(cr1.Vessel))
            {
                if (cr2 == null || String.IsNullOrEmpty(cr2.Vessel))
                {
                    // Here, both are null so they are the same
                    return 0;
                }
                // Since cr2 is NOT null is greater. 
                return -1;
            }

            // In this case, if cr2 is null then cr1 is greater
            if (cr2 == null || String.IsNullOrEmpty(cr2.Vessel))
            {
                return 1;
            }

            // Now the actual values need to be compared
            long vessel1 = long.Parse(cr1.Vessel);
            long vessel2 = long.Parse(cr2.Vessel);

            if (vessel1 > vessel2)
            {
                return 1;
            }

            if (vessel1 == vessel2)
            {
                return 0;
            }

            // cr2 MUST be greater than cr1
            return -1;

        }

    }
}
