using log4net.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;


namespace Spreadsheet.Handler
{
    public class AssayLevel
    {
        public int NoSArea { get; set; }
        public int NoSAssay { get; set; }
        public int NoSLC { get; set; }

        public decimal DecRSD1 { get; set; }
        public string StrOperetor { get; set; }
    }

    public class ImpurityLevel
    {
        /// <summary>
        /// Previously referred to as "Number of Impurities".
        /// </summary>
        public int NumPeaks { get; set; }

        /// <summary>
        /// Previously referred to as "Number of Samples".
        /// </summary>
        public int NumSolutions { get; set; }

        public string StrQuantitationType { get; set; }

        public string RangeOperator1 { get; set; }
        public string RangeOperator2 { get; set; }
        public decimal RangeValue1 { get; set; }
        public decimal RangeValue2 { get; set; }

        public string DiffAbsRel { get; set; }
        public decimal DiffValue1 { get; set; }
        public decimal DiffValue2 { get; set; }
        public decimal DiffValue3 { get; set; }
    }

    public class SNRatio
    {
        public int SNNumberOfPeaks { get; set; }
        public int SNNumberOfSolutions { get; set; }

        public string StrSNOperator { get; set; }

        public decimal DecSNValue { get; set; }
    }

    public class PeaksInAgedSolution
    {
        public int PASNumberOfPeaks { get; set; }
        public int PASNumberOfResPair { get; set; }

        public string StrResolutionOperator { get; set; }

        public decimal DecResolutionValue { get; set; }
    }

    public class SolutionStability
    {
        private static Application _app;

        private const int DefaultNumConditions = 2;
        private const int DefaultImpurityCols = 4;
        private const int DefaultSNResolutionCols = 3;
        private const int DefaultSNResolutionSummaryCols = 1;

        private const string TempDirectoryName = "ABD_TempFiles";

        public static string UpdateSolutionStabilitySheet(
            string sourcePath,
            // General
            string strcmbProtocolType,
            string strcmbProductType,
            string strcmbTestType,
            int numStorageConditions,
            int numDataPoints,
            int numNoSolsArea,
            int numNoSolsAssay,
            int numNoSolsLC,
            decimal assayLevelValue,
            string assayLevelOperator,
            int numPeaks,
            int numNoSols,
            string strCmbNoSols,
            string impurityRangeOperator1,
            decimal impurityRangeValue1,
            decimal impurityRangeValue2,
            string impurityRangeOperator2,
            string impurityDiffAbsRel,
            decimal impurityDiffValue1,
            decimal impurityDiffValue2,
            decimal impurityDiffValue3,
            int numSNNoPeaks,
            int numSNNoSols,
            string snOperator,
            decimal snValue,
            int numNRNoPeaks,
            int numNoResPair,
            string resolutionOperator,
            decimal resolutionValue
        )
        {
            string returnPath = "";
            try
            {
                returnPath = UpdateSolutionStabilitySheet2(
                    sourcePath,
                    strcmbProtocolType,
                    strcmbProductType,
                    strcmbTestType,
                    numStorageConditions,
                    numDataPoints,
                    numNoSolsArea,
                    numNoSolsAssay,
                    numNoSolsLC,
                    assayLevelValue,
                    assayLevelOperator,
                    numPeaks,
                    numNoSols,
                    strCmbNoSols,
                    impurityRangeOperator1,
                    impurityRangeValue1,
                    impurityRangeValue2,
                    impurityRangeOperator2,
                    impurityDiffAbsRel,
                    impurityDiffValue1,
                    impurityDiffValue2,
                    impurityDiffValue3,
                    numSNNoPeaks,
                    numSNNoSols,
                    snOperator,
                    snValue,
                    numNRNoPeaks,
                    numNoResPair,
                    resolutionOperator,
                    resolutionValue);
            }
            catch (Exception ex)
            {
                Logger.LogMessage("An error occurred in the call to SolutionStability.UpdateSolutionStabilitySheet. Message and stack trace are:\r\n" + ex.Message + "\r\n" + ex.StackTrace, Level.Error);

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
                            Logger.LogMessage("An error occurred in the call to SolutionStability.UpdateSolutionStabilitySheet. Failed to save current workbook changes and to get path.", Level.Error);
                        }

                        _app.Workbooks.Close();
                    }
                    _app = null;
                }
                catch
                {
                    Logger.LogMessage("An error occurred in the call to SolutionStability.UpdateSolutionStabilitySheet. Application failed to close workbooks. Message and stack trace are:\r\n" + ex.Message + "\r\n" + ex.StackTrace, Level.Error);
                }
                finally
                {
                    WorksheetUtilities.ReleaseExcelApp();
                }
            }
            return returnPath;
        }

        private static string UpdateSolutionStabilitySheet2(
            string sourcePath,
            // General
            string strcmbProtocolType,
            string strcmbProductType,
            string strcmbTestType,
            int numStorageConditions,
            int numDataPoints,
            int numNoSolsArea,
            int numNoSolsAssay,
            int numNoSolsLC,
            decimal assayLevelValue,
            string assayLevelOperator,
            int numPeaks,
            int numNoSols,
            string strCmbNoSols,
            string impurityRangeOperator1,
            decimal impurityRangeValue1,
            decimal impurityRangeValue2,
            string impurityRangeOperator2,
            string impurityDiffAbsRel,
            decimal impurityDiffValue1,
            decimal impurityDiffValue2,
            decimal impurityDiffValue3,
            int numSNNoPeaks,
            int numSNNoSols,
            string snOperator,
            decimal snValue,
            int numNRNoPeaks,
            int numNoResPair,
            string resolutionOperator,
            decimal resolutionValue
        )
        {
            if (numStorageConditions <= 0 || numDataPoints <= 0)
            {
                Logger.LogMessage("Error in call to SolutionStability.UpdateSolutionStabilitySheet. Invalid Storage Conditions or Data Points!", Level.Error);
                return "";
            }

            if (!File.Exists(sourcePath))
            {
                Logger.LogMessage("Error in call to SolutionStability.UpdateSolutionStabilitySheet. Invalid source file path specified.", Level.Error);
                return "";
            }

            // Generate an random temp path to save new workbook
            string savePath = WorksheetUtilities.CopyWorkbook(sourcePath, TempDirectoryName, "Solution Stability Results.xls");
            if (String.IsNullOrEmpty(savePath)) return "";

            // Try to open the file
            _app = WorksheetUtilities.GetExcelApp();
            _app.Workbooks.Open(savePath, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            //_app.Visible = true;

            Workbook book = _app.Workbooks[1];
            Worksheet sheet = book.Worksheets[1] as Worksheet;

            AssayLevel assayLevel = new AssayLevel
            {
                NoSArea = numNoSolsArea,
                NoSAssay = numNoSolsAssay,
                NoSLC = numNoSolsLC,
                StrOperetor = assayLevelOperator,
                DecRSD1 = assayLevelValue
            };

            ImpurityLevel impurityLevel = new ImpurityLevel
            {
                NumPeaks = numPeaks,
                NumSolutions = numNoSols,
                StrQuantitationType = strCmbNoSols,
                RangeOperator1 = impurityRangeOperator1,
                RangeOperator2 = impurityRangeOperator2,
                RangeValue1 = impurityRangeValue1,
                RangeValue2 = impurityRangeValue2,
                DiffAbsRel = impurityDiffAbsRel,
                DiffValue1 = impurityDiffValue1,
                DiffValue2 = impurityDiffValue2,
                DiffValue3 = impurityDiffValue3
            };

            SNRatio snRatio = new SNRatio
            {
                SNNumberOfPeaks = numSNNoPeaks,
                SNNumberOfSolutions = numSNNoSols,
                StrSNOperator = snOperator,
                DecSNValue = snValue
            };

            PeaksInAgedSolution peaksInAgedSolution = new PeaksInAgedSolution
            {
                PASNumberOfPeaks = numNRNoPeaks,
                PASNumberOfResPair = numNoResPair,
                StrResolutionOperator = resolutionOperator,
                DecResolutionValue = resolutionValue
            };

            WorksheetUtilities.SetMetadataValues(sheet, strcmbProtocolType, strcmbProductType, strcmbTestType);

            UpdateWorksheet(sheet,
                strcmbTestType,
                numStorageConditions,
                numDataPoints,
                assayLevel,
                impurityLevel,
                snRatio,
                peaksInAgedSolution);

            _app.Workbooks[1].Save();

            WorksheetUtilities.ReleaseComObject(book);
            _app.Workbooks.Close();

            //WorksheetUtilities.ReleaseComObject(_app);
            _app = null;
            WorksheetUtilities.ReleaseExcelApp();

            // Return the path
            return savePath;
        }

        private static void UpdateWorksheet(Worksheet sheet, string testType, int numStorageConditions, int numDataPoints, AssayLevel assayLevel, ImpurityLevel impurityLevel, SNRatio snRatio, PeaksInAgedSolution peaksInAgedSolution)
        {
            if (sheet == null)
            {
                return;
            }

            bool wasProtected = WorksheetUtilities.SetSheetProtection(sheet, null, false);

            

            // Expand or contract the data tables based on the number of conditions
            string[] namedConditonTable = new string[]{
                    "ImpurityConditionsTable1",
                    "ResultsConditionsTable1" ,
                    "AreaConditionsTable1",
                    "AssayConditionsTable1",
                    "ClaimConditionsTable1",
                    "ImpuritySummaryConditionsTable1"
                };

            int totalStorageConditions = numStorageConditions * numDataPoints;

            ProcessMainLevel(sheet, totalStorageConditions, namedConditonTable);
            
            ProcessImpurityLevel(sheet, impurityLevel);

            HandleSamples(sheet, assayLevel, totalStorageConditions);

            HandleImpuritySamples(sheet, totalStorageConditions, impurityLevel.NumPeaks, impurityLevel.NumSolutions);

            WorksheetUtilities.DeleteInvalidNamedRanges(sheet);

            HandleLargeNumberOfImpurities(sheet, impurityLevel.NumPeaks);

            ProcessSNRatio(sheet, totalStorageConditions, snRatio);
            ProcessSNRatioSummary(sheet, totalStorageConditions, snRatio);

            ProcessResolution(sheet, totalStorageConditions, peaksInAgedSolution);

            ProcessResolutionSummary(sheet, totalStorageConditions, peaksInAgedSolution);

            if (testType != "Dissolution")
            {
                DeleteVerticalRangeShiftLeft(sheet);
            }

            WorksheetUtilities.PostProcessSheet(sheet);
        }
        
        private static void ProcessMainLevel(Worksheet sheet, int conditionsCount, string[] namedConditonTable)
        {
            if (conditionsCount > DefaultNumConditions)
            {
                int numRowsToInsert = conditionsCount - DefaultNumConditions;
                WorksheetUtilities.InsertRowsIntoNamedRange(numRowsToInsert, sheet, "ImpurityConditionsTable1", true, XlDirection.xlDown, XlPasteType.xlPasteAll);
                UpdateDifferenceFormulas(sheet, "ImpurityConditionsTable1", 1, 3, 0);
                UpdateResultFormulas(sheet, "ImpuritySampleStandards1", 3, 5, 0, 2, 2);
                WorksheetUtilities.InsertRowsIntoNamedRange(numRowsToInsert, sheet, "ImpuritySummaryConditionsTable1", true, XlDirection.xlDown, XlPasteType.xlPasteAll);
                UpdateDifferenceFormulas(sheet, "ImpuritySummaryConditionsTable1", 1, 3, 0);
                WorksheetUtilities.InsertRowsIntoNamedRange(numRowsToInsert, sheet, "AreaConditionsTable1", true, XlDirection.xlDown, XlPasteType.xlPasteAll);
                UpdateDifferenceFormulas(sheet, "AreaConditionsTable1", 1, 5, 0);
                WorksheetUtilities.InsertRowsIntoNamedRange(numRowsToInsert, sheet, "AssayConditionsTable1", true, XlDirection.xlDown, XlPasteType.xlPasteAll);
                UpdateDifferenceFormulas(sheet, "AssayConditionsTable1", 1, 5, 0);
                WorksheetUtilities.InsertRowsIntoNamedRange(numRowsToInsert, sheet, "ClaimConditionsTable1", true, XlDirection.xlDown, XlPasteType.xlPasteAll);
                UpdateDifferenceFormulas(sheet, "ClaimConditionWithInitialTable", 1, 4, 5);
                UpdateDifferenceFormulas(sheet, "ClaimConditionsTable1", 1, 5, 0);
                WorksheetUtilities.InsertRowsIntoNamedRange(numRowsToInsert, sheet, "ResultsConditionsTable1", true, XlDirection.xlDown, XlPasteType.xlPasteAll);

            }
            else if (conditionsCount < DefaultNumConditions)
            {
                int numRowsToRemove = DefaultNumConditions - conditionsCount;
                // There needs to be at least 1 row in order to not corrupt the sheet's formulas
                if (DefaultNumConditions - numRowsToRemove < 1) numRowsToRemove = DefaultNumConditions - 1;
                foreach (string condTable in namedConditonTable)
                {
                    //Needed to remove last row from Impurity Condition table, as formula gets corrupted if only one Condition is present.
                    if (condTable == "ImpurityConditionsTable1" || condTable == "ImpuritySummaryConditionsTable1")
                    {
                        WorksheetUtilities.DeleteRowsFromNamedRange(numRowsToRemove, sheet, condTable, XlDirection.xlUp);
                    }
                    else
                    {
                        WorksheetUtilities.DeleteRowsFromNamedRange(numRowsToRemove, sheet, condTable, XlDirection.xlDown);
                    }
                }
            }
        }

        private static void ProcessImpurityLevel(Worksheet sheet, ImpurityLevel impurityLevel)
        {
            int numImpurities = impurityLevel.NumPeaks;

            WorksheetUtilities.SetNamedRangeValue(sheet, "ImpurityRangeOperator1", impurityLevel.RangeOperator1, 1, 1);
            WorksheetUtilities.SetNamedRangeValue(sheet, "ImpurityRangeOperator2", impurityLevel.RangeOperator2, 1, 1);
            WorksheetUtilities.SetNamedRangeValue(sheet, "ImpurityRangeValue1", impurityLevel.RangeValue1.ToString(), 1, 1);
            WorksheetUtilities.SetNamedRangeValue(sheet, "ImpurityRangeValue2", impurityLevel.RangeValue2.ToString(), 1, 1);
            WorksheetUtilities.SetNamedRangeValue(sheet, "ImpurityDiffValue1", impurityLevel.DiffValue1.ToString(), 1, 1);
            WorksheetUtilities.SetNamedRangeValue(sheet, "ImpurityDiffValue2", impurityLevel.DiffValue2.ToString(), 1, 1);
            WorksheetUtilities.SetNamedRangeValue(sheet, "ImpurityDiffValue3", impurityLevel.DiffValue3.ToString(), 1, 1);
            WorksheetUtilities.SetNamedRangeValue(sheet, "ImpurityDiffRelAbs", impurityLevel.DiffAbsRel, 1, 1);
            WorksheetUtilities.SetNamedRangeValue(sheet, "ImpurityQuantitationType", impurityLevel.StrQuantitationType, 1, 1);

            //--------------- Add Impurity series ----------------------------------
            // As of 08 - 07, it was requested that if the number of impurities is higher than 6, replicate vertically
            //And start adding new columns from there
            for (int i = 1; i <= numImpurities; i++)
            {
                // Named ranges for current and previous impurities
                string impurityRangName = "Impurity" + i;
                string previouseImpurityRangName = "Impurity" + (i - 1);
                string summaryImpRangeName = "ImpuritySummaryColumn" + i;
                string previousSummaryImpRangeName = "ImpuritySummaryColumn" + (i - 1);
                string prevImpName = "ImpurityName" + (i - 1);
                string impName = "ImpurityName" + i;
                string prevImpSummaryName = "ImpuritySummaryName" + (i - 1);
                string impSummaryName = "ImpuritySummaryName" + i;

                // Linking ranges for summary table
                string differenceRangeName = "ImpurityDifferenceConditions" + i + "1";
                string previousDifferenceRangeName = "ImpurityDifferenceConditions" + (i - 1) + "1";
                string summaryConditionsRangeName = "ImpuritySummaryConditions" + i + "1";
                string previousSummaryConditionsRangeName = "ImpuritySummaryConditions" + (i - 1) + "1";

                // Difference percentage ranges
                string previousDifferencePercentage = "ImpurityDifference" + (i - 1) + "1";
                string differencePercentage = "ImpurityDifference" + i + "1";
                string previousSummaryPercentage = "ImpuritySummaryDifference" + (i - 1) + "1";
                string summaryDifferencePercentage = "ImpuritySummaryDifference" + i + "1";

                if (!WorksheetUtilities.NamedRangeExist(sheet, impurityRangName))
                {
                    // Add Impurity Validation Report NamedRange to the right
                    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, previouseImpurityRangName, impurityRangName, 1, DefaultImpurityCols + 1, XlPasteType.xlPasteAll);
                    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, prevImpName, impName, 1, DefaultImpurityCols + 1, XlPasteType.xlPasteAll);
                    WorksheetUtilities.ResizeNamedRange(sheet, "ImpuritySampleStandards1", 0, DefaultImpurityCols + 1);

                    // Add to Impurity Summary Table
                    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, previousSummaryImpRangeName, summaryImpRangeName, 1, 2, XlPasteType.xlPasteAll);
                    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, prevImpSummaryName, impSummaryName, 1, 2, XlPasteType.xlPasteAll);
                    WorksheetUtilities.ResizeNamedRange(sheet, "ImpuritySummarySample1", 0, 1);
                    WorksheetUtilities.ResizeNamedRange(sheet, "TableForImpurity", 0, 1);

                    // Set impurity labels
                    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, prevImpName, impName, 1, DefaultImpurityCols + 1, XlPasteType.xlPasteAll);
                    WorksheetUtilities.SetNamedRangeValue(sheet, impName, "Impurity " + i, 1, 1);
                    WorksheetUtilities.SetNamedRangeValue(sheet, impSummaryName, "Impurity" + i, 1, 1);

                    // Link new impurities to summary table
                    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, previousDifferenceRangeName, differenceRangeName, 1, DefaultImpurityCols + 1, XlPasteType.xlPasteAll);
                    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, previousSummaryConditionsRangeName, summaryConditionsRangeName, 1, 2, XlPasteType.xlPasteAll);
                    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, previousDifferencePercentage, differencePercentage, 1, DefaultImpurityCols + 1, XlPasteType.xlPasteAll);
                    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, previousSummaryPercentage, summaryDifferencePercentage, 1, 2, XlPasteType.xlPasteAll);

                    // Link cells using wrapper method
                    WorksheetUtilities.LinkTwoNamedRangeCellsWrapper(sheet, differenceRangeName, summaryConditionsRangeName, 1 /* sourceCol */, 1 /* targetCol */, true /* roundForImpurity */);
                    WorksheetUtilities.LinkTwoNamedRangeCellsWrapper(sheet, differencePercentage, summaryDifferencePercentage, 1 /* sourceCol */, 1 /* targetCol */, true /* roundForImpurity */);
                    WorksheetUtilities.LinkTwoNamedRangeCellsWrapper(sheet, impName, impSummaryName, 1 /* sourceCol */, 1 /* targetCol */);
                }
            }
        }

        private static void HandleSamples(Worksheet sheet, AssayLevel assayLevel, int conditionsCount)
        {
            WorksheetUtilities.SetNamedRangeValue(sheet, "AssayLevelOperator", assayLevel.StrOperetor, 1, 1);
            WorksheetUtilities.SetNamedRangeValue(sheet, "AssayLevelValue", assayLevel.DecRSD1.ToString(), 1, 1);

            List<string> samplesStandards = new List<string>();
            samplesStandards.AddRange(GenerateSampleNames("Area", assayLevel.NoSArea, "Area"));
            samplesStandards.AddRange(GenerateSampleNames("Assay", assayLevel.NoSAssay, "Assay"));
            samplesStandards.AddRange(GenerateSampleNames("LabelClaim", assayLevel.NoSLC, "Label Claim"));

            // Handle the number of samples/standards
            int sampleCount = 0;
            int areaCount = 0;
            int assayCount = 0;
            int claimCount = 0;
            foreach (string sample in samplesStandards)
            {
                sampleCount++;
                string label = "";
                string type = "";
                string header = "";
                string[] sampleStandardsArray = sample.Split(',');
                label = sampleStandardsArray[0].Trim();
                if (sampleStandardsArray.Length >= 2)
                {
                    type = sampleStandardsArray[1].Trim();
                }

                // handle Results table

                if (!WorksheetUtilities.NamedRangeExist(sheet, "ResultsSampleStandards" + sampleCount))
                {
                    // add new named table
                    // Insert a range between the ranges numbered 1 & 2, copy from range 1 to new range (copy all)
                    WorksheetUtilities.InsertRowsIntoNamedRange(conditionsCount + 2, sheet, "ResultsData", false, XlDirection.xlUp, XlPasteType.xlPasteAll);
                    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ResultsSampleStandards" + (sampleCount - 1), "ResultsSampleStandards" + (sampleCount), conditionsCount + 3, 1, XlPasteType.xlPasteAll);
                    //Need a new NamedRange to handle the linking
                    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ResultSamplesLinking" + (sampleCount - 1), "ResultSamplesLinking" + sampleCount, conditionsCount + 3, 1, XlPasteType.xlPasteAll);
                    //New val Report Table
                    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "TableForAssay" + (sampleCount - 1), "TableForAssay" + (sampleCount), conditionsCount + 3, 1, XlPasteType.xlPasteAll);
                }

                // set sample name
                WorksheetUtilities.SetNamedRangeValue(sheet, "ResultsSampleStandards" + sampleCount, label, 1, 1);

                string linkingName = null;
                // handle AreaSampleStandards Table
                if (!string.IsNullOrEmpty(type))
                {
                    switch (type.ToUpper())
                    {
                        case "AREA":
                            header = "AreaSampleHeader";
                            HandleSampleType(sheet, "Area", "AreaData", ref areaCount, ref linkingName, conditionsCount, label);
                            break;

                        case "ASSAY":
                            header = "AssaySampleHeader";
                            HandleSampleType(sheet, "Assay", "AssayData", ref assayCount, ref linkingName, conditionsCount, label);
                            break;

                        case "LABEL CLAIM":
                            header = "LabelClaimHeader";
                            HandleSampleType(sheet, "Claim", "ClaimData", ref claimCount, ref linkingName, conditionsCount, label);
                            break;
                    }
                }

                // link Results cell to raw data and others

                WorksheetUtilities.LinkTwoNamedRangeCellsWrapper(sheet, linkingName, "ResultSamplesLinking" + sampleCount, 4 /* sourceCol */, 3 /* targetCol */, true /* addRound */);
                WorksheetUtilities.LinkTwoNamedRangeCellsWrapper(sheet, linkingName, "ResultSamplesLinking" + sampleCount, 5 /* sourceCol */, 4 /* targetCol */, true /* addRound */);

                if (!string.IsNullOrEmpty(header))
                {
                    WorksheetUtilities.LinkTwoNamedRangeCellsWrapper(sheet, header, "ResultsSampleStandards" + sampleCount, 4 /* sourceCol */, 3 /* targetCol */);
                    WorksheetUtilities.LinkTwoNamedRangeCellsWrapper(sheet, header, "ResultsSampleStandards" + sampleCount, 5 /* sourceCol */, 4 /* targetCol */);

                    if (header != "LabelClaimHeader")
                    {
                        WorksheetUtilities.LinkTwoNamedRangeCellsWrapper(sheet, header, "ResultsSampleStandards" + sampleCount, 7 /* sourceCol */, 2 /* targetCol */);
                        WorksheetUtilities.LinkTwoNamedRangeCellsWrapper(sheet, linkingName, "ResultSamplesLinking" + sampleCount, 7 /* sourceCol */, 2 /* targetCol */);
                    }
                    else
                    {
                        WorksheetUtilities.LinkTwoNamedRangeCellsWrapper(sheet, header, "ResultsSampleStandards" + sampleCount, 10 /* sourceCol */, 2 /* targetCol */);
                        WorksheetUtilities.LinkTwoNamedRangeCellsWrapper(sheet, linkingName, "ResultSamplesLinking" + sampleCount, 10 /* sourceCol */, 2 /* targetCol */);
                    }
                }

                LinkConcatenatedABPairsVertically(sheet, linkingName, "ResultSamplesLinking" + sampleCount);
            }// end each sample loop
        }

        private static void HandleImpuritySamples(Worksheet sheet, int conditionsCount, int numImpurities, int impuritySampleNumber)
        {
            if (impuritySampleNumber > 1)
            {
                for (var x = 1; x <= impuritySampleNumber; x++)
                {
                    // handle Impurity table
                    // below is true when i != 1
                    if (!WorksheetUtilities.NamedRangeExist(sheet, "ImpuritySampleStandards" + x))
                    {
                        // add new named table
                        // Insert a range between the ranges numbered 1 & 2, copy from range 1 to new range (copy all)
                        WorksheetUtilities.InsertRowsIntoNamedRange(conditionsCount + 2, sheet, "ImpurityData", false, XlDirection.xlUp, XlPasteType.xlPasteAll);
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ImpuritySampleStandards" + (x - 1), "ImpuritySampleStandards" + x, conditionsCount + 3, 1, XlPasteType.xlPasteAll);
                        //Add Conditions for Linking to summary table
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ImpurityConditions" + (x - 1), "ImpurityConditions" + x, conditionsCount + 3, 1, XlPasteType.xlPasteAll);
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ImpurityDifference1" + (x - 1), "ImpurityDifference1" + x, conditionsCount + 3, 1, XlPasteType.xlPasteAll);
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ImpurityDifferenceConditions1" + (x - 1), "ImpurityDifferenceConditions1" + x, conditionsCount + 3, 1, XlPasteType.xlPasteAll);

                        WorksheetUtilities.InsertRowsIntoNamedRange(conditionsCount + 4, sheet, "ImpuritySummaryData", false, XlDirection.xlUp, XlPasteType.xlPasteAll);
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ImpuritySummarySample" + (x - 1), "ImpuritySummarySample" + x, conditionsCount + 5, 1, XlPasteType.xlPasteAll);
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ImpuritySummaryDifference1" + (x - 1), "ImpuritySummaryDifference1" + x, conditionsCount + 5, 1, XlPasteType.xlPasteAll);
                        //Conditions of summary table
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ImpuritySummaryDefinitions" + (x - 1), "ImpuritySummaryDefinitions" + x, conditionsCount + 5, 1, XlPasteType.xlPasteAll);
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ImpuritySummaryConditions1" + (x - 1), "ImpuritySummaryConditions1" + x, conditionsCount + 5, 1, XlPasteType.xlPasteAll);

                        // link names of Samples
                        WorksheetUtilities.LinkFirstCell(sheet, "ImpuritySampleStandards" + x, "ImpuritySummarySample" + x);

                        for (var i = 1; i <= numImpurities; i++)
                        {
                            if (i != 1)
                            {
                                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ImpuritySummaryDifference" + (i - 1) + x, "ImpuritySummaryDifference" + i + x, 1, 2, XlPasteType.xlPasteAll);
                                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ImpurityDifference" + (i - 1) + x, "ImpurityDifference" + i + x, 1, DefaultImpurityCols + 1, XlPasteType.xlPasteAll);


                                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ImpuritySummaryConditions" + (i - 1) + x, "ImpuritySummaryConditions" + i + x, 1, 2, XlPasteType.xlPasteAll);
                                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ImpurityDifferenceConditions" + (i - 1) + x, "ImpurityDifferenceConditions" + i + x, 1, DefaultImpurityCols + 1, XlPasteType.xlPasteAll);

                                LinkImpurityDifference(sheet, "ImpurityDifference" + i + x, "ImpuritySummaryDifference" + i + x, "ImpuritySummaryColumn" + i);
                                LinkImpurityDifference(sheet, "ImpurityDifferenceConditions" + i + x, "ImpuritySummaryConditions" + i + x, "ImpuritySummaryColumn" + i);
                            }
                            else
                            {
                                WorksheetUtilities.LinkTwoNamedRangeCellsWrapper(sheet, "ImpurityDifference" + i + x, "ImpuritySummaryDifference" + i + x, 1 /* sourceCol */, 1 /* targetCol */, false /* roundForImpurity */);
                                WorksheetUtilities.LinkTwoNamedRangeCellsWrapper(sheet, "ImpurityDifferenceConditions" + i + x, "ImpuritySummaryConditions" + i + x, 1 /* sourceCol */, 1 /* targetCol */, true /* roundForImpurity */);
                            }
                        }

                        // set sample name 
                        string label = ("Sample" + x).Trim();

                        WorksheetUtilities.SetNamedRangeValue(sheet, "ImpuritySampleStandards" + x, label, 1, 1);

                        //Link ImpuritySummaryTable with ImpuritySamples
                        LinkConcatenatedABPairsVertically(sheet, "ImpurityConditions" + x, "ImpuritySummaryDefinitions" + x);

                        WorksheetUtilities.ResizeNamedRange(sheet, "TableForImpurity", conditionsCount + 1, 0);
                        //08-07 - Code for new Replication
                        WorksheetUtilities.ResizeNamedRange(sheet, "ImpuritySummaryReplication1", conditionsCount + 4, 0);
                        for (var i = 1; i <= numImpurities; i++)
                        {
                            WorksheetUtilities.ResizeNamedRange(sheet, "ImpuritySummaryColumn" + i, conditionsCount + 4, 0);
                        }
                    }
                    else
                    {
                        string label = "Sample1".Trim();

                        WorksheetUtilities.SetNamedRangeValue(sheet, "ImpuritySampleStandards1", label, 1, 1);
                    }
                }
            }
            else if (impuritySampleNumber == 0)
            {
                //Added as 30/03 Comments
                WorksheetUtilities.DeleteNamedRangeRows(sheet, "ToDeleteImpuritySection");
            }
            else
            {
                string label = "Sample1".Trim();

                WorksheetUtilities.SetNamedRangeValue(sheet, "ImpuritySampleStandards1", label, 1, 1);
            }

            // extend impurity range
            for (int i = 1; i <= numImpurities; i++)
            {
                string impurityRangName = "Impurity" + i;

                if (WorksheetUtilities.NamedRangeExist(sheet, impurityRangName))
                {
                    WorksheetUtilities.ResizeNamedRange(sheet, impurityRangName, (conditionsCount + 2) * (impuritySampleNumber - 1), 0);
                }
            }
        }

        private static void HandleLargeNumberOfImpurities(Worksheet sheet, int numImpurities)
        {
            //08-07 - Replication Code for more than 6 Impurities - ValidationReportIssue
            if (numImpurities > 5)
            {
                int repOrder = 1;
                int rowsToInsert = WorksheetUtilities.GetNamedRangeRowCount(sheet, "ImpuritySummaryReplication1");
                for (var i = 1; i <= numImpurities; i++)
                {

                    if (i % 5 == 0)
                    {
                        repOrder = repOrder + 1;
                        Range namedRangeForLastRow = WorksheetUtilities.GetNamedRange(sheet, "ImpuritySummaryData");
                        //int indexToInsertConditions = namedRangeForLastRow.SpecialCells(XlCellType.xlCellTypeLastCell,Type.Missing).Row;
                        int indexToInsertFrom = WorksheetUtilities.GetNamedRangeRowCount(sheet, "ImpuritySummaryData");
                        Range baseCell = (Range)namedRangeForLastRow.Cells[indexToInsertFrom, 1];
                        string baseAddress = baseCell.get_AddressLocal(false, false, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

                        //Regex to replace Address letter and leave only the number
                        string pattern = "[a-zA-Z]";
                        string replacement = "";

                        string result = Regex.Replace(baseAddress, pattern, replacement);

                        int indexToInsertConditions = int.Parse(result);

                        //ImpuritySummaryData - First column w/ last row empty

                        WorksheetUtilities.InsertRowsIntoNamedRangeFromRow(rowsToInsert, sheet, "ImpuritySummaryData", true, XlDirection.xlDown, XlPasteType.xlPasteAll, indexToInsertFrom);
                        //Copy first column into the new rows

                        //int copyAtRow = previouseRange.SpecialCells(XlCellType.xlCellTypeLastCell).Row
                        //We take first the formatting, and then the values, we cannot do a direct copy as there are formula in the condition columns
                        WorksheetUtilities.CopyNamedRangeToNewLocationWithNewNamedRange(sheet, "ImpuritySummaryReplication1", "ImpuritySummaryReplication" + repOrder, indexToInsertConditions, 2, XlPasteType.xlPasteFormats);
                        WorksheetUtilities.LinkTwoNamedRangeCells(sheet, "ImpuritySummaryReplication1", "ImpuritySummaryReplication" + repOrder, -1, 1, -1, 1, false, false);

                        //Cut the sixth and every following column and paste them in the new location
                        int impnumber = i;

                        Range impToMove = null;
                        Range conditionsRange = null;
                        int colToResize = 0;

                        for (var x = 1; x <= 5; x++)
                        {
                            //Search Named range to move
                            impToMove = WorksheetUtilities.GetNamedRange(sheet, "ImpuritySummaryColumn" + (impnumber + x));
                            conditionsRange = WorksheetUtilities.GetNamedRange(sheet, "ImpuritySummaryReplication" + repOrder);

                            if (impToMove == null)
                            {
                                break;
                            }
                            else
                            {
                                //variables to get where to move the range

                                Range firstCell = conditionsRange.Cells[1, 1];

                                //Get values as integers
                                int rowOffset = firstCell.Row;
                                int colOffset = firstCell.Column + x;
                                WorksheetUtilities.MoveNamedRange(sheet, "ImpuritySummaryColumn" + (impnumber + x), rowOffset, colOffset);

                                colToResize = colToResize + 1;
                                //Clean
                                WorksheetUtilities.ReleaseComObject(firstCell);
                                WorksheetUtilities.ReleaseComObject(impToMove);
                            }
                        }

                        //Resize original namedRange (Final Step)
                        WorksheetUtilities.ResizeNamedRange(sheet, "TableForImpurity", rowsToInsert, 0);
                        if (colToResize != 0)
                        {
                            WorksheetUtilities.ResizeNamedRange(sheet, "TableForImpurity", 0, -(colToResize));
                        }

                        //Clean
                        WorksheetUtilities.ReleaseComObject(conditionsRange);
                        WorksheetUtilities.ReleaseComObject(impToMove);
                        WorksheetUtilities.ReleaseComObject(namedRangeForLastRow);
                        WorksheetUtilities.ReleaseComObject(baseCell);
                    }
                }
            }
        }

        private static void ProcessSNRatio(Worksheet sheet, int conditionsCount, SNRatio snRatio)
        {
            ProcessSNRatioResolution(sheet, conditionsCount, snRatio.SNNumberOfPeaks, snRatio.SNNumberOfSolutions, snRatio.StrSNOperator, snRatio.DecSNValue, "SN");
        }

        private static void ProcessSNRatioSummary(Worksheet sheet, int conditionsCount, SNRatio snRatio)
        {
            ProcessSNResolutionSummary(sheet, conditionsCount, snRatio.SNNumberOfPeaks, snRatio.SNNumberOfSolutions, "SN");
        }

        private static void ProcessResolution(Worksheet sheet, int conditionsCount, PeaksInAgedSolution resolution)
        {
            ProcessSNRatioResolution(sheet, conditionsCount, resolution.PASNumberOfPeaks, resolution.PASNumberOfResPair, resolution.StrResolutionOperator, resolution.DecResolutionValue, "Resolution");
        }

        private static void ProcessResolutionSummary(Worksheet sheet, int conditionsCount, PeaksInAgedSolution resolution)
        {
            ProcessSNResolutionSummary(sheet, conditionsCount, resolution.PASNumberOfPeaks, resolution.PASNumberOfResPair, "Resolution");
        }

        // the named ranges for SN Ratio and Resolution are named similar with a difference of the word SN vs Resolution in either prefix.
        private static void ProcessSNRatioResolution(Worksheet sheet, int conditionsCount, int numReplicatesHorizontal, int numReplicatesVertical, string strOperator, decimal decValue, string prefix)
        {
            WorksheetUtilities.SetNamedRangeValue(sheet, $"{prefix}Operator", strOperator, 1, 1);
            WorksheetUtilities.SetNamedRangeValue(sheet, $"{prefix}Value", decValue.ToString(), 1, 1);

            // Add rows based on conditions count
            if (conditionsCount > DefaultNumConditions)
            {
                int numRowsToInsert = conditionsCount - DefaultNumConditions;
                WorksheetUtilities.InsertRowsIntoNamedRange(numRowsToInsert, sheet, $"{prefix}ConditionsTable1", true, XlDirection.xlDown, XlPasteType.xlPasteAll);
            }

            // Replicate horizontally based on number of peaks
            for (int i = 2; i <= numReplicatesHorizontal; i++)
            {
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"{prefix}Column" + (i - 1), $"{prefix}Column" + i, 1, DefaultSNResolutionCols + 1, XlPasteType.xlPasteAll);
                WorksheetUtilities.ResizeNamedRange(sheet, $"{prefix}SampleStandards1", 0, DefaultSNResolutionCols);
            }

            // Replicate vertically based on number of solution pairs
            for (int j = 2; j <= numReplicatesVertical; j++)
            {
                WorksheetUtilities.InsertRowsIntoNamedRange(conditionsCount + 2, sheet, $"{prefix}Data", false, XlDirection.xlUp, XlPasteType.xlPasteAll);
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"{prefix}Conditions" + (j - 1), $"{prefix}Conditions" + j, conditionsCount + 3, 1, XlPasteType.xlPasteAll);
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"{prefix}SampleStandards" + (j - 1), $"{prefix}SampleStandards" + j, conditionsCount + 3, 1, XlPasteType.xlPasteAll);
                WorksheetUtilities.SetNamedRangeValue(sheet, $"{prefix}SampleStandards" + j, $"Sample{j}", 1, 1);

                for (int i = 1; i <= numReplicatesHorizontal; i++)
                {
                    WorksheetUtilities.ResizeNamedRange(sheet, $"{prefix}Column" + i, conditionsCount + 2, 0);
                }
            }

            CopyEmpowerRelatedBlocksSNResolution(sheet, numReplicatesHorizontal, numReplicatesVertical, $"{prefix}EmpowerBlock", conditionsCount + 3, 4);
        }

        private static void ProcessSNResolutionSummary(Worksheet sheet, int conditionsCount, int numReplicatesHorizontal, int numReplicatesVertical, string prefix)
        {
            // Add rows based on conditions count
            if (conditionsCount > DefaultNumConditions)
            {
                int numRowsToInsert = conditionsCount - DefaultNumConditions;
                WorksheetUtilities.InsertRowsIntoNamedRange(numRowsToInsert, sheet, $"{prefix}SummaryConditionsTable1", true, XlDirection.xlDown, XlPasteType.xlPasteAll);
            }

            // Replicate horizontally based on number of peaks
            for (int i = 2; i <= numReplicatesHorizontal; i++)
            {
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"{prefix}SummaryColumn" + (i - 1), $"{prefix}SummaryColumn" + i, 1, DefaultSNResolutionSummaryCols + 1, XlPasteType.xlPasteAll);
                WorksheetUtilities.ResizeNamedRange(sheet, $"{prefix}SummarySampleStandards", 0, DefaultSNResolutionSummaryCols + 1);

                // Link Peak Names
                LinkPeakName(sheet, $"{prefix}Column" + i, $"{prefix}SummaryColumn" + i);
            }

            // Replicate vertically based on number of solution pairs
            for (int j = 2; j <= numReplicatesVertical; j++)
            {
                WorksheetUtilities.InsertRowsIntoNamedRange(conditionsCount + 2, sheet, $"{prefix}SummaryData", false, XlDirection.xlUp, XlPasteType.xlPasteAll);
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"{prefix}SummaryDefinitions" + (j - 1), $"{prefix}SummaryDefinitions" + j, conditionsCount + 3, 1, XlPasteType.xlPasteAll);
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"{prefix}SummarySampleStandards" + (j - 1), $"{prefix}SummarySampleStandards" + j, conditionsCount + 3, 1, XlPasteType.xlPasteAll);
            }

            CopyEmpowerRelatedBlocksSNResolution(sheet, numReplicatesHorizontal, numReplicatesVertical, $"{prefix}SummaryBlock", conditionsCount + 3, 2);
            LinkNamedRangesSNResolution(sheet, $"{prefix}EmpowerBlock", $"{prefix}SummaryBlock", $"{prefix}SummaryColumn", numReplicatesHorizontal, numReplicatesVertical);
        }

        private static void HandleSampleType(Worksheet sheet, string typePrefix, string dataRangeName, ref int typeCount, ref string linkingName, int conditionsCount, string label)
        {
            ++typeCount;
            string rangeName = $"{typePrefix}SampleStandards{typeCount}";
            linkingName = $"{typePrefix}SampleLinking{typeCount}";

            if (!WorksheetUtilities.NamedRangeExist(sheet, rangeName))
            {
                WorksheetUtilities.InsertRowsIntoNamedRange(conditionsCount + 2, sheet, dataRangeName, false, XlDirection.xlUp, XlPasteType.xlPasteAll);
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"{typePrefix}SampleStandards{typeCount - 1}", rangeName, conditionsCount + 3, 1, XlPasteType.xlPasteAll);
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, $"{typePrefix}SampleLinking{typeCount - 1}", linkingName, conditionsCount + 3, 1, XlPasteType.xlPasteAll);
            }

            WorksheetUtilities.SetNamedRangeValue(sheet, rangeName, label, 1, 1);
        }

        private static void DeleteVerticalRangeShiftLeft(Worksheet sheet)
        {
            string nameRangeToDelete = "AssaySummaryTimepointsCol";
            Range rangeToDelete = sheet.Range[nameRangeToDelete];

            if (rangeToDelete != null)
            {
                rangeToDelete.Delete(XlDeleteShiftDirection.xlShiftToLeft);
                WorksheetUtilities.ReleaseComObject(rangeToDelete);
            }
        }

        // helper functions below here

        /// <summary>
        /// Updates the Difference formulas to use correct cell address(es)
        /// Copy formula from base cell down, only update current cell address
        /// </summary>
        /// <param name="sheet">The worksheet</param>
        /// <param name="namedRangeBaseName">The base named range to update.</param>
        /// <param name="baseCellRow">The row number of the base cell (from which the base formula is gotten).</param>
        /// <param name="baseCellCol">The column number of the base cell (from which the base formula is gotten).</param>
        /// <param name="colOffset">The column off set number from each cell to get reference address (from which the base formula is gotten).</param>
        private static void UpdateDifferenceFormulas(_Worksheet sheet, string namedRangeBaseName, int baseCellRow, int baseCellCol, int colOffset)
        {
            if (sheet == null || string.IsNullOrEmpty(namedRangeBaseName))
            {
                return;
            }

            Name name = null;
            Range range = null;

            try
            {
                name = GetNamedRange(sheet, namedRangeBaseName);
                if (name == null)
                {
                    return;
                }

                range = name.RefersToRange;
                if (range == null)
                {
                    return;
                }

                UpdateFormulasInRange(range, baseCellRow, baseCellCol, colOffset);
            }
            finally
            {
                WorksheetUtilities.ReleaseComObject(name);
                WorksheetUtilities.ReleaseComObject(range);
            }
        }

        /// <summary>
        /// Updates the Result formulas to use correct cell address(es)
        /// Copy formula from base cell down, only update current cell address
        /// </summary>
        /// <param name="sheet">The worksheet</param>
        /// <param name="namedRangeBaseName">The base named range to update.</param>
        /// <param name="baseCellRow">The row number of the base cell (from which the base formula is gotten).</param>
        /// <param name="baseCellCol">The column number of the base cell (from which the base formula is gotten).</param>
        /// <param name="colOffset">The column off set number from each cell to get reference address (from which the base formula is gotten).</param>
        private static void UpdateResultFormulas(_Worksheet sheet, string namedRangeBaseName, int baseCellRow, int baseCellCol, int colOffset, int staticCellRow, int staticCellColumn)
        {
            if (sheet == null || string.IsNullOrEmpty(namedRangeBaseName))
            {
                return;
            }

            Name name = null;
            Range range = null;

            try
            {
                name = GetNamedRange(sheet, namedRangeBaseName);
                if (name == null)
                {
                    return;
                }

                range = name.RefersToRange;
                if (range == null)
                {
                    return;
                }

                UpdateFormulasInRange(range, baseCellRow, baseCellCol, colOffset, staticCellRow, staticCellColumn);
            }
            finally
            {
                WorksheetUtilities.ReleaseComObject(name);
                WorksheetUtilities.ReleaseComObject(range);
            }
        }

        private static void UpdateFormulasInRange(Range range, int baseCellRow, int baseCellCol, int colOffset, int? staticCellRow = null, int? staticCellColumn = null)
        {
            Range baseCell = range.Cells[baseCellRow, baseCellCol] as Range;
            if (baseCell == null)
            {
                return;
            }

            string baseFormula = baseCell.FormulaLocal.ToString();
            string baseAddress = ((Range)baseCell.Cells[1, colOffset]).get_AddressLocal(false, false, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

            WorksheetUtilities.ReleaseComObject(baseCell);

            for (int j = 1; j <= range.Rows.Count - 1; j++)
            {
                Range cell = range.Cells[baseCellRow + j, baseCellCol] as Range;

                if (cell != null)
                {
                    if (staticCellRow.HasValue && staticCellColumn.HasValue)
                    {
                        Range pointCell = range.Cells[staticCellRow.Value, staticCellColumn.Value] as Range;
                        if (pointCell != null)
                        {
                            string newCellAddress = ((Range)pointCell.Cells[1, 1]).get_AddressLocal(false, false, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                            string addressToReplace = ((Range)pointCell.Cells[1 + j, 1]).get_AddressLocal(false, false, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                            cell.Formula = cell.Formula.Replace(addressToReplace, newCellAddress);

                            WorksheetUtilities.ReleaseComObject(pointCell);
                        }
                    }
                    else
                    {
                        string newCellAddress = ((Range)cell.Cells[1, colOffset]).get_AddressLocal(false, false, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                        cell.Formula = baseFormula.Replace(baseAddress, newCellAddress);
                    }

                    WorksheetUtilities.ReleaseComObject(cell);
                }
            }
        }

        private static void LinkConcatenatedABPairsVertically(Worksheet sheet, string sourceNamedRange, string destinationNamedRange)
        {
            // Retrieve named ranges
            Range sourceRange = sheet.Range[sourceNamedRange];
            Range destinationRange = sheet.Range[destinationNamedRange];

            int rowCount = sourceRange.Rows.Count;

            for (int i = 2; i <= rowCount; i++)
            {
                // Get the first two columns of the source range for this row
                Range sourceCell1 = sourceRange.Cells[i, 1] as Range; // First column in source range
                Range sourceCell2 = sourceRange.Cells[i, 2] as Range; // Second column in source range

                // Get the corresponding destination cell
                Range destinationCell = destinationRange.Cells[i, 1] as Range;

                if (sourceCell1 != null && sourceCell2 != null && destinationCell != null)
                {
                    string addr1 = sourceCell1.get_Address(false, false, XlReferenceStyle.xlA1);
                    string addr2 = sourceCell2.get_Address(false, false, XlReferenceStyle.xlA1);

                    destinationCell.Formula = $"=IF(AND({addr1}<>\"\",{addr2}<>\"\"),CONCATENATE({addr1},\"_\",{addr2}),\"\")";

                    WorksheetUtilities.ReleaseComObject(sourceCell1);
                    WorksheetUtilities.ReleaseComObject(sourceCell2);
                    WorksheetUtilities.ReleaseComObject(destinationCell);
                }
            }

            WorksheetUtilities.ReleaseComObject(sourceRange);
            WorksheetUtilities.ReleaseComObject(destinationRange);
        }

        private static void LinkImpurityDifference(Worksheet sheet, string srcNamedRange, string destNamedRange, string decimalPrecisionNamedRange)
        {
            // Retrieve named ranges
            Range sourceRange = sheet.Range[srcNamedRange];
            Range destinationRange = sheet.Range[destNamedRange];
            Range decimalPrecisionRange = sheet.Range[decimalPrecisionNamedRange];

            // Get the second cell of the decimal precision range
            Range fixedDecimalCell = decimalPrecisionRange.Cells[2, 1] as Range;
            string fixedDecimalCellAddr = fixedDecimalCell.get_Address(true, true, XlReferenceStyle.xlA1); // Absolute reference

            int rowCount = sourceRange.Rows.Count;

            for (int i = 1; i <= rowCount; i++)
            {
                Range srcCell = sourceRange.Cells[i, 1] as Range;
                Range destCell = destinationRange.Cells[i, 1] as Range;

                if (srcCell != null && destCell != null)
                {
                    string srcAddress = srcCell.get_Address(false, false, XlReferenceStyle.xlA1);
                    destCell.Formula = $"=IF(ISNUMBER({srcAddress}), FIXED({srcAddress}, {fixedDecimalCellAddr}), IF({srcAddress}<>\"\", {srcAddress}, \"\"))";

                    WorksheetUtilities.ReleaseComObject(srcCell);
                    WorksheetUtilities.ReleaseComObject(destCell);
                }
            }

            // Release COM objects
            WorksheetUtilities.ReleaseComObject(fixedDecimalCell);
            WorksheetUtilities.ReleaseComObject(sourceRange);
            WorksheetUtilities.ReleaseComObject(destinationRange);
            WorksheetUtilities.ReleaseComObject(decimalPrecisionRange);
        }

        private static void LinkPeakName(Worksheet sheet, string columnNamedRange, string summaryColumnNamedRange)
        {
            // Retrieve the named ranges
            Range sourceRange = sheet.Range[columnNamedRange];
            Range destinationRange = sheet.Range[summaryColumnNamedRange];

            // Ensure there are cells in both named ranges
            if (sourceRange.Cells.Count > 0 && destinationRange.Cells.Count > 1)
            {
                // Get the first cell of the first named range
                Range firstCellOfSource = sourceRange.Cells[1, 1] as Range;
                string sourceAddress = firstCellOfSource.Address[false, false, XlReferenceStyle.xlA1];

                // Get the second cell of the second named range
                Range secondCellOfDestination = destinationRange.Cells[2, 1] as Range;

                // Set the formula in the second cell of the destination range
                secondCellOfDestination.Formula = $"=IF({sourceAddress}=\"\",\"\",{sourceAddress})";

                // Release COM objects
                WorksheetUtilities.ReleaseComObject(firstCellOfSource);
                WorksheetUtilities.ReleaseComObject(secondCellOfDestination);
            }

            // Release COM objects
            WorksheetUtilities.ReleaseComObject(sourceRange);
            WorksheetUtilities.ReleaseComObject(destinationRange);
        }

        private static void CopyEmpowerRelatedBlocksSNResolution(Worksheet sheet, int x, int y, string namedRangeBase, int rowOffset, int colOffset)
        {
            for (int i = 0; i < x; i++)
            {
                for (int j = 0; j < y; j++)
                {
                    if (i == 0 && j == 0)
                    {
                        continue; // Skip the top-left block
                    }

                    string destLetterCombo = GetLetterCombination(i, j);
                    string destNamedRange = namedRangeBase + destLetterCombo;

                    string srcNamedRange;

                    int rowIndex = rowOffset;       // rows down per solution
                    int colIndex = colOffset;   // columns over per peak


                    if (j == 0)
                    {
                        // Copy from the block to the left
                        string srcLetterCombo = GetLetterCombination(i - 1, j);
                        srcNamedRange = namedRangeBase + srcLetterCombo;
                        rowIndex = 1;
                    }
                    else
                    {
                        // Copy from the block above
                        string srcLetterCombo = GetLetterCombination(i, j - 1);
                        srcNamedRange = namedRangeBase + srcLetterCombo;
                        colIndex = 1;
                    }

                    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, srcNamedRange, destNamedRange, rowIndex, colIndex, XlPasteType.xlPasteAll);
                }
            }
        }

        private static void LinkNamedRangesSNResolution(Worksheet sheet, string srcBaseName, string destBaseName, string summaryColumnBaseName, int x, int y)
        {
            for (int i = 0; i < x; i++)
            {
                for (int j = 0; j < y; j++)
                {
                    if (i == 0 && j == 0)
                    {
                        continue; // Skip the top-left block
                    }

                    // Generate letter combination suffix
                    string srcLetterCombo = GetLetterCombination(i, j);
                    string srcNamedRange = srcBaseName + srcLetterCombo;

                    string destLetterCombo = GetLetterCombination(i, j);
                    string destNamedRange = destBaseName + destLetterCombo;

                    LinkNamedRangeSNResolution(sheet, srcNamedRange, destNamedRange, summaryColumnBaseName + (i + 1));
                }
            }
        }

        private static void LinkNamedRangeSNResolution(Worksheet sheet, string srcNamedRange, string destNamedRange, string summaryColumnNamedRange)
        {
            // Retrieve named ranges
            Range sourceRange = sheet.Range[srcNamedRange];
            Range destinationRange = sheet.Range[destNamedRange];
            Range summaryColumnRange = sheet.Range[summaryColumnNamedRange];

            // Get the first cell address of the fixed decimal range
            Range fixedDecimalCell = summaryColumnRange.Cells[1, 1] as Range;
            string fixedDecimalCellAddr = fixedDecimalCell.get_Address(false, false, XlReferenceStyle.xlA1);

            int rowCount = sourceRange.Rows.Count;

            for (int k = 1; k <= rowCount; k++)
            {
                Range srcCell = sourceRange.Cells[k, 1] as Range;
                Range destCell = destinationRange.Cells[k, 1] as Range;

                if (srcCell != null && destCell != null)
                {
                    string srcAddress = srcCell.get_Address(false, false, XlReferenceStyle.xlA1);
                    destCell.Formula = $"=IF(ISNUMBER({srcAddress}),FIXED({srcAddress},{fixedDecimalCellAddr}),IF({srcAddress}=\"\",\"\",{srcAddress}))";

                    WorksheetUtilities.ReleaseComObject(srcCell);
                    WorksheetUtilities.ReleaseComObject(destCell);
                }
            }

            WorksheetUtilities.ReleaseComObject(sourceRange);
            WorksheetUtilities.ReleaseComObject(destinationRange);
        }

        private static Name GetNamedRange(_Worksheet sheet, string namedRangeBaseName)
        {
            try
            {
                return sheet.Names.Item(namedRangeBaseName, Type.Missing, Type.Missing) as Name;
            }
            catch
            {
                return null;
            }
        }

        private static List<string> GenerateSampleNames(string prefix, int count, string methodName)
        {
            List<string> samples = new List<string>();

            for (int i = 1; i <= count; i++)
            {
                samples.Add(prefix + "Sample" + i + ", " + methodName);
            }

            return samples;
        }

        // assume input is within range of [0, 25]
        private static string GetLetterCombination(int i, int j)
        {
            char[] letters = "abcdefghijklmnopqrstuvwxyz".ToCharArray();

            return $"{letters[i]}{letters[j]}";
        }
    }
}
