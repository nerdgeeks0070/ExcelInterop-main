using Spreadsheet.Handler;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;

namespace TestExcelLauncher
{
    public class Program
    {
        // change Excel directory
        static string excelDirectory = @"C:\Internal_GitHub\Internal_ABDValidationReport\trunk\TestExcelLauncher\ExcelToTest";

        public static void Main(string[] args)
        {
            // set file name
            string fileName = "5_Injection Repeatability_Results.xls";

            string fullPath = System.IO.Path.Combine(excelDirectory, fileName);

            if (!File.Exists(fullPath))
            {
                Console.WriteLine("File doesn't exist! Put the Excel file to test in the directory.");
                Console.WriteLine("Press any key to continue...");

                // Pause until user presses a key
                Console.ReadKey();
                return;
            }

            Console.WriteLine("Generating Result Excel...");


            string returnPath = TestInjectionRepeatabilitySheet(fullPath);

            // this will be the return path - C:\ProgramData\Internal Technologies\ABD_TempFiles
            try
            {
                Application excelApp = new Application();
                Workbook workbook = excelApp.Workbooks.Open(returnPath);

                // Make Excel visible
                excelApp.Visible = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error opening Excel file: " + ex.Message);
            }
        }

        private static string TestAccuracy(string sourcePath)
        {
            return AccuracyNew.UpdateAccuracySheet(
                sourcePath: sourcePath,
                strcmbProtocolType: "NDA",
                strcmbProductType: "SDD",
                strcmbTestType: "Impurities"
            );
        }

        private static string TestInjectionRepeatabilitySheet(string sourcePath)
        {
            return InjectionRepeatability.UpdateInjectionRepeatabilitySheet(
                sourcePath: sourcePath,
                numInjections: 4,
                isAssayLevel: true,
                signAssayRSD: "<=",
                valueAssayRSD: 1.25m,
                numPeaksAssay: 3,
                isImpurityLevel: true,
                signImpurityRSD: ">",
                valueImpurityRSD: 0.75m,
                numPeaksImpurity: 2,
                cmbProtocolType: "PAV",
                cmbProductType: "Drug Product",
                cmbTestType: "Volatile");
        }

        private static string TestRobustness(string sourcePath)
        {
            return Robustness.UpdateRobustnessSheet(
                sourcePath,

                // General Fields
                strcmbProtocolType: "PAV",
                strcmbProductType: "Drug Product",
                strcmbTestType: "AssayLevel_Impurities",

                // Assay Level Parameters
                numNumSamples: 5,
                numNosArea: 6,
                numNosAssay: 3,
                numNosLC: 0,
                strCmbDiff: "<",
                valTxtDiff: 0.05m,

                // Impurity Level Parameters
                numNoP: 2,
                numNoS: 3,
                numNumConditionIV: 7,
                strCmbNoS: "%GC",
                strCmbOperator1: ">",
                valAC1: 0.01m,
                valAC2: 0.02m,
                valAC3: 0.03m,
                strCmbAbsoluteRelative1: "Absolute",
                strTxtoperator1: ">",
                valacceptancecriteria1: 0.015m,
                strCmbOperator2: "<",
                valAC4: 0.04m,
                valAC5: 0.05m,
                strTxtoperator2: "<",
                valacceptancecriteria3: 0.025m,

                // Water Content Parameters
                numNumSamplesWC: 4,
                numNumConditionWC: 2,
                strCmbWaterContent1: "Low",
                valAC6: 0.1m,
                strCmbWaterContent3: "Medium",
                valAC7: 0.2m,
                strCmbWaterContent2: "High",
                valAC8: 0.3m,
                strCmbWaterContent4: "Very High",
                valAC9: 0.4m,

                // Dissolution Parameters
                numNoSDisso: 3,
                numNumConditionDisso: 3,
                numNumTimepointsDisso: 6,

                // Dissolution Criteria (Recoveries)
                strCmbCB1Recoveries: "Yes",
                valRecoveriesTB1AccCriteria: 90.0m,
                strCmbCB2Recoveries: "No",
                valRecoveriesTB2AccCriteria: 95.0m,
                valRecoveriesCBAcceptanceCriteria: 92.5m,

                // Dissolution Criteria (Dissolved)
                strCmbCB1Dissolved: "Yes",
                valDissolvedTB1AccCriteria: 85.0m,
                strCmbCB2Dissolved: "No",
                valDissolvedTB2AccCriteria: 80.0m,
                valDissolvedCBAcceptanceCriteria: 82.5m
            );
        }

        private static string TestSystemSuitability(string sourcePath)
        {
            return SystemSuitability.UpdateSystemSuitabilitySheet(
                sourcePath,
                strcmbProtocolType: "ads",
                strcmbProductType: "asd",
                strcmbTestType: "Dissolution",
                numBlankInterference: 2,
                numSensitivity: 3,
                numRSD: 2,
                numStandardAgreement: 4,
                numTailingFactor: 2,
                numResolutionTest: 2,
                numTheoreticalPlates: 3,
                numPeakToValleyRatio: 1,
                numRetentionFactor: 2,
                numDetectability: 13,
                numStdRecovery: 2,
                numAvgTiterVal: 5,
                numOther: 6
            );
        }

        private static string TestSampleRepeatabilityAPI(string sourcePath)
        {
            return SampleRepeatabilityApi.UpdateSampleRepeatabilityApiSheet(
                sourcePath: sourcePath,
                numReplicates: 3,
                protocolType: "PAV",
                productType: "Drug Substance",
                testType: "Cleaning Verification",

                // Assay
                assayNumSamples: 6,
                rsdOperator: "<=",
                rsdValue: 2.5m,

                // Impurity
                impurityNumSamples: 2,
                impurityQuantitationType: "JPsn",
                impurityNumPeaks: 3,
                impurityOperator1: "<",
                impurityValue1: 0.05m,
                impurityValue2: 0.01m,
                impurityAutoOperator1: "<=",
                impurityAutoValue1: 0.02m,
                impurityOperator2: ">=",
                impurityValue3: 0.005m,
                impurityOperator4: "<",
                impurityValue4: 3.5m,
                impurityAutoOperator2: ">",
                impurityAutoValue2: 0.001m,
                impurityOperator5: "<=",
                impurityValue5: 0.03m,

                // Water Content
                wcNumSamples: 4,
                wcOperator1: "<",
                wcValue1: 3.0m,
                wcOperator2: "<=",
                wcValue2: 2.5m,
                wcOperator3: ">",
                wcValue3: 1.59m,
                wcOperator4: "<=",
                wcValue4: 2.0m
            );
        }

        private static string TestSampleRepeatability(string sourcePath)
        {
            return SampleRepeatabilityNew.UpdateSampleRepeatabilitySheet(
                sourcePath: sourcePath,
                numRepsGeneral: 7,
                strcmbProtocolType: "PAV",
                strcmbProductType: "SDD",
                strcmbTestType: "Impurities",

                // Assay
                numSamplesAssay: 3,
                strcmbRSD: "<=",
                valRSD1: 2.5m,

                // Content Uniformity
                numSamplesCU: 5,

                // Impurity
                numSamplesImp: 3,
                strcmbNoS: "JPsn",
                numPeaksImp: 5,
                strcmbOperator1Imp: "<",
                valAC1Imp: 0.05m,
                valAC2Imp: 0.01m,
                strAutoOperator1Imp: "<",
                strAutoOperator2Imp: "<=",
                valacceptancecriteria1Imp: 0.03m,
                valacceptancecriteria3Imp: 3.5m,
                strcmbOperator2Imp: ">=",
                valAC3Imp: 0.001m,
                valAC4Imp: 0.02m,
                valAC5Imp: 0.03m,
                strcmbOperator4Imp: "<",
                strcmbOperator5Imp: "<=",

                // Water Content
                numSamplesWC: 3,
                strcmbWaterContent1WC: "<",
                valAC6WC: 3.0m,
                strcmbWaterContent3WC: ">",
                valAC7WC: 1.59m,
                strcmbWaterContent2WC: "<=",
                valAC8WC: 2.5m,
                strcmbWaterContent4WC: "<=",
                valAC9WC: 2.0m,

                // Dissolution
                numSamplesDisso: 3,
                numRepsDisso: 9,
                strcmbOperator1Disso: "<",
                valAC1Disso: 5.0m,
                strcmbOperator2Disso: "<=",
                valAC2Disso: 4.5m,
                strcmbOperator3Disso: ">",
                strcmbOperator4Disso: ">=",
                valAC3Disso: 3.0m,
                strcmbOperator5Disso: "<",
                valAC4Disso: 2.5m,
                strcmbOperator6Disso: "<="
            );
        }

        private static string TestIntermediatePrecisionSheet(string sourcePath)
        {
            return IntermediatePrecision.UpdateIntermediatePrecisionSheet(
                sourcePath: sourcePath,
                strcmbProtocolType: "NDA",
                strcmbProductType: "SDD",
                strcmbTestType: "Impurities"
            );
        }

        private static string TestReproducibilitySheet(string sourcePath)
        {
            return Reproducibility.UpdateReproducibilitySheet(
                sourcePath: sourcePath,
                strcmbProtocolType: "NDA",
                strcmbProductType: "SDD",
                strcmbTestType: "Impurities"
            );
        }

        private static string TestSolutionStability(string sourcePath)
        {
            // General
            string strcmbProtocolType = "NDA";
            string strcmbProductType = "Cleaning Verification";
            string strcmbTestType = "Impurity";
            int numStorageConditions = 3;
            int numDataPoints = 2;
            int numNoSolsArea = 2;
            int numNoSolsAssay = 3;
            int numNoSolsLC = 3;
            decimal assayLevelValue = 9.5m;
            string assayLevelOperator = ">=";
            int numPeaks = 2;
            int numNoSols = 3;
            string strCmbNoSols = "EPsn"; // Quantitation type
            string impurityRangeOperator1 = "<=";
            decimal impurityRangeValue1 = 0.3m;
            decimal impurityRangeValue2 = 0.7m;
            string impurityRangeOperator2 = ">";
            string impurityDiffAbsRel = "Relative";
            decimal impurityDiffValue1 = 0.02m;
            decimal impurityDiffValue2 = 0.05m;
            decimal impurityDiffValue3 = 0.08m;
            int numSNNoPeaks = 2;
            int numSNNoSols = 3;
            string snOperator = ">";
            decimal snValue = 3.5m;
            int numNRNoPeaks = 3;
            int numNoResPair = 4;
            string resolutionOperator = ">=";
            decimal resolutionValue = 1.2m;

            // Method call
            return SolutionStability.UpdateSolutionStabilitySheet(
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
                resolutionValue
            );
        }

        private static string TestFilterBinding(string sourcePath)
        {
            // General
            string strcmbProtocolType = "NDA";
            string strcmbProductType = "Cleaning Verification";
            string strcmbTestType = "Dissolution";
            int numFilters = 7;

            // Assay
            int numSolsArea = 5;
            int numSolsAssay = 3;
            int numSolsLC = 2;
            decimal valRSD1 = 1.25m;
            string strcmbDifference = "<=";

            // Impurity
            int numPeaks = 2;
            int numSols = 3;
            string strcmbQuantitationType = "EPsn";
            string strcmbOperator1 = ">";
            decimal valAC1 = 0.5m;
            decimal valAC2 = 0.75m;
            string strcmbAbsoluteRelative1 = "Absolute";
            string strOperator1 = "<=";
            decimal valacceptancecriteria1 = 1.0m;
            string strcmbOperator2 = "<";
            decimal valAC3 = 0.3m;
            decimal valAC4 = 0.6m;
            string strOperator2 = ">=";
            decimal valacceptancecriteria3 = 0.9m;
            decimal valAC5 = 0.45m;

            return FilterBinding.UpdateFilterBindingSheet(
                sourcePath,
                strcmbProtocolType,
                strcmbProductType,
                strcmbTestType,
                numFilters,
                numSolsArea,
                numSolsAssay,
                numSolsLC,
                valRSD1,
                strcmbDifference,
                numPeaks,
                numSols,
                strcmbQuantitationType,
                strcmbOperator1,
                valAC1,
                valAC2,
                strcmbAbsoluteRelative1,
                strOperator1,
                valacceptancecriteria1,
                strcmbOperator2,
                valAC3,
                valAC4,
                strOperator2,
                valacceptancecriteria3,
                valAC5);
        }

        private static string TestIdentificationSheetUpdate(string sourcePath)
        {
            return Identification.UpdateIdentificationSheet(
                sourcePath: sourcePath,
                numSamples: 5,
                hasRetentionLevel: true,
                numInjections: 3,
                rtr1Operator: ">=",
                rtr1Value: 0.80m,
                rtr2Operator: "<",
                rtr2Value: 1.05m,
                hasUVLevel: true,
                numLambdaMax: 2,
                lambdaMaxValue: 25.5m,
                cmbValidation: "PAV",
                cmbProduct: "Drug Product");
        }

        private static string TestSensitivity(string sourcePath)
        {
            int tbNumOfPeaks = 3;
            bool isChkRL_RSD = true;
            bool isChkRL_SN = true;
            bool isChkDL_SN = true;

            var numParams = new Dictionary<string, int>
            {
                { "txtNumReps", 10 },
                { "txtNumRepsDL", 10 }
            };

                    var acceptanceCriteria = new Dictionary<string, string>
            {
                { "cmbValidation", "NDA" },
                { "cmbProduct", "cmbProduct" },
                { "cmbQuantitativeType", "cmbQuantitativeType" },
                { "cmbRLDL", "Yes" },
                { "cmbRL_RSD", "≤" },
                { "TBRL_RSD", "10" },
                { "cmbRL_SN", "≥" },
                { "cmbDL_SN", "≥" },
                { "TBDL_SN", "10" }
            };

            return ImpuritySensitivity.UpdateImpSensitivitySheet(
                sourcePath,
                tbNumOfPeaks,
                isChkRL_SN,
                isChkRL_RSD,
                isChkDL_SN,
                numParams,
                acceptanceCriteria
            );
        }

        private static string TestSampleSizeRobustness(string sourcePath)
        {
            return SampleSizeRobustness.UpdateSampleSizeRobustnessSheet(
                sourcePath: sourcePath,
                strcmbProtocolType: "PAV",
                strcmbProductType: "SDD",
                numSamples: 5,
                numReps100pct: 7,
                numLevels: 3,
                numRepsExcluding100pct: 4,
                wcOperator1: ">",
                wcValue1: 0.75m,
                wcOperator3: "<",
                wcValue2: 1.25m,
                strcmbAbsoluteRelative1: "Absolute",
                wcOperator2: ">",
                wcValue3: 3.59m,
                wcOperator4: ">=",
                wcValue4: 0.0050m);
        }

        private static string TestLinearity(string sourcePath)
        {
            return Linearity.UpdateLinearitySheet(
                sourcePath: sourcePath,
                strcmbProtocolType: "PAV",
                strcmbProductType: "SDD",
                strcmbTestType: "Impurity",
                numReps: 5,
                numPeaks1: 3,
                numPeaks2: 2,
                numPeaks3: 6,
                numLevel1: 4,
                numLevel2: 5,
                numLevel3: 3, // test less than 1
                strcmbRRF: "Yes",
                strcmbR: ">=",
                strcmbY: "<",
                rValue: 0.987m,
                yInterceptValue: 0.0125m
            );
        }

        private static string TestSpecify(string sourcePath)
        {
            return Specificity.UpdateSpecificitySheet(sourcePath: sourcePath,
                numPeaks: 5,
                numSamples: 6,
                numSolPeakPurity: 6,
                numSolDissol: 3,
                cmbProtocolType: "cmbProtocolType",
                cmbProductType: "Drug Product",
                cmbTestType: "Volatiles");
            //Drug Substance,Drug Product,SDD
            //Volatiles,Water content
        }

        private static string TestDissoManuvsAuto(string sourcePath)
        {
            Dictionary<string, string> acceptanceCriteria = new Dictionary<string, string>() {
            { "ValidationType", "NDA" }, { "CBProduct", "Drug Product" },
            { "RecoveriesOperator1", "<=" }, { "RecoveriesValue1", "10" },
            { "RecoveriesOperator2", "<=" }, { "RecoveriesValue2", "15" },
            { "DissolvedOperator1", "<" }, { "DissolvedValue1", "10" },
            { "DissolvedOperator2", "<" }, { "DissolvedValue2", "5" }
            };

            return DissoManuvsAuto.UpdateDissoManuvsAutoSheet(sourcePath: sourcePath,
                compSet: 7,
                timepoints: 6,
                acceptanceCriteria: acceptanceCriteria
                );
            //Drug Substance,Drug Product,SDD
            //Volatiles,Water content

        }

        private static string TestAlternateAutomatedSampliing(string sourcePath)
        {
            Dictionary<string, string> acceptanceCriteria = new Dictionary<string, string>() {
            { "ValidationType", "NDA" }, { "CBProduct", "Drug Product" },
            { "RecoveriesOperator1", "<=" }, { "RecoveriesValue1", "10" },
            { "RecoveriesOperator2", "<=" }, { "RecoveriesValue2", "15" },
            { "DissolvedOperator1", "<" }, { "DissolvedValue1", "10" },
            { "DissolvedOperator2", "<" }, { "DissolvedValue2", "5" }
            };

            return AlternateAutomatedSampling.UpdateAlternateAutomatedSamplingSheet(sourcePath: sourcePath,
                compSet: 5,
                timepoints: 7,
                acceptanceCriteria: acceptanceCriteria
                );
        }

        private static string TestSinglePointVsProfile(string sourcePath)
        {
            Dictionary<string, string> acceptanceCriteria = new Dictionary<string, string>() {
            { "cmbProtocolType", "NDA" }, { "cmbProductType", "Drug Product" },
            { "cmbDiff", "<=" }, { "txtdiff", "10" },
            { "cmbQTP", "<=" }, { "txtQTP", "15" },
            { "txtTP", "30" }
            };

            return SinglePointVsProfile.UpdateSinglePointVsProfileSheet(sourcePath: sourcePath,
                compSet: 4,
                acceptanceCriteria: acceptanceCriteria
                );
        }

        //    private static string TestCompoundInformation(string sourcePath)
        //    {
        //        // Step 1: Define your new reference-based input dynamically
        //        var referenceList = new List<ReferenceRowRequest>()
        //{
        //    new ReferenceRowRequest { ReferenceType = "Reference_Standard_1", Count = 3 },
        //    new ReferenceRowRequest { ReferenceType = "Impurity_Mixture_1", Count = 2 },
        //    //new ReferenceRowRequest { ReferenceType = "SDD_Sample", Count = 1 },
        //    //new ReferenceRowRequest { ReferenceType = "Drug_Product_Peak_Sample", Count = 2 },
        //    //new ReferenceRowRequest { ReferenceType = "Drug_Substance_Peak_Sample", Count = 2 },
        //    //new ReferenceRowRequest { ReferenceType = "SDD_Peak_Sample", Count = 2 },
        //    //new ReferenceRowRequest { ReferenceType = "Drug_Product_Summary1", Count = 1 },
        //    //new ReferenceRowRequest { ReferenceType = "SDD_Summary1", Count = 1 }
        //};

        //        // Step 2: Call your new dynamic method instead of the older one
        //        string updatedFilePath = CompoundInformation.UpdateSpecificitySheet(sourcePath, referenceList);

        //        // Step 3: Optionally log or print the result
        //        if (!string.IsNullOrEmpty(updatedFilePath))
        //            Console.WriteLine($"✅ Excel updated successfully: {updatedFilePath}");
        //        else
        //            Console.WriteLine("⚠️ Excel update failed.");

        //        return updatedFilePath;
        //    }

        private static string TestCompoundInformation(string sourcePath)
        {
            try
            {
                // Dummy test data
                int txtReferenceStandard = 4;
                int txtImpurityMixture = 4;
                int txtDrugSubstance = 4;
                int txtSDD = 4;
                int txtDrugProduct = 4;
                int txtPlacebo = 4;
                int txtPolymer = 4;
                int txtIndividualImpurity = 4;
                int txtVolatiles = 4;
                int txtimpurity1 = 5;

                string cmbProtocolType = "PAV";
                string cmbProductType = "Drug ";
                string cmbTestType = "Dropdown ";

                // Call your real function
                string resultPath = CompoundInformation.UpdateCompoundInformationSheet2(
                    sourcePath,
                    txtReferenceStandard,
                    txtImpurityMixture,
                    txtDrugSubstance,
                    txtSDD,
                    txtDrugProduct,
                    txtPlacebo,
                    txtPolymer,
                    txtIndividualImpurity,
                    txtVolatiles,
                    cmbProtocolType,
                    cmbProductType,
                    cmbTestType, txtimpurity1);

                //Logger.LogMessage($"✅ TestCompoundInformation completed. Saved file path: {resultPath}", Level.Info);

                return resultPath;
            }
            catch (Exception ex)
            {
                //Logger.LogMessage($"❌ Error in TestCompoundInformation: {ex.Message}\r\n{ex.StackTrace}", Level.Error);
                return "";
            }
        }


    }
}
