using log4net.Core;
using Microsoft.Office.Interop.Excel;
using Internal.Framework.Collections;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace Spreadsheet.Handler
{
	public class InjectionRepeatability
	{
		private static Application _app;

		private const int DefaultNumInjections = 6;

		private const int MinNumInjections = 2;

		private const int DefaultNumDataTables = 1;

		private const int RawDataColOffset = 5;

		private const string TempDirectoryName = "ABD_TempFiles";

        public static string UpdateInjectionRepeatabilitySheet(
			string sourcePath,
			int numInjections,
			bool isAssayLevel,
			string signAssayRSD,
			decimal valueAssayRSD,
			int numPeaksAssay,
			bool isImpurityLevel,
			string signImpurityRSD,
			decimal valueImpurityRSD,
			int numPeaksImpurity,
			string cmbProtocolType,
			string cmbProductType,
			string cmbTestType)
        {
            string result = "";
			try
			{
				result = UpdateInjectionRepeatabilitySheet2(sourcePath, numInjections, isAssayLevel, signAssayRSD, valueAssayRSD, numPeaksAssay, isImpurityLevel, signImpurityRSD, valueImpurityRSD, numPeaksImpurity, cmbProtocolType, cmbProductType, cmbTestType);
			}
			catch (Exception ex)
			{
				Logger.LogMessage("An error occurred in the call to InjectionRepeatability.UpdateInjectionRepeatabilitySheet. Message and stack trace are:\r\n" + ex.Message + "\r\n" + ex.StackTrace, Level.Error);
				try
				{
					if (_app.Workbooks.Count > 0)
					{
						try
						{
							_app.Workbooks[0].Save();
							result = _app.Workbooks[0].FullName;
						}
						catch
						{
							Logger.LogMessage("An error occurred in the call to InjectionRepeatability.UpdateInjectionRepeatabilitySheet. Failed to save current workbook changes and to get path.", Level.Error);
						}
						_app.Workbooks.Close();
					}
					_app = null;
				}
				catch
				{
					Logger.LogMessage("An error occurred in the call to InjectionRepeatability.UpdateInjectionRepeatabilitySheet. Application failed to close workbooks. Message and stack trace are:\r\n" + ex.Message + "\r\n" + ex.StackTrace, Level.Error);
				}
				finally
				{
					WorksheetUtilities.ReleaseExcelApp();
				}
			}
			return result;
		}

        private static string UpdateInjectionRepeatabilitySheet2(
			string sourcePath,
			int numInjections,
			bool isAssayLevel,
			string signAssayRSD,
			decimal valueAssayRSD,
			int numPeaksAssay,
			bool isImpurityLevel,
			string signImpurityRSD,
			decimal valueImpurityRSD,
			int numPeaksImpurity,
			string strcmbProtocolType,
			string strcmbProductType,
			string strcmbTestType)
        {
            if (!isAssayLevel && !isImpurityLevel)
			{
				Logger.LogMessage("Error in call to InjectionRepeatability.UpdateInjectionRepeatabilitySheet. Should be either of Assay Level or Impurity Level!", Level.Error);
				return "";
			}

			if (!File.Exists(sourcePath))
			{
				Logger.LogMessage("Error in call to InjectionRepeatability.UpdateInjectionRepeatabilitySheet. Invalid source file path specified.", Level.Error);
				return "";
			}
			string text = WorksheetUtilities.CopyWorkbook(sourcePath, "ABD_TempFiles", "Injection Repeatability Results.xls");
			if (string.IsNullOrEmpty(text))
			{
				return "";
			}
			_app = WorksheetUtilities.GetExcelApp();
			_app.Workbooks.Open(text, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
			Workbook workbook = _app.Workbooks[1];
			if (workbook.Worksheets[1] is Worksheet worksheet)
			{
				bool wasProtected = WorksheetUtilities.SetSheetProtection(worksheet, null, protect: false);

                WorksheetUtilities.SetMetadataValues(worksheet, strcmbProtocolType, strcmbProductType, strcmbTestType);

                if (isAssayLevel)
				{
					SetSignAndRSD(worksheet, signAssayRSD, valueAssayRSD, true);
					numInjections = AdjustRows(worksheet, numInjections, true);

					CreateReplicatesAndAdjustFormulas(worksheet, numPeaksAssay, numInjections, true);
				}
				else
				{
					// delete Assay level section
					DeleteLevelSection(worksheet, true);
				}

				if (isImpurityLevel)
				{
					SetSignAndRSD(worksheet, signImpurityRSD, valueImpurityRSD, false);
					numInjections = AdjustRows(worksheet, numInjections, false);

					CreateReplicatesAndAdjustFormulas(worksheet, numPeaksImpurity, numInjections, false);
				}
				else
				{
					// delete Impurity Level section
					DeleteLevelSection(worksheet, false);
				}

				WorksheetUtilities.PostProcessSheet(worksheet);
			}

			_app.Workbooks[1].Save();
			WorksheetUtilities.ReleaseComObject(workbook);
			_app.Workbooks.Close();
			_app = null;
			WorksheetUtilities.ReleaseExcelApp();
			return text;
		}

		private static void DeleteLevelSection(Worksheet sheet, bool isAssay)
		{
			WorksheetUtilities.DeleteRowsFromNamedRange(isAssay ? 41 : 40, sheet, AddPrefix("LevelSection", isAssay), XlDirection.xlDown);
		}

		private static void SetSignAndRSD(Worksheet sheet, string sign, decimal valueRSD, bool isAssay)
		{
			WorksheetUtilities.SetNamedRangeValue(sheet, AddPrefix("Sign", isAssay), sign, 1, 1);
			WorksheetUtilities.SetNamedRangeValue(sheet, AddPrefix("RSD", isAssay), valueRSD.ToString(), 1, 1);
		}

		private static int AdjustRows(Worksheet worksheet, int numInjections, bool isAssay)
		{
			if (numInjections > DefaultNumInjections)
			{
				int numRowsToInsert = numInjections - DefaultNumInjections;
				WorksheetUtilities.InsertRowsIntoNamedRange(numRowsToInsert, worksheet, AddPrefix("InjectNumsRawData1", isAssay), fillRows: false, XlDirection.xlDown, XlPasteType.xlPasteFormulas);
				WorksheetUtilities.InsertRowsIntoNamedRange(numRowsToInsert, worksheet, AddPrefix("InjectNumsValidationResults", isAssay), fillRows: true, XlDirection.xlDown, XlPasteType.xlPasteFormulas);
			}
			else if (numInjections < DefaultNumInjections)
			{
				int num = DefaultNumInjections - numInjections;
				if (DefaultNumInjections - num < MinNumInjections)
				{
					num = 4;
				}
				WorksheetUtilities.DeleteRowsFromNamedRange(num, worksheet, AddPrefix("InjectNumsRawData1", isAssay), XlDirection.xlDown);
				WorksheetUtilities.DeleteRowsFromNamedRange(num, worksheet, AddPrefix("InjectNumsValidationResults", isAssay), XlDirection.xlDown);
			}
			if (numInjections < MinNumInjections)
			{
				numInjections = 2;
			}
			List<string> list = new List<string>(0);
			for (int i = 1; i <= numInjections; i++)
			{
				list.Add(i.ToString());
			}
			WorksheetUtilities.SetNamedRangeValues(worksheet, AddPrefix("InjectNumsRawData1", isAssay), list);
			WorksheetUtilities.SetNamedRangeValues(worksheet, AddPrefix("InjectNumsValidationResults", isAssay), list);
			return numInjections;
		}

		private static void CreateReplicatesAndAdjustFormulas(Worksheet worksheet, int numPeaks, int numInjections, bool isAssay)
		{
			// replicate horizontally
			if (numPeaks > 1)
			{
				int num2 = numPeaks - 1;
				int colOffset = 5;
				for (int j = 1; j <= num2; j++)
				{
					int num3 = j + 1;
					WorksheetUtilities.CopyNamedRangeToNewNamedRange(worksheet, AddPrefix("RawDataTable1", isAssay), AddPrefix("RawDataTable" + num3, isAssay), 1, colOffset * j + 1, XlPasteType.xlPasteAll);
					WorksheetUtilities.SetNamedRangeValue(worksheet, AddPrefix("RawDataTable" + num3, isAssay), "Raw Data Table " + num3, 1, 1);
					WorksheetUtilities.CopyNamedRangeToNewNamedRange(worksheet, AddPrefix("RoundDecimals1", isAssay), AddPrefix("RoundDecimals" + num3, isAssay), 1, colOffset * j + 1, XlPasteType.xlPasteAll);
					WorksheetUtilities.CopyNamedRangeToNewNamedRange(worksheet, AddPrefix("Calculations1", isAssay), AddPrefix("Calculations" + num3, isAssay), 1, colOffset * j + 1, XlPasteType.xlPasteAll);
					WorksheetUtilities.CopyNamedRangeToNewNamedRange(worksheet, AddPrefix("ValidAreaCounts1", isAssay), AddPrefix("ValidAreaCounts" + num3, isAssay), 1, j + 1, XlPasteType.xlPasteAll);
					WorksheetUtilities.CopyNamedRangeToNewNamedRange(worksheet, AddPrefix("ValidationAreaCounts1", isAssay), AddPrefix("ValidationAreaCounts" + num3, isAssay), 1, j + 1, XlPasteType.xlPasteAll);
					WorksheetUtilities.CopyNamedRangeToNewNamedRange(worksheet, AddPrefix("ValidationsCalculations1", isAssay), AddPrefix("ValidationsCalculations" + num3, isAssay), 1, j + 1, XlPasteType.xlPasteAll);
					WorksheetUtilities.CopyNamedRangeToNewNamedRange(worksheet, AddPrefix("RawDataPeakName1", isAssay), AddPrefix("RawDataPeakName" + num3, isAssay), 1, colOffset * j + 1, XlPasteType.xlPasteAll);
					WorksheetUtilities.CopyNamedRangeToNewNamedRange(worksheet, AddPrefix("RawDataConc1", isAssay), AddPrefix("RawDataConc" + num3, isAssay), 1, colOffset * j + 1, XlPasteType.xlPasteAll);
					WorksheetUtilities.CopyNamedRangeToNewNamedRange(worksheet, AddPrefix("PeakName1", isAssay), AddPrefix("PeakName" + num3, isAssay), 1, colOffset * j + 1, XlPasteType.xlPasteAll);
					WorksheetUtilities.CopyNamedRangeToNewNamedRange(worksheet, AddPrefix("StrengthValue1", isAssay), AddPrefix("StrengthValue" + num3, isAssay), 1, j + 1, XlPasteType.xlPasteAll);
					WorksheetUtilities.CopyNamedRangeToNewNamedRange(worksheet, AddPrefix("ValidationTableHeader1", isAssay), AddPrefix("ValidationTableHeader" + num3, isAssay), 1, j + 1, XlPasteType.xlPasteAll);
					WorksheetUtilities.ResizeNamedRange(worksheet, AddPrefix("ValidationTable", isAssay), 0, 1);
				}
			}

			UpdateAreaCountTableFormulas_old(worksheet, numPeaks, numInjections, isAssay);
		}

		private static void UpdateAreaCountTableFormulas_old(_Worksheet sheet, int numPeaks, int numInjections, bool isAssay)
		{
			if (sheet == null || numPeaks <= 0 || numInjections <= 0)
			{
				return;
			}

			const int rawDataPrepsRow = 2;
			const int rawDataTableColIndex = 2;
			const int areaCountRow = 1;

			object objRawDataNamedRange = null;
			Name rawDataTableName = null;
			Range rawDataTableRange = null;
			object objAreaCountNamedRange = null;
			Name areaCountName = null;
			Range areaCountRange = null;
			Range srcCell = null;
			Range destCell = null;
			object objRoundDecimalNamedRange = null;

			//Decimal rounding in validation table
			Name roundDecimalName = null;
			Range roundDecimalRange = null;
			Range decimalSrcCell = null;
			string decimalSrcCellAddress = "";

			//Correct formulas in validation table
			object objCalculationsNamedRange = null;
			Name calculationsName = null;
			Range calculationsRange = null;
			object objValCalculationsNamedRange = null;
			Name valCalculationsName = null;
			Range valCalculationsRange = null;

			//Update formulas in rawdatatable
			object objPeakNamedRange = null;
			Name PeakName = null;
			Range PeakRange = null;
			string SrcCellAddress = "";
			object objPeakNameNamedRange = null;
			Name peakNameName = null;
			Range peakNameRange = null;

			//Updates formulas in validation header table
			object objValTableHeader = null;
			Name valTableHeaderName = null;
			Range valTableHeaderRange = null;
			for (int i = 1; i <= numPeaks; i++)
			{
				objRawDataNamedRange = sheet.Names.Item(AddPrefix("RawDataTable", isAssay) + i, Type.Missing, Type.Missing);
				if (!(objRawDataNamedRange is Name))
				{
					break;
				}
				rawDataTableName = objRawDataNamedRange as Name;
				rawDataTableRange = rawDataTableName.RefersToRange;
				objAreaCountNamedRange = sheet.Names.Item(AddPrefix("ValidationAreaCounts", isAssay) + i, Type.Missing, Type.Missing);
				if (!(objAreaCountNamedRange is Name))
				{
					continue;
				}
				areaCountName = objAreaCountNamedRange as Name;
				areaCountRange = areaCountName.RefersToRange;

                //Get decimals to round
                objRoundDecimalNamedRange = sheet.Names.Item(AddPrefix("RoundDecimals", isAssay) + i, Type.Missing, Type.Missing);
                if (!(objRoundDecimalNamedRange is Name))
                {
                    continue;
                }
                roundDecimalName = objRoundDecimalNamedRange as Name;
                roundDecimalRange = roundDecimalName.RefersToRange;
                decimalSrcCell = roundDecimalRange.Cells[1, 1] as Range;
                if (decimalSrcCell != null)
                {
                    decimalSrcCellAddress = decimalSrcCell.get_Address(Type.Missing, Type.Missing, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                }

                for (int j = 0; j < areaCountRange.Rows.Count; j++)
				{
					srcCell = rawDataTableRange.Cells[j + rawDataPrepsRow, rawDataTableColIndex] as Range;
					if (srcCell == null)
					{
						continue;
					}
					string srcCellAddress = srcCell.get_Address(Type.Missing, Type.Missing, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
					srcCellAddress = srcCellAddress.Replace("$", "");
					destCell = areaCountRange.Cells[j + areaCountRow, 1] as Range;
					if (destCell != null)
					{
						if (j < 3)
						{
							destCell.Value2 = string.Format("=IF({0}=\"\",\" \",{0})", srcCellAddress);
						}
						else
						{
							destCell.Value2 = string.Format("=IF({0}=\"\",\" \",ROUND({0},{1}))", srcCellAddress, decimalSrcCellAddress);
						}
						WorksheetUtilities.ReleaseComObject(destCell);
					}
					WorksheetUtilities.ReleaseComObject(srcCell);
				}

				objCalculationsNamedRange = sheet.Names.Item(AddPrefix("Calculations", isAssay) + i, Type.Missing, Type.Missing);
				if (!(objCalculationsNamedRange is Name))
				{
					continue;
				}
				calculationsName = objCalculationsNamedRange as Name;
				calculationsRange = calculationsName.RefersToRange;
				objValCalculationsNamedRange = sheet.Names.Item(AddPrefix("ValidationsCalculations", isAssay) + i, Type.Missing, Type.Missing);
				if (!(objValCalculationsNamedRange is Name))
				{
					continue;
				}
				valCalculationsName = objValCalculationsNamedRange as Name;
				valCalculationsRange = valCalculationsName.RefersToRange;
				for (int k = 0; k < calculationsRange.Rows.Count; k++)
				{
					srcCell = calculationsRange.Cells[k + 1, 1] as Range;
					if (srcCell == null)
					{
						continue;
					}
					string srcCellAddress = srcCell.get_Address(Type.Missing, Type.Missing, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
					srcCellAddress = srcCellAddress.Replace("$", "");
					destCell = valCalculationsRange.Cells[k + 1, 1] as Range;
					if (destCell != null)
					{
						destCell.Value2 = string.Format("=IF({0}=\"\",\" \",{0})", srcCellAddress);
						WorksheetUtilities.ReleaseComObject(destCell);
					}
					WorksheetUtilities.ReleaseComObject(srcCell);
				}
				objPeakNamedRange = sheet.Names.Item(AddPrefix("RawDataPeakName", isAssay) + i, Type.Missing, Type.Missing);
				if (!(objPeakNamedRange is Name))
				{
					continue;
				}
				PeakName = objPeakNamedRange as Name;
				PeakRange = PeakName.RefersToRange;
				objPeakNameNamedRange = sheet.Names.Item(AddPrefix("PeakName", isAssay) + i, Type.Missing, Type.Missing);
				if (!(objPeakNameNamedRange is Name))
				{
					continue;
				}
				peakNameName = objPeakNameNamedRange as Name;
				peakNameRange = peakNameName.RefersToRange;
				srcCell = PeakRange.Cells[1, 1] as Range;
				if (srcCell != null)
				{
					SrcCellAddress = srcCell.get_Address(Type.Missing, Type.Missing, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
				}
				for (int l = 0; l < peakNameRange.Rows.Count; l++)
				{
					SrcCellAddress = SrcCellAddress.Replace("$", "");
					destCell = peakNameRange.Cells[l + 1, 1] as Range;
					if (destCell != null)
					{
						destCell.Value2 = string.Format("=IF({0}=\"\",\" \",{0})", SrcCellAddress);
						WorksheetUtilities.ReleaseComObject(destCell);
					}
				}
				objValTableHeader = sheet.Names.Item(AddPrefix("ValidationTableHeader", isAssay) + i, Type.Missing, Type.Missing);
				if (!(objValTableHeader is Name))
				{
					continue;
				}
				valTableHeaderName = objValTableHeader as Name;
				valTableHeaderRange = valTableHeaderName.RefersToRange;
				for (int m = 0; m < valTableHeaderRange.Rows.Count; m++)
				{
					destCell = valTableHeaderRange.Cells[1 + m, 1] as Range;
					if (destCell != null)
					{
						if (m == 0)
						{
							destCell.Value2 = string.Format("=IF({0}=\"\",\" \",{0})", SrcCellAddress);
						}
						WorksheetUtilities.ReleaseComObject(destCell);
					}
				}
				try
				{
					WorksheetUtilities.ReleaseComObject(objRawDataNamedRange);
					WorksheetUtilities.ReleaseComObject(rawDataTableRange);
					WorksheetUtilities.ReleaseComObject(rawDataTableName);
					WorksheetUtilities.ReleaseComObject(objAreaCountNamedRange);
					WorksheetUtilities.ReleaseComObject(areaCountRange);
					WorksheetUtilities.ReleaseComObject(areaCountName);
					WorksheetUtilities.ReleaseComObject(roundDecimalRange);
					WorksheetUtilities.ReleaseComObject(roundDecimalName);
					WorksheetUtilities.ReleaseComObject(objRoundDecimalNamedRange);
					WorksheetUtilities.ReleaseComObject(decimalSrcCell);
					WorksheetUtilities.ReleaseComObject(objCalculationsNamedRange);
					WorksheetUtilities.ReleaseComObject(calculationsRange);
					WorksheetUtilities.ReleaseComObject(calculationsName);
					WorksheetUtilities.ReleaseComObject(objValCalculationsNamedRange);
					WorksheetUtilities.ReleaseComObject(valCalculationsRange);
					WorksheetUtilities.ReleaseComObject(valCalculationsName);
					WorksheetUtilities.ReleaseComObject(objPeakNamedRange);
					WorksheetUtilities.ReleaseComObject(PeakRange);
					WorksheetUtilities.ReleaseComObject(PeakName);
					WorksheetUtilities.ReleaseComObject(objPeakNameNamedRange);
					WorksheetUtilities.ReleaseComObject(peakNameRange);
					WorksheetUtilities.ReleaseComObject(peakNameName);
					WorksheetUtilities.ReleaseComObject(objValTableHeader);
					WorksheetUtilities.ReleaseComObject(valTableHeaderRange);
					WorksheetUtilities.ReleaseComObject(valTableHeaderName);
				}
				catch
				{
				}
			}
			try
			{
				WorksheetUtilities.ReleaseComObject(objRawDataNamedRange);
				WorksheetUtilities.ReleaseComObject(rawDataTableName);
				WorksheetUtilities.ReleaseComObject(rawDataTableRange);
				WorksheetUtilities.ReleaseComObject(objAreaCountNamedRange);
				WorksheetUtilities.ReleaseComObject(areaCountName);
				WorksheetUtilities.ReleaseComObject(areaCountRange);
				WorksheetUtilities.ReleaseComObject(srcCell);
				WorksheetUtilities.ReleaseComObject(destCell);

				sheet = null;
			}
			catch
			{
			}
		}

		private static string AddPrefix(string input, bool isAssay)
		{
			string prefix = isAssay ? "Assay" : "Impurity";
			return $"{prefix}{input}";
		}
	}
}