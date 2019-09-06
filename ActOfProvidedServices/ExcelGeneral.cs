using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ActOfProvidedServices {
	class ExcelGeneral {
		private readonly BackgroundWorker bw;
		public static string AssemblyDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\";

		public ExcelGeneral(BackgroundWorker bw) {
			this.bw = bw;
		}

		protected bool OpenWorkbook(string workbook, out Excel.Application xlApp, out Excel.Workbook wb, out Excel.Worksheet ws, string sheetName = "") {
			xlApp = null;
			wb = null;
			ws = null;

			xlApp = new Excel.Application();

			if (xlApp == null) {
				bw.ReportProgress(0, "Не удалось открыть приложение Excel");
				return false;
			}

			xlApp.Visible = false;

			wb = xlApp.Workbooks.Open(workbook);

			if (wb == null) {
				bw.ReportProgress(0, "Не удалось открыть книгу " + workbook);
				return false;
			}

			if (string.IsNullOrEmpty(sheetName))
				sheetName = "Данные";

			ws = wb.Sheets[sheetName];

			if (ws == null) {
				bw.ReportProgress(0, "Не удалось открыть лист Данные");
				return false;
			}

			return true;
		}

		protected static void SaveAndCloseWorkbook(Excel.Application xlApp, Excel.Workbook wb, Excel.Worksheet ws) {
			if (ws != null) {
				Marshal.ReleaseComObject(ws);
				ws = null;
			}

			if (wb != null) {
				wb.Save();
				wb.Close(0);
				Marshal.ReleaseComObject(wb);
				wb = null;
			}

			if (xlApp != null) {
				xlApp.Quit();
				Marshal.ReleaseComObject(xlApp);
				xlApp = null;
			}

			GC.Collect();
			GC.WaitForPendingFinalizers();
		}

		public DataTable ReadExcelFile(string fileName, string sheetName) {
			bw.ReportProgress(0, "Считывание книги: " + fileName + ", лист: " + sheetName);
			DataTable dataTable = new DataTable();

			if (!File.Exists(fileName))
				return dataTable;

			try {
				using (OleDbConnection conn = new OleDbConnection()) {
					conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Mode=Read;" +
						"Extended Properties='Excel 12.0 Xml;HDR=NO;'";

					using (OleDbCommand comm = new OleDbCommand()) {
						if (string.IsNullOrEmpty(sheetName)) {
							conn.Open();
							DataTable dtSchema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables,
								new object[] { null, null, null, "TABLE" });
							sheetName = dtSchema.Rows[0].Field<string>("TABLE_NAME");
							conn.Close();
						} else
							sheetName += "$";

						comm.CommandText = "Select * from [" + sheetName + "]";
						comm.Connection = conn;

						using (OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter()) {
							oleDbDataAdapter.SelectCommand = comm;
							oleDbDataAdapter.Fill(dataTable);
						}
					}
				}
			} catch (Exception e) {
				bw.ReportProgress(0, e.Message + Environment.NewLine + e.StackTrace);
			}

			return dataTable;
		}

		private static string GetExcelColumnName(int columnNumber) {
			int dividend = columnNumber;
			string columnName = String.Empty;
			int modulo;

			while (dividend > 0) {
				modulo = (dividend - 1) % 26;
				columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
				dividend = (int)((dividend - modulo) / 26);
			}

			return columnName;
		}

		public static string ColumnIndexToColumnLetter(int colIndex) {
			int div = colIndex;
			string colLetter = String.Empty;
			int mod = 0;

			while (div > 0) {
				mod = (div - 1) % 26;
				colLetter = (char)(65 + mod) + colLetter;
				div = (int)((div - mod) / 26);
			}

			return colLetter;
		}

		protected void AddBoldBorder(Excel.Range range) {
			try {
				//foreach (Excel.XlBordersIndex item in new Excel.XlBordersIndex[] {
				//	Excel.XlBordersIndex.xlDiagonalDown,
				//	Excel.XlBordersIndex.xlDiagonalUp,
				//	Excel.XlBordersIndex.xlInsideHorizontal,
				//	Excel.XlBordersIndex.xlInsideVertical}) 
				//	range.Borders[item].LineStyle = Excel.Constants.xlNone;

				foreach (Excel.XlBordersIndex item in new Excel.XlBordersIndex[] {
					Excel.XlBordersIndex.xlEdgeBottom,
					Excel.XlBordersIndex.xlEdgeLeft,
					Excel.XlBordersIndex.xlEdgeRight,
					Excel.XlBordersIndex.xlEdgeTop}) {
					range.Borders[item].LineStyle = Excel.XlLineStyle.xlContinuous;
					range.Borders[item].ColorIndex = 0;
					range.Borders[item].TintAndShade = 0;
					range.Borders[item].Weight = Excel.XlBorderWeight.xlMedium;
				}
			} catch (Exception e) {
				bw.ReportProgress(0, e.Message + Environment.NewLine + e.StackTrace);
			}
		}

		public string WritePatientsToExcel(List<ItemPatient> patients, string resultFilePrefix, string templateFileName,
			string sheetName = "", bool createNew = true) {
			IWorkbook workbook = null;
			ISheet sheet = null;
			string resultFile = string.Empty;

			if (createNew) {
				if (!CreateNewIWorkbook(resultFilePrefix, templateFileName,
					out workbook, out sheet, out resultFile, sheetName))
					return string.Empty;
			} else {
				try {
					using (FileStream stream = new FileStream(templateFileName, FileMode.Open, FileAccess.Read))
						workbook = new XSSFWorkbook(stream);

					sheet = workbook.GetSheet(sheetName);
					resultFile = templateFileName;
				} catch (Exception e) {
					bw.ReportProgress(0, e.Message + Environment.NewLine + e.StackTrace);
					return string.Empty;
				}
			}

			int rowNumber = 16;
			IFont fontBold9 = workbook.CreateFont();
			IFont fontBold8 = workbook.CreateFont();
			IFont fontNormal8 = workbook.CreateFont();
			fontBold9.FontName = "Calibri";
			fontBold8.FontName = "Calibri";
			fontNormal8.FontName = "Calibri";
			fontBold9.FontHeightInPoints = 9;
			fontBold8.FontHeightInPoints = 8;
			fontNormal8.FontHeightInPoints = 8;
			fontBold9.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;
			fontBold8.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;
			fontNormal8.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Normal;
			ICellStyle cellStyleBold9 = workbook.CreateCellStyle();
			ICellStyle cellStyleBold8 = workbook.CreateCellStyle();
			ICellStyle cellStyleNormal8 = workbook.CreateCellStyle();
			ICellStyle cellStyleBold9BorderLeft = workbook.CreateCellStyle();
			ICellStyle cellStyleBold9BorderBottomTop = workbook.CreateCellStyle();
			ICellStyle cellStyleBold9BorderRight = workbook.CreateCellStyle();

			cellStyleBold9BorderLeft.BorderLeft = BorderStyle.Thin;
			cellStyleBold9BorderLeft.BorderRight = BorderStyle.None;
			cellStyleBold9BorderLeft.BorderTop = BorderStyle.Thin;
			cellStyleBold9BorderLeft.BorderBottom = BorderStyle.Thin;

			cellStyleBold9BorderRight.BorderLeft = BorderStyle.None;
			cellStyleBold9BorderRight.BorderRight = BorderStyle.Thin;
			cellStyleBold9BorderRight.BorderTop = BorderStyle.Thin;
			cellStyleBold9BorderRight.BorderBottom = BorderStyle.Thin;

			cellStyleBold9BorderBottomTop.BorderLeft = BorderStyle.None;
			cellStyleBold9BorderBottomTop.BorderRight = BorderStyle.None;
			cellStyleBold9BorderBottomTop.BorderTop = BorderStyle.Thin;
			cellStyleBold9BorderBottomTop.BorderBottom = BorderStyle.Thin;

			foreach (ItemPatient patient in patients) {
				double patientCostTotal = 0;
				WriteArrayToRow(sheet, ref rowNumber, new object[] { patient.Name, "", "", patient.Documents }, fontBold9, cellStyleBold9);

				foreach (ItemTreatment treatment in patient.Treatments) {
					WriteArrayToRow(sheet, ref rowNumber, 
						new object[] { "", treatment.Doctor, treatment.Date, "", "", treatment.TreatmentCostTotal }, fontBold8, cellStyleBold8);

					foreach (ItemService service in treatment.Services) {
						WriteArrayToRow(sheet, ref rowNumber, 
							new object[] { "", string.Join(Environment.NewLine, treatment.Diagnoses), 
							service.Code, service.Name, service.Count, service.Cost }, fontNormal8, cellStyleNormal8, wrapText:true);
						patientCostTotal += service.Count * service.Cost;
					}
				}

				WriteArrayToRow(sheet,
					ref rowNumber,
					new object[] { "Итого по пациенту:", "", "", "", "", patientCostTotal },
					fontBold9,
					cellStyleBold9BorderLeft,
					true,
					cellStyleBold9BorderBottomTop,
					cellStyleBold9BorderRight);
			}

			if (!SaveAndCloseIWorkbook(workbook, resultFile))
				return string.Empty;

			return resultFile;
		}

		private void WriteArrayToRow(ISheet sheet,
							   ref int rowNumber,
							   object[] values,
							   IFont font,
							   ICellStyle cellStyle,
							   bool createBorder = false,
							   ICellStyle cellStyleMedium = null,
							   ICellStyle cellStyleLast = null,
							   bool wrapText = false) {
			IRow row = null;
			try { row = sheet.GetRow(rowNumber); } catch (Exception) { }

			if (row == null)
				row = sheet.CreateRow(rowNumber);

			int columnNumber = 0;

			for (int i = 0; i < values.Length; i++) {
				object value = values[i];
			    ICell cell = null;
				try { cell = row.GetCell(columnNumber); } catch (Exception) { }

				if (cell == null)
					cell = row.CreateCell(columnNumber);
				
				if (createBorder) {
					if (i == 0)
						cell.CellStyle = cellStyle;
					if (i == values.Length - 1)
						cell.CellStyle = cellStyleLast;
					else
						cell.CellStyle = cellStyleMedium;
				} else
					cell.CellStyle = cellStyle;

				cell.CellStyle.SetFont(font);

				if (wrapText) {
					cell.CellStyle.WrapText = true;
					cell.CellStyle.VerticalAlignment = VerticalAlignment.Center;
				}

				if (value is double) {
					cell.SetCellValue((double)value);
				} else if (value is DateTime) {
					cell.SetCellValue((DateTime)value);
				} else {
					cell.SetCellValue(value.ToString());
				}

				columnNumber++;
			}

			rowNumber++;
		}


		private bool CreateNewIWorkbook(string resultFilePrefix, string templateFileName,
			out IWorkbook workbook, out ISheet sheet, out string resultFile, string sheetName) {
			workbook = null;
			sheet = null;
			resultFile = string.Empty;

			try {
				if (!GetTemplateFilePath(ref templateFileName))
					return false;

				string resultPath = GetResultFilePath(resultFilePrefix, templateFileName);

				using (FileStream stream = new FileStream(templateFileName, FileMode.Open, FileAccess.Read))
					workbook = new XSSFWorkbook(stream);

				if (string.IsNullOrEmpty(sheetName))
					sheetName = "Данные";

				sheet = workbook.GetSheet(sheetName);
				resultFile = resultPath;

				return true;
			} catch (Exception e) {
				bw.ReportProgress(0, e.Message + Environment.NewLine + e.StackTrace);
				return false;
			}
		}

		protected bool GetTemplateFilePath(ref string templateFileName) {
			templateFileName = Path.Combine(Path.Combine(AssemblyDirectory, "Templates\\"), templateFileName);

			if (!File.Exists(templateFileName)) {
				bw.ReportProgress(0, "Не удалось найти файл шаблона: " + templateFileName);
				return false;
			}

			return true;
		}

		public static string GetResultFilePath(string resultFilePrefix, string templateFileName = "", bool isPlainText = false) {
			string resultPath = Path.Combine(AssemblyDirectory, "Results");
			if (!Directory.Exists(resultPath))
				Directory.CreateDirectory(resultPath);

			foreach (char item in Path.GetInvalidFileNameChars())
				resultFilePrefix = resultFilePrefix.Replace(item, '-');

			string fileEnding = ".xlsx";
			if (isPlainText)
				fileEnding = ".txt";

			string resultFile = Path.Combine(resultPath, resultFilePrefix + " " + DateTime.Now.ToString("yyyyMMdd_HHmmss") + fileEnding);

			if (isPlainText && !string.IsNullOrEmpty(templateFileName))
				File.Copy(templateFileName, resultFile, true);

			return resultFile;
		}

		protected bool SaveAndCloseIWorkbook(IWorkbook workbook, string resultFile) {
			try {
				using (FileStream stream = new FileStream(resultFile, FileMode.Create, FileAccess.Write))
					workbook.Write(stream);

				workbook.Close();

				return true;
			} catch (Exception e) {
				bw.ReportProgress(0, e.Message + Environment.NewLine + e.StackTrace);
				return false;
			}
		}

		public bool Process(string resultFile,
					  double progressCurrent,
					  string organization,
					  string period,
					  string contract,
					  string dateDischarged,
					  List<ItemPatient> patients) {
			if (!OpenWorkbook(resultFile, out Excel.Application xlApp, out Excel.Workbook wb,
				out Excel.Worksheet ws))
				return false;

			try {
				int usedRows = ws.UsedRange.Rows.Count;

				ws.Range["A3"].Value2 = "выписан " + dateDischarged;
				ws.Range["A5"].Value2 = "Для организации: " + organization;
				ws.Range["A6"].Value2 = "Договор №: " + contract;
				ws.Range["A8"].Value2 = "За период: " + period;

				wb.Sheets["Итог"].Activate();
				wb.ActiveSheet.Range["A1:F13"].Select();
				xlApp.Selection.Copy();
				wb.Sheets["Данные"].Activate();
				ws.Range["A" + (usedRows + 1)].Select();
				ws.Paste();
				xlApp.DisplayAlerts = false;
				wb.Sheets["Итог"].Delete();
				xlApp.DisplayAlerts = true;

				double totalCost = 0;
				double totalCostWithDiscount = 0;
				double totalServices = 0;

				foreach (ItemPatient patient in patients) 
					foreach (ItemTreatment treatment in patient.Treatments) 
						foreach (ItemService service in treatment.Services) {
							totalCost += service.Count * service.Cost;
							totalCostWithDiscount += service.Count * service.CostFinal;
							totalServices += service.Count;
						}

				ws.Range["A" + (usedRows + 8)].EntireRow.RowHeight = 30;
				ws.Range["F" + (usedRows + 1)].Value2 = totalCost;
				ws.Range["F" + (usedRows + 2)].Value2 = totalCost - totalCostWithDiscount;
				ws.Range["C" + (usedRows + 3)].Value2 = totalCostWithDiscount;
				string cost = Slepov.Russian.СуммаПрописью.Сумма.Пропись(totalCostWithDiscount, Slepov.Russian.СуммаПрописью.Валюта.Рубли);
				ws.Range["C" + (usedRows + 4)].Value2 = "''" + cost.Substring(0, 1).ToUpper() + cost.Substring(1, cost.Length - 1) + "'";
				ws.Range["C" + (usedRows + 4) + ":F" + (usedRows + 4)].MergeCells = true;
				ws.Range["C" + (usedRows + 4)].WrapText = true;
				ws.Range["A" + (usedRows + 4)].EntireRow.RowHeight = 30;
				ws.Range["C" + (usedRows + 5)].Value2 = patients.Count;
				ws.Range["C" + (usedRows + 6)].Value2 = totalServices;

				ws.Range["A1"].Select();
			} catch (Exception e) {
				bw.ReportProgress(0, e.Message + Environment.NewLine + e.StackTrace);
			}

			SaveAndCloseWorkbook(xlApp, wb, ws);

			return true;
		}

	}
}
