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
						"Extended Properties='Excel 12.0 Xml;HDR=NO;IMEX=1'";

					using (OleDbCommand comm = new OleDbCommand()) {
						if (string.IsNullOrEmpty(sheetName)) {
							conn.Open();
							DataTable dtSchema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables,
								new object[] { null, null, null, "TABLE" });
							sheetName = dtSchema.Rows[0].Field<string>("TABLE_NAME");
							conn.Close();
						} else
							sheetName += "$";

#pragma warning disable CA2100 // Review SQL queries for security vulnerabilities
						comm.CommandText = "Select * from [" + sheetName + "]";
#pragma warning restore CA2100 // Review SQL queries for security vulnerabilities
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



		public string WritePatientsToExcelRenessans(List<ItemPatient> patients, MainModel.Type type, string resultFilePrefix, string templateFileName,
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

			int rowNumber = 0;
			switch (type) {
				case MainModel.Type.Renessans:
					rowNumber = 16;
					break;
				case MainModel.Type.VTB:
					rowNumber = 19;
					break;
				case MainModel.Type.Rosgosstrakh:
					rowNumber = 19;
					break;
				case MainModel.Type.Reso:
					rowNumber = 19;
					break;
				default:
					break;
			}

			#region styles
			IFont fontBold10 = workbook.CreateFont();
			fontBold10.FontName = "Calibri";
			fontBold10.FontHeightInPoints = 10;
			fontBold10.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;

			IFont fontBold9 = workbook.CreateFont();
			fontBold9.FontName = "Calibri";
			fontBold9.FontHeightInPoints = 9;
			fontBold9.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;

			IFont fontBold8 = workbook.CreateFont();
			fontBold8.FontName = "Calibri";
			fontBold8.FontHeightInPoints = 8;
			fontBold8.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;

			IFont fontNormal8 = workbook.CreateFont();
			fontNormal8.FontName = "Calibri";
			fontNormal8.FontHeightInPoints = 8;
			fontNormal8.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Normal;
			
			IFont fontNormal9 = workbook.CreateFont();
			fontNormal9.FontName = "Calibri";
			fontNormal9.FontHeightInPoints = 9;
			fontNormal9.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Normal;

			ICellStyle cellStyleBold10 = workbook.CreateCellStyle();
			cellStyleBold10.SetFont(fontBold10);

			ICellStyle cellStyleBold9 = workbook.CreateCellStyle();
			cellStyleBold9.SetFont(fontBold9);

			ICellStyle cellStyleBold8 = workbook.CreateCellStyle();
			cellStyleBold8.SetFont(fontBold8);

			ICellStyle cellStyleNormal8 = workbook.CreateCellStyle();
			cellStyleNormal8.SetFont(fontNormal8);

			ICellStyle cellStyleNormal9 = workbook.CreateCellStyle();
			cellStyleNormal9.SetFont(fontNormal9);

			ICellStyle cellStyleBold9BorderLeft = workbook.CreateCellStyle();
			cellStyleBold9BorderLeft.CloneStyleFrom(cellStyleBold9);
			cellStyleBold9BorderLeft.BorderLeft = BorderStyle.Thin;
			cellStyleBold9BorderLeft.BorderRight = BorderStyle.None;
			cellStyleBold9BorderLeft.BorderTop = BorderStyle.Thin;
			cellStyleBold9BorderLeft.BorderBottom = BorderStyle.Thin;

			ICellStyle cellStyleBold9BorderRight = workbook.CreateCellStyle();
			cellStyleBold9BorderRight.CloneStyleFrom(cellStyleBold9);
			cellStyleBold9BorderRight.BorderLeft = BorderStyle.None;
			cellStyleBold9BorderRight.BorderRight = BorderStyle.Thin;
			cellStyleBold9BorderRight.BorderTop = BorderStyle.Thin;
			cellStyleBold9BorderRight.BorderBottom = BorderStyle.Thin;

			ICellStyle cellStyleBold9BorderBottomTop = workbook.CreateCellStyle();
			cellStyleBold9BorderBottomTop.CloneStyleFrom(cellStyleBold9);
			cellStyleBold9BorderBottomTop.BorderLeft = BorderStyle.None;
			cellStyleBold9BorderBottomTop.BorderRight = BorderStyle.None;
			cellStyleBold9BorderBottomTop.BorderTop = BorderStyle.Thin;
			cellStyleBold9BorderBottomTop.BorderBottom = BorderStyle.Thin;

			ICellStyle cellStyleBold9Wrap = workbook.CreateCellStyle();
			cellStyleBold9Wrap.CloneStyleFrom(cellStyleBold9);
			cellStyleBold9Wrap.WrapText = true;
			cellStyleBold9Wrap.VerticalAlignment = VerticalAlignment.Center;

			ICellStyle cellStyleBold9Centered = workbook.CreateCellStyle();
			cellStyleBold9Centered.CloneStyleFrom(cellStyleBold9);
			cellStyleBold9Centered.Alignment = HorizontalAlignment.Center;

			ICellStyle cellStyleBold8Wrap = workbook.CreateCellStyle();
			cellStyleBold8Wrap.CloneStyleFrom(cellStyleBold8);
			cellStyleBold8Wrap.WrapText = true;
			cellStyleBold8Wrap.VerticalAlignment = VerticalAlignment.Center;

			ICellStyle cellStyleNormal9Wrap = workbook.CreateCellStyle();
			cellStyleNormal9Wrap.CloneStyleFrom(cellStyleNormal9);
			cellStyleNormal9Wrap.WrapText = true;
			cellStyleNormal9Wrap.VerticalAlignment = VerticalAlignment.Center;

			ICellStyle cellStyleNormal8Wrap = workbook.CreateCellStyle();
			cellStyleNormal8Wrap.CloneStyleFrom(cellStyleNormal8);
			cellStyleNormal8Wrap.WrapText = true;
			cellStyleNormal8Wrap.VerticalAlignment = VerticalAlignment.Center;

			ICellStyle cellStyleBold9BorderAll = workbook.CreateCellStyle();
			cellStyleBold9BorderAll.CloneStyleFrom(cellStyleBold9);
			cellStyleBold9BorderAll.BorderLeft = BorderStyle.Medium;
			cellStyleBold9BorderAll.BorderRight = BorderStyle.Medium;
			cellStyleBold9BorderAll.BorderTop = BorderStyle.Medium;
			cellStyleBold9BorderAll.BorderBottom = BorderStyle.Medium;

			if (type == MainModel.Type.VTB ||
				type == MainModel.Type.Rosgosstrakh) {
				cellStyleBold9.Alignment = HorizontalAlignment.Center;
				cellStyleBold9.VerticalAlignment = VerticalAlignment.Center;

				cellStyleBold10.Alignment = HorizontalAlignment.Center;
				cellStyleBold10.VerticalAlignment = VerticalAlignment.Center;

				cellStyleNormal9.Alignment = HorizontalAlignment.Center;
				cellStyleNormal9.VerticalAlignment = VerticalAlignment.Center;

				cellStyleNormal9Wrap.Alignment = HorizontalAlignment.Center;
				cellStyleNormal9Wrap.VerticalAlignment = VerticalAlignment.Center;
			}

			#endregion

			foreach (ItemPatient patient in patients) {
				double patientCostTotal = 0;
				if (type == MainModel.Type.Renessans ||
					type == MainModel.Type.Reso)
					WriteArrayToRow(sheet, ref rowNumber, 
						new (object, ICellStyle)[] {
							(patient.Name, type == MainModel.Type.Renessans ? cellStyleBold9 : cellStyleBold10),
							("", null), 
							("", null),
							(patient.Documents.Replace("@number", patient.GarantyMail), cellStyleBold9) });

				foreach (ItemTreatment treatment in patient.Treatments) {
					if (type == MainModel.Type.Renessans)
						WriteArrayToRow(sheet, ref rowNumber,
							new (object, ICellStyle)[] { 
								("", null),
								(treatment.Doctor, cellStyleBold8), 
								(treatment.Date, cellStyleBold8), 
								("", null),
								("", null),
								(treatment.TreatmentCostTotal, cellStyleBold8) });

					foreach (ItemService service in treatment.Services) {
						if (type == MainModel.Type.Renessans) {
							WriteArrayToRow(sheet, ref rowNumber,
								new (object, ICellStyle)[] {
									(service.ToothNumber, cellStyleNormal8Wrap),
									(string.Join(Environment.NewLine, treatment.Diagnoses), cellStyleBold8Wrap),
									(service.Code, cellStyleNormal8Wrap),
									(service.Name, cellStyleNormal8Wrap),
									(service.Count, cellStyleNormal8Wrap),
									(service.Cost * service.Count, cellStyleNormal8Wrap) });
							patientCostTotal += service.Count * service.Cost;

						} else if (type == MainModel.Type.Reso) {
							WriteArrayToRow(sheet, ref rowNumber,
								new (object, ICellStyle)[] {
									("", null),
									(treatment.Date, cellStyleBold9Centered),
									(treatment.Doctor, cellStyleBold9),
									("", null),
									("", null),
									("", null),
									(treatment.Filial, cellStyleNormal9)});

							WriteArrayToRow(sheet, ref rowNumber,
								new (object, ICellStyle)[] {
									(service.ToothNumber, cellStyleNormal9Wrap),
									(string.Join(";", treatment.Diagnoses), cellStyleNormal9),
									(service.Code, cellStyleNormal9),
									(service.Name, cellStyleNormal9Wrap),
									(service.Count, cellStyleNormal9),
									(service.Cost * service.Count, cellStyleBold9) });
							patientCostTotal += service.Count * service.Cost;

						} else if (type == MainModel.Type.VTB ||
							type == MainModel.Type.Rosgosstrakh) {
							WriteArrayToRow(sheet, ref rowNumber, new (object, ICellStyle)[] {
								(patient.Documents, cellStyleBold9),
								(treatment.GarantyMail, cellStyleBold9),
								(patient.Name, cellStyleBold10),
								(treatment.Date, cellStyleBold9),
								(string.Join(";", treatment.Diagnoses), cellStyleNormal9),
								(service.ToothNumber, cellStyleNormal9),
								(service.Code, cellStyleNormal9),
								(service.Name, cellStyleNormal9Wrap),
								(treatment.Doctor, cellStyleBold9),
								(service.Count, cellStyleNormal9),
								(service.Cost, cellStyleBold9),
								(service.Count * service.Cost, cellStyleBold9)
							});
						}
					}
				}

				if (type == MainModel.Type.Renessans)
					WriteArrayToRow(sheet,
						ref rowNumber,
						new (object, ICellStyle)[] {
							("Итого по пациенту:", cellStyleBold9BorderLeft), 
						    ("", cellStyleBold9BorderBottomTop), 
						    ("", cellStyleBold9BorderBottomTop), 
						    ("", cellStyleBold9BorderBottomTop), 
						    ("", cellStyleBold9BorderBottomTop),
						    (patientCostTotal, cellStyleBold9BorderRight) });

				else if (type == MainModel.Type.Reso)
					WriteArrayToRow(sheet,
						ref rowNumber,
						new (object, ICellStyle)[] {
							("Итого по клиенту:", cellStyleBold9BorderAll),
							("", cellStyleBold9BorderAll),
							("", cellStyleBold9BorderAll),
							("", cellStyleBold9BorderAll),
							("", cellStyleBold9BorderAll),
							(patientCostTotal, cellStyleBold9BorderAll),
							("", cellStyleBold9BorderAll) });
			}

			if (!SaveAndCloseIWorkbook(workbook, resultFile))
				return string.Empty;

			return resultFile;
		}

		public bool WriteEndingRenessans(MainModel.Type type, string resultFile, double progressCurrent,
			string period, string contract, string dateDischarged, List<ItemPatient> patients) {
			if (!OpenWorkbook(resultFile, out Excel.Application xlApp, out Excel.Workbook wb,
				out Excel.Worksheet ws))
				return false;

			try {
				int usedRows = ws.UsedRange.Rows.Count;

				ws.Range["A3"].Value2 = "выписан " + dateDischarged;

				if (type == MainModel.Type.Renessans)
					ws.Range["A6"].Value2 = "Договор №: " + contract;
				else if (type == MainModel.Type.VTB ||
					type == MainModel.Type.Rosgosstrakh)
					ws.Range["C6"].Value2 = contract;
				else if (type == MainModel.Type.Reso)
					ws.Range["B6"].Value2 = contract;

				ws.Range["A8"].Value2 = "За период: " + period;

				wb.Sheets["Итог"].Activate();
				string rangeToCopy = "A1:F13";
				if (type == MainModel.Type.VTB ||
					type == MainModel.Type.Rosgosstrakh ||
					type == MainModel.Type.Reso)
					rangeToCopy = "A1:G15";

				wb.ActiveSheet.Range[rangeToCopy].Select();
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

				if (type == MainModel.Type.Renessans) {
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
				} else if (type == MainModel.Type.VTB ||
					type == MainModel.Type.Rosgosstrakh) {
					ws.Range["D" + (usedRows + 2)].Value2 = string.Format("{0:0.##} руб", totalCost);
					ws.Range["D" + (usedRows + 3)].Value2 = string.Format("{0:0.##} руб", totalCost - totalCostWithDiscount);
					ws.Range["D" + (usedRows + 4)].Value2 = string.Format("{0:0.##} руб", totalCostWithDiscount);
					string cost = Slepov.Russian.СуммаПрописью.Сумма.Пропись(
						double.Parse(string.Format("{0:0.##}", totalCostWithDiscount)), Slepov.Russian.СуммаПрописью.Валюта.Рубли);
					cost = cost.Substring(0, 1).ToUpper() + cost.Substring(1, cost.Length - 1);
					ws.Range["D" + (usedRows + 5)].Value2 = cost;
					ws.Range["D" + (usedRows + 6)].Value2 = "'" + patients.Count;
					ws.Range["D" + (usedRows + 7)].Value2 = "'" + totalServices;
				} else if (type == MainModel.Type.Reso) {
					ws.Range["C" + (usedRows + 2)].Value2 = string.Format("{0:0.##} руб", totalCost);
					ws.Range["C" + (usedRows + 3)].Value2 = string.Format("{0:0.##} руб", totalCost - totalCostWithDiscount);
					ws.Range["C" + (usedRows + 4)].Value2 = string.Format("{0:0.##} руб", totalCostWithDiscount);
					string cost = Slepov.Russian.СуммаПрописью.Сумма.Пропись(
						double.Parse(string.Format("{0:0.##}", totalCostWithDiscount)), Slepov.Russian.СуммаПрописью.Валюта.Рубли);
					cost = cost.Substring(0, 1).ToUpper() + cost.Substring(1, cost.Length - 1);
					ws.Range["C" + (usedRows + 5)].Value2 = cost;
					ws.Range["C" + (usedRows + 6)].Value2 = "'" + patients.Count;
					ws.Range["C" + (usedRows + 7)].Value2 = "'" + totalServices;
				}

				ws.Range["A1"].Select();
			} catch (Exception e) {
				bw.ReportProgress(0, e.Message + Environment.NewLine + e.StackTrace);
			}

			SaveAndCloseWorkbook(xlApp, wb, ws);

			return true;
		}


		private void WriteArrayToRow(ISheet sheet, ref int rowNumber, (object, ICellStyle)[] valuesStyles) {
			IRow row = null;
			try { row = sheet.GetRow(rowNumber); } catch (Exception) { }

			if (row == null)
				row = sheet.CreateRow(rowNumber);

			int columnNumber = 0;

			for (int i = 0; i < valuesStyles.Length; i++) {
				(object, ICellStyle) valueStyle = valuesStyles[i];
				ICell cell = null;
				try { cell = row.GetCell(columnNumber); } catch (Exception) { }

				if (cell == null)
					cell = row.CreateCell(columnNumber);

				if (valueStyle.Item2 != null)
					cell.CellStyle = valueStyle.Item2;

				object value = valueStyle.Item1;
				if (value is double) {
					cell.SetCellValue((double)value);
				} else if (value is DateTime) {
					cell.SetCellValue((DateTime)value);
				} else {
					try {
						cell.SetCellValue(value.ToString());
					} catch (Exception e) {
						Console.WriteLine(e.Message + Environment.NewLine + e.StackTrace);
					}
				}

				columnNumber++;
			}

			rowNumber++;
		}


		public static List<string> ReadSheetNames(string file) {
			List<string> sheetNames = new List<string>();

			using (OleDbConnection conn = new OleDbConnection()) {
				conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + file + ";Mode=Read;" +
					"Extended Properties='Excel 12.0 Xml;HDR=NO;'";

				using (OleDbCommand comm = new OleDbCommand()) {
					conn.Open();
					DataTable dtSchema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables,
						new object[] { null, null, null, "TABLE" });
					foreach (DataRow row in dtSchema.Rows) {
						string name = row.Field<string>("TABLE_NAME");
						if (name.Contains("FilterDatabase"))
							continue;

						sheetNames.Add(name.Replace("$", "").TrimStart('\'').TrimEnd('\''));
					}
				}
			}

			return sheetNames;
		}
	}
}
