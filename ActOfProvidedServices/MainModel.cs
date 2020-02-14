using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActOfProvidedServices {
	class MainModel {
		private double progressCurrent = 0;
		private readonly BackgroundWorker bw;
		private readonly string workbookFilePath;
		private readonly string worksheetName;
		private string template = string.Empty;

		private const string COLUMN_PATIENT_POLICY = "F3";
		private const string COLUMN_PATIENT_NAME = "F4";
		private const string COLUMN_PATIENT_CARD = "F7";
		private const string COLUMN_DIAGNOSIS = "F8";
		private const string COLUMN_TOOTH_NUMBER = "F9";
		private const string COLUMN_DATE = "F10";
		private const string COLUMN_SERVICE_CODE = "F11";
		private const string COLUMN_SERVICE_NAME = "F12";
		private const string COLUMN_SERVICE_COUNT = "F13";
		private const string COLUMN_SERVICE_COST = "F14";
		private const string COLUMN_SERVICE_DISCOUNTED_COST = "F16";
		private const string COLUMN_FILIAL = "F17";
		private const string COLUMN_DEPARTMENT = "F18";
		private const string COLUMN_EMPLOYEE_NAME = "F19";
		private const string COLUMN_GARANTY_MAIL = "F42";

		public MainModel(BackgroundWorker bw,
			string workbookFilePath,
			string worksheetName) {
			this.bw = bw;
			this.workbookFilePath = workbookFilePath;
			this.worksheetName = worksheetName;
		}

		public enum Type {
			Renessans,
			VTB,
			Rosgosstrakh,
			Reso
		}

		public void CreateAct(Type type, string period, string contract, string dateDischarged) {
			ExcelGeneral excelGeneral = new ExcelGeneral(bw);

			switch (type) {
				case Type.Renessans:
					template = "TemplateRenessans.xlsx";
					break;
				case Type.VTB:
					template = "TemplateVTB.xlsx";
					break;
				case Type.Rosgosstrakh:
					template = "TemplateRosgosstrakh.xlsx";
					break;
				case Type.Reso:
					template = "TemplateReso.xlsx";
					break;
				default:
					bw.ReportProgress((int)progressCurrent, "!!! Неизвестный тип организации, пропуск: " + type);
					break;
			}

			using (DataTable dataTable = excelGeneral.ReadExcelFile(workbookFilePath, worksheetName)) {
				bw.ReportProgress((int)progressCurrent, "Считано строк: " + dataTable.Rows.Count);

				if (dataTable.Rows.Count == 0)
					return;

				progressCurrent += 10;
				bw.ReportProgress((int)progressCurrent, "Анализ считанных данных");

				List<ItemPatient> patients = new List<ItemPatient>();
				ItemPatient currentPatient = null;
				ItemTreatment currentTreatment = null;

				double progressStep = 20.0d / (double)dataTable.Rows.Count;
				int i = 0;
				foreach (DataRow dataRow in dataTable.Rows) {
					i++;
					try {
						bw.ReportProgress((int)(progressCurrent += progressStep));
						string patientDocuments = dataRow[COLUMN_PATIENT_POLICY].ToString();
						if (string.IsNullOrEmpty(patientDocuments)) {
							bw.ReportProgress((int)progressCurrent, "!!! Отсутсвует номер полиса в строке: " + (i + 1) + ", пропуск");
							continue;
						} else if (patientDocuments.Equals("Полис"))
							continue;

						string patientCode = dataRow[COLUMN_PATIENT_CARD].ToString();
						string patientName = dataRow[COLUMN_PATIENT_NAME].ToString();

						if (type == Type.Renessans)
							patientName = ClearNameString(patientName);

						if (currentPatient != null && 
							(!patientCode.Equals(currentPatient.Code) || !currentPatient.Documents.Contains(patientDocuments))) {
							if (currentTreatment != null) {
								currentPatient.Treatments.Add(currentTreatment);
								currentTreatment = null;
							}

							patients.Add(currentPatient);
							currentPatient = null;
						}

						if (currentPatient == null) {
							currentPatient = new ItemPatient() {
								Name = patientName,
								Documents = patientDocuments,
								Code = patientCode
							};

							//if (type == Type.Renessans)
							//	currentPatient.Documents = "№ СП: " + currentPatient.Documents;

							if (type == Type.Reso ||
								type == Type.Renessans)
								currentPatient.Documents = "ГП № @number № СП: " + currentPatient.Documents;
						}

						string treatmentDoctor = ClearNameString(dataRow[COLUMN_EMPLOYEE_NAME].ToString());

						string treatmentDate = dataRow[COLUMN_DATE].ToString();
						if (double.TryParse(treatmentDate, out double rawDateDouble)) {
							try {
								treatmentDate = DateTime.FromOADate(rawDateDouble).ToShortDateString();
							} catch (Exception) { }
						}

						string treatmentFilial = dataRow[COLUMN_FILIAL].ToString();
						if (treatmentFilial.Equals("12"))
							treatmentFilial = "СУЩ";
						else if (treatmentFilial.Equals("5"))
							treatmentFilial = "М-СРЕТ";
						else if (treatmentFilial.Equals("1"))
							treatmentFilial = "МДМ";
						else if (treatmentFilial.Equals("6"))
							treatmentFilial = "КУТУЗ";

						if (currentTreatment != null &&
							(!treatmentDoctor.Equals(currentTreatment.Doctor) ||
							!treatmentDate.Equals(currentTreatment.Date))) {
							currentPatient.Treatments.Add(currentTreatment);
							currentTreatment = null;
						}

						if (currentTreatment == null) {
							currentTreatment = new ItemTreatment() {
								Doctor = treatmentDoctor,
								Date = treatmentDate,
								Filial = treatmentFilial
							};
						}

						string serviceName = dataRow[COLUMN_SERVICE_NAME].ToString();
						double serviceCount = Convert.ToDouble(dataRow[COLUMN_SERVICE_COUNT].ToString());
						string serviceCode = dataRow[COLUMN_SERVICE_CODE].ToString();
						double serviceCost = Convert.ToDouble(dataRow[COLUMN_SERVICE_COST].ToString());

						string serviceCostFinalValue = dataRow[COLUMN_SERVICE_DISCOUNTED_COST].ToString();

						double serviceCostFinal = 0;
						if (!string.IsNullOrEmpty(serviceCostFinalValue))
							serviceCostFinal = Convert.ToDouble(serviceCostFinalValue) / serviceCount;

						string serviceDiagnosis = dataRow[COLUMN_DIAGNOSIS].ToString();

						ItemService itemService = new ItemService() {
							Name = serviceName,
							Count = serviceCount,
							Code = serviceCode,
							Cost = serviceCost,
							CostFinal = serviceCostFinal
						};

						currentTreatment.Services.Add(itemService);

						if (!string.IsNullOrEmpty(serviceDiagnosis) &&
							serviceDiagnosis.Contains(" ")) {
							serviceDiagnosis = serviceDiagnosis.Split(' ')[0];
							if (!currentTreatment.Diagnoses.Contains(serviceDiagnosis))
								currentTreatment.Diagnoses.Add(serviceDiagnosis);
						}

						currentTreatment.TreatmentCostTotal += serviceCount * serviceCost;

						string department = dataRow[COLUMN_DEPARTMENT].ToString();
						if (!string.IsNullOrEmpty(department))
							if (department.ToLower().Contains("стоматология")) {
								string toothNumber = dataRow[COLUMN_TOOTH_NUMBER].ToString();
								itemService.ToothNumber = toothNumber;
							}

						string garantyMail = dataRow[COLUMN_GARANTY_MAIL].ToString();
						if (!string.IsNullOrEmpty(garantyMail)) {
							currentTreatment.GarantyMail = garantyMail;

							if (string.IsNullOrEmpty(currentPatient.GarantyMail))
								currentPatient.GarantyMail = garantyMail;
						}
					} catch (Exception e) {
						bw.ReportProgress((int)progressCurrent, "!!! Строка: " + i + ", ошибка: " + e.Message);
					}
				}

				if (currentPatient != null) {
					if (currentTreatment != null)
						currentPatient.Treatments.Add(currentTreatment);

					patients.Add(currentPatient);
				}

				bw.ReportProgress((int)progressCurrent, "Пациентов в списке: " + patients.Count);
				if (patients.Count == 0) {
					bw.ReportProgress((int)progressCurrent, "Не удалось получить информацию о пациентах. Пропуск выгрузки в Excel");
					return;
				}

				bw.ReportProgress((int)progressCurrent, "Выгрузка информации в акт о выполненных работах");

				string resultFile = excelGeneral.WritePatientsToExcelRenessans(
					patients, type, "Акт о выполненных работах", template, "Данные");

				if (string.IsNullOrEmpty(resultFile)) {
					bw.ReportProgress((int)progressCurrent, "Не удалось записать данные в файл");
					return;
				}

				progressCurrent += 20;
				bw.ReportProgress((int)progressCurrent, "Данные записаны в файл: " + resultFile);
				bw.ReportProgress((int)progressCurrent, "Применение форматирования. Пост-обработка");

				if (excelGeneral.WriteEndingRenessans(type, resultFile, progressCurrent, period,
					contract, dateDischarged, patients))
					bw.ReportProgress(100, "Обработка завершена успешно");
			}
		}

		private string ClearNameString(string name) {
			string patientName = name;
			if (patientName.Contains(" ") && !patientName.StartsWith("КДЛ")) {
				string[] splittedName = patientName.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
				patientName = splittedName[0];
				for (int i = 1; i < splittedName.Length; i++)
					patientName += " " + splittedName[i].Substring(0, 1) + ".";
			}

			return patientName;
		}
	}
}
