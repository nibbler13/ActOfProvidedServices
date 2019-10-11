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
						string patientDocuments = dataRow["F2"].ToString();
						if (string.IsNullOrEmpty(patientDocuments)) {
							bw.ReportProgress((int)progressCurrent, "!!! Отсутсвует номер полиса в строке: " + i + ", пропуск");
							continue;
						} else if (patientDocuments.Equals("№ полиса"))
							continue;

						string patientCode = dataRow["F5"].ToString();
						string patientName = dataRow["F1"].ToString();

						if (type == Type.Renessans)
							patientName = ClearNameString(patientName);

						if (currentPatient != null &&
							!patientCode.Equals(currentPatient.Code)) {
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

							if (type == Type.Renessans ||
								type == Type.Reso)
								currentPatient.Documents = "№ СП: " + currentPatient.Documents;
						}

						string treatmentDoctor = ClearNameString(dataRow["F16"].ToString());
						string treatmentDate = dataRow["F7"].ToString().Replace(" 0:00:00", "");
						string treatmentFilial = dataRow["F15"].ToString();
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

						string serviceName = dataRow["F9"].ToString();
						double serviceCount = Convert.ToDouble(dataRow["F10"].ToString());
						string serviceCode = dataRow["F8"].ToString();
						double serviceCost = Convert.ToDouble(dataRow["F11"].ToString());
						double serviceCostFinal = Convert.ToDouble(dataRow["F13"].ToString());
						string serviceDiagnosis = dataRow["F6"].ToString();

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
