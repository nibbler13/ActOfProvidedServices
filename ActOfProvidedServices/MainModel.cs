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

		public MainModel(BackgroundWorker bw,
			string workbookFilePath,
			string worksheetName) {
			this.bw = bw;
			this.workbookFilePath = workbookFilePath;
			this.worksheetName = worksheetName;
		}

		public void CreateAct(string organization, string period, string contract, string dateDischarged) {
			ExcelGeneral excelGeneral = new ExcelGeneral(bw);

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
				foreach (DataRow dataRow in dataTable.Rows) {
					try {
						bw.ReportProgress((int)(progressCurrent += progressStep));
						string patientDocuments = dataRow["F2"].ToString();
						if (string.IsNullOrEmpty(patientDocuments) ||
							patientDocuments.Equals("№ полиса"))
							continue;

						string patientCode = dataRow["F5"].ToString();
						string patientName = ClearNameString(dataRow["F1"].ToString());

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
								Documents = "№ СП: " + patientDocuments,
								Code = patientCode
							};
						}

						string treatmentDoctor = ClearNameString(dataRow["F16"].ToString());
						string treatmentDate = dataRow["F7"].ToString().Replace(" 0:00:00", "");

						if (currentTreatment != null &&
							(!treatmentDoctor.Equals(currentTreatment.Doctor) ||
							!treatmentDate.Equals(currentTreatment.Date))) {
							currentPatient.Treatments.Add(currentTreatment);
							currentTreatment = null;
						}

						if (currentTreatment == null) {
							currentTreatment = new ItemTreatment() {
								Doctor = treatmentDoctor,
								Date = treatmentDate
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
					} catch (Exception) { }
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

				string resultFile = excelGeneral.WritePatientsToExcel(
					patients, "Акт о выполненных работах", "Template.xlsx", "Данные");

				if (string.IsNullOrEmpty(resultFile)) {
					bw.ReportProgress((int)progressCurrent, "Не удалось записать данные в файл");
					return;
				}

				progressCurrent += 20;
				bw.ReportProgress((int)progressCurrent, "Данные записаны в файл: " + resultFile);
				bw.ReportProgress((int)progressCurrent, "Применение форматирования. Пост-обработка");

				if (excelGeneral.Process(resultFile,
							 progressCurrent,
							 organization,
							 period,
							 contract,
							 dateDischarged,
							 patients))
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
