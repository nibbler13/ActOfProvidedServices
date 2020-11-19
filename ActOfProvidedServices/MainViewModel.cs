using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace ActOfProvidedServices {
	class MainViewModel : INotifyPropertyChanged {
		public event PropertyChangedEventHandler PropertyChanged;
		private void NotifyPropertyChanged([CallerMemberName] string propertyName = "", bool recalculate = true) {
			PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
		}

		public DateTime? DateDischarged { get; set; } = DateTime.Now;
		public string TextOrganization { get; set; }
		public string TextContract { get; set; }
		public string TextPeriod { get; set; }

		private string textWorkbookPath;
		public string TextWorkbookPath { 
		   get { return textWorkbookPath; }
		   set {
				if (value != textWorkbookPath) {
					textWorkbookPath = value;
					NotifyPropertyChanged();
				}
			}
		}


		private ICommand commandSelectWorkbook;
		public ICommand CommandSelectWorkbook {
			get {
				return commandSelectWorkbook ?? (commandSelectWorkbook = new CommandHandler(() => SelectWorkbookFile(), () => true));
			}
		}

		private ICommand commandExecute;
		public ICommand CommandExecute {
			get {
				return commandExecute ?? (commandExecute = new CommandHandler(() => Execute(), () => true));
			}
		}


		private Visibility gridMainVisibility = Visibility.Visible;
		public Visibility GridMainVisibility {
			get { return gridMainVisibility; }
			set {
				if (value != gridMainVisibility) {
					gridMainVisibility = value;
					NotifyPropertyChanged();
				}
			}
		}

		private Visibility gridResultVisibility = Visibility.Hidden;
		public Visibility GridResultVisibility {
			get { return gridResultVisibility; }
			set {
				if (value != gridResultVisibility) {
					gridResultVisibility = value;
					NotifyPropertyChanged();
				}
			}
		}

		private string textResult;
		public string TextResult {
			get { return textResult; }
			set {
				if (value != textResult) {
					textResult = value;
					NotifyPropertyChanged();
				}
			}
		}

		private int progressValue;
		public int ProgressValue { 
			get { return progressValue; } 
			set {
				if (value != progressValue) {
					progressValue = value;
					NotifyPropertyChanged();
				}
			}
		}




		private ICommand commandCloseResults;
		public ICommand CommandCloseResults {
			get {
				return commandCloseResults ?? (commandCloseResults = new CommandHandler(() => CloseResult(), () => CanUseButtonCloseResults));
			}
		}

		private bool canUseButtonCloseResults = false;
		public bool CanUseButtonCloseResults {
			get {
				return canUseButtonCloseResults;
			}
			set {
				if (value != canUseButtonCloseResults) {
					canUseButtonCloseResults = value;
					NotifyPropertyChanged();
				}
			}
		}

		public bool IsCheckedRenessans { get; set; }
		public bool IsCheckedVTB { get; set; }
		public bool IsCheckedRosgosstrakh { get; set; }
		public bool IsCheckedResoGaranty { get; set; }

		public ObservableCollection<string> SheetNames { get; set; } = new ObservableCollection<string>();

		private string selectedSheetName;
		public string SelectedSheetName {
			get { return selectedSheetName; }
			set {
				if (value != selectedSheetName) {
					selectedSheetName = value;
					NotifyPropertyChanged();
				}
			}
		}


		private bool sheetNamesComboboxEnabled = false;
		public bool SheetNamesComboboxEnabled {
			get { return sheetNamesComboboxEnabled; }
			set {
				if (value != sheetNamesComboboxEnabled) {
					sheetNamesComboboxEnabled = value;
					NotifyPropertyChanged();
				}
			}
		}


		public void CloseResult() {
			TextResult = string.Empty;
			ProgressValue = 0;
			GridResultVisibility = Visibility.Hidden;
			GridMainVisibility = Visibility.Visible;
		}
		
		public void SelectWorkbookFile() {
			SheetNames.Clear();
			SheetNamesComboboxEnabled = false;
			SelectedSheetName = string.Empty;

			OpenFileDialog openFileDialog = new OpenFileDialog();
			openFileDialog.Filter = "Книга Excel (*.xls*)|*.xls*";
			openFileDialog.CheckFileExists = true;
			openFileDialog.CheckPathExists = true;
			openFileDialog.Multiselect = false;
			openFileDialog.RestoreDirectory = true;

			if (openFileDialog.ShowDialog() == true) {
				TextWorkbookPath = openFileDialog.FileName;
				ExcelGeneral.ReadSheetNames(TextWorkbookPath).ForEach(SheetNames.Add);
				SheetNamesComboboxEnabled = true;

				if (SheetNames.Count > 0)
					SelectedSheetName = SheetNames[0];
			}
		}





		public void Execute() {
			string errors = string.Empty;

			if (string.IsNullOrEmpty(TextWorkbookPath))
				errors = "Не выбран файл книги Excel" + Environment.NewLine;

			if (string.IsNullOrEmpty(SelectedSheetName))
				errors += "Не указано имя листа" + Environment.NewLine;

			if (!IsCheckedRenessans && !IsCheckedResoGaranty &&
				!IsCheckedRosgosstrakh && !IsCheckedVTB)
				errors += "Не выбрана организация";

			if (!string.IsNullOrEmpty(errors)) {
				MessageBox.Show(Application.Current.MainWindow,
					"Невозможно выполнить обработку: " + Environment.NewLine + errors,
					string.Empty,
					MessageBoxButton.OK,
					MessageBoxImage.Information);
				return;
			}

			GridMainVisibility = Visibility.Hidden;
			GridResultVisibility = Visibility.Visible;

			using (BackgroundWorker bw = new BackgroundWorker()) {
				bw.ProgressChanged += Bw_ProgressChanged;
				bw.WorkerReportsProgress = true;
				bw.DoWork += Bw_DoWork;
				bw.RunWorkerCompleted += Bw_RunWorkerCompleted;
				bw.RunWorkerAsync();
			}
		}

		private void Bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) {
			CanUseButtonCloseResults = true;

			if (e.Error != null) {
				MessageBox.Show(Application.Current.MainWindow,
					e.Error.Message + Environment.NewLine + e.Error.StackTrace,
					"Ошибки во время выполнения",
					MessageBoxButton.OK,
					MessageBoxImage.Error);
				return;
			}

			ProgressValue = 100;

			MessageBox.Show(Application.Current.MainWindow,
					"Выполнение завершено",
					string.Empty,
					MessageBoxButton.OK,
					MessageBoxImage.Information);
		}

		private void Bw_DoWork(object sender, DoWorkEventArgs e) {
			MainModel mainModel = new MainModel(
				(sender as BackgroundWorker),
				TextWorkbookPath,
				SelectedSheetName);

			MainModel.Type type;
			if (IsCheckedRenessans)
				type = MainModel.Type.Renessans;
			else if (IsCheckedResoGaranty)
				type = MainModel.Type.Reso;
			else if (IsCheckedRosgosstrakh)
				type = MainModel.Type.Rosgosstrakh;
			else if (IsCheckedVTB)
				type = MainModel.Type.VTB;
			else {
				(sender as BackgroundWorker).ReportProgress(0, "!!! Неизвестный тип организации, пропуск");
				return;
			}

			mainModel.CreateAct(type, TextPeriod, TextContract,
				DateDischarged.HasValue ? DateDischarged.Value.ToShortDateString() : "");
		}

		private void Bw_ProgressChanged(object sender, ProgressChangedEventArgs e) {
			ProgressValue = e.ProgressPercentage;

			if (e.UserState != null)
				TextResult += DateTime.Now.ToLongTimeString() + ": " + 
					e.UserState.ToString() + Environment.NewLine;
		}
	}
}
