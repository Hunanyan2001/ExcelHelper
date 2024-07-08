using ExcelHelper.Services;
using System;
using System.Threading.Tasks;
using System.Windows.Input;
using System.Collections.ObjectModel;
using ExcelHelper.Models;

namespace ExcelHelper.ViewModel
{
    public class ExcelViewModel : ViewModelBase
    {
        private string? _filePath;

        public string FilePath
        {
            get => _filePath;
            set
            {
                _filePath = value;
                OnPropertyChanged(FilePath);
            }
        }

        public ObservableCollection<ExcelDataModel> _excelDataModels;

        public ObservableCollection<ExcelDataModel> ExcelDataModel
        {
            get => _excelDataModels;
            set
            {
                _excelDataModels = value;
                OnPropertyChanged(nameof(ExcelDataModel));
            }
        }

        #region Command
        public ICommand? ExportCommand { get; set; }

        public ICommand? ImportCommand { get; set; }

        public ICommand? DeleteCommand { get; set; }
        #endregion

        public async Task Loaded()
        {
            LoadGridData();
            ExportCommand = new CommandService(OnExport, CanExport);
            ImportCommand = new CommandService(OnImport, CanImport);
        }

        private void OnExport(object parameter)
        {
            try
            {

            }
            catch(Exception ex)
            {

            }
        }

        private bool CanExport(object parameter)
        {
            return !string.IsNullOrEmpty(FilePath);
        }

        private void OnImport(object parameter)
        {
           
        }

        private bool CanImport(object parameter)
        {
            return true;
        }

        public async Task LoadGridData()
        {
            //var data = await _yourDataService.GetData();
            //GridData = new ObservableCollection<YourDataModel>(data);
        }
    }
}
