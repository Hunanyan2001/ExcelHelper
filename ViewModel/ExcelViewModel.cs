using ExcelHelper.Services;
using System;
using System.Threading.Tasks;
using System.Windows.Input;
using System.Collections.ObjectModel;
using ExcelHelper.Models;
using Microsoft.Win32;
using System.IO;
using System.Collections.Generic;
using ClosedXML.Excel;
using System.Linq;
using NPOI.HSSF.UserModel;
using System.Windows;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelHelper.ViewModel
{
    public class ExcelViewModel : ViewModelBase
    {

        public ExcelViewModel() 
        {
            ExportCommand = new CommandService(OnExport, CanExport);
            ImportCommand = new CommandService(OnImport, CanImport);
        }

        private string? _filePath;

        public string FilePath
        {
            get => _filePath;
            set
            {
                _filePath = value;
                OnPropertyChanged(nameof(FilePath));
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
        }

        public void OnExport(object parameter)
        {
            try
            {

            }
            catch(Exception ex)
            {

            }
        }

        public bool CanExport(object parameter)
        {
            return !string.IsNullOrEmpty(FilePath);
        }

        public async void OnImport(object parameter)
        {
            try
            {
                var openFileDialog = new OpenFileDialog
                {
                    Filter = "Excel Files|*.xls;*.xlsx"
                };

                if (openFileDialog.ShowDialog() == true)
                {
                    var excelData = new List<ExcelDataModel>();
                    FilePath = openFileDialog.FileName;
                    var fileExtension = Path.GetExtension(FilePath).ToLower();

                    if (fileExtension == ".xlsx" || fileExtension == ".xlsm" || fileExtension == ".xltx" || fileExtension == ".xltm")
                    {
                        excelData = await ReadExcelFileXlsx(FilePath);
                    }
                    else
                    {
                        excelData = await ReadExcelFileXls(FilePath);
                    }
                    ExcelDataModel = new ObservableCollection<ExcelDataModel>(excelData);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        

        public bool CanImport(object parameter)
        {
            return true;
        }

        public async Task LoadGridData()
        {
            //var data = await _yourDataService.GetData();
            //GridData = new ObservableCollection<YourDataModel>(data);
        }

        private async Task<List<ExcelDataModel>> ReadExcelFileXlsx(string filePath)
        {
            var dataModels = new List<ExcelDataModel>();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheets.First();
                var rows = worksheet.RowsUsed();
                var columnNames = worksheet.Row(1).Cells().Select(cell => cell.Value.ToString()).ToList();

                foreach (var row in rows.Skip(1))
                {
                    var model = new ExcelDataModel
                    {
                    };

                    dataModels.Add(model);
                }
            }
            return dataModels;
        }
        private async Task<List<ExcelDataModel>> ReadExcelFileXls(string filePath)
        {
            try
            {
                var dataModels = new List<ExcelDataModel>();

                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    var workbook = new HSSFWorkbook(stream);
                    var sheet = workbook.GetSheetAt(0);
                    int rowCount = sheet.LastRowNum;

                    var headerRow = sheet.GetRow(0);
                    var columnNames = new List<string>();

                    foreach (var cell in headerRow.Cells)
                    {
                        columnNames.Add(cell.ToString());
                    }

                    for (int row = 1; row <= rowCount; row++)
                    {
                        var excelRow = sheet.GetRow(row);
                        var model = new ExcelDataModel
                        {
                        };

                        dataModels.Add(model);
                    }
                }

                return dataModels;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
                return null;
            }
        }
    }
}
