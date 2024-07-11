using ExcelHelper.Services;
using System;
using System.Windows.Input;
using System.Collections.ObjectModel;
<<<<<<< HEAD
using System.Collections.Generic;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows;
using System.Data;
using Microsoft.Win32;
using System.Linq;
using ClosedXML.Excel;
using ExcelHelper.Constants;
=======
using ExcelHelper.Models;
using Microsoft.Win32;
using System.IO;
using System.Collections.Generic;
using ClosedXML.Excel;
using System.Linq;
using NPOI.HSSF.UserModel;
using System.Windows;
using DocumentFormat.OpenXml.Spreadsheet;
>>>>>>> main

namespace ExcelHelper.ViewModel
{
    public class ExcelViewModel : ViewModelBase
    {
<<<<<<< HEAD
        public DataGrid ExcelDataGrid { get; set; }

        public ExcelViewModel()
=======

        public ExcelViewModel() 
>>>>>>> main
        {
            ExportCommand = new CommandService(OnExport, CanExport);
            ImportCommand = new CommandService(OnImport, CanImport);
        }

<<<<<<< HEAD
        private ObservableCollection<Dictionary<string, object>> _dynamicData;


        public ObservableCollection<Dictionary<string, object>> DynamicData
        {
            get => _dynamicData;
            set
            {
                _dynamicData = value;
                OnPropertyChanged(nameof(DynamicData));
            }
        }

        private ObservableCollection<string> _columnHeaders;

        public ObservableCollection<string> ColumnHeaders
        {
            get => _columnHeaders;
            set
            {
                _columnHeaders = value;
                OnPropertyChanged(nameof(ColumnHeaders));
            }
        }
=======
>>>>>>> main
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

        #region Command
        public ICommand? ExportCommand { get; set; }

        public ICommand? ImportCommand { get; set; }

        public ICommand? DeleteCommand { get; set; }
        #endregion

        private bool CanExport(object parameter)
        {
<<<<<<< HEAD
            return !string.IsNullOrEmpty(FilePath);
=======
            LoadGridData();
>>>>>>> main
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

<<<<<<< HEAD
        private bool CanImport(object parameter)
=======
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
>>>>>>> main
        {
            return true;
        }

        private void OnImport(object parameter)
        {
            try
            {
                var dialog = new OpenFileDialog
                {
                    DefaultExt = FileConstantHelper.DefaultExt,
                    Filter = FileConstantHelper.Filter
                };

                bool? result = dialog.ShowDialog();

                if (result == true)
                {
                    FilePath = dialog.FileName;

                    using (var workbook = new XLWorkbook(FilePath))
                    {
                        var worksheet = workbook.Worksheets.First();

                        var dataTable = new DataTable();

                        foreach (var cell in worksheet.Row(1).CellsUsed())
                        {
                            dataTable.Columns.Add(cell.Value.ToString());
                        }

                        foreach (var row in worksheet.RowsUsed().Skip(1))
                        {
                            var dataRow = dataTable.NewRow();
                            int i = 0;
                            foreach (var cell in row.CellsUsed())
                            {
                                dataRow[i++] = cell.Value.ToString();
                            }
                            dataTable.Rows.Add(dataRow);
                        }

                        ColumnHeaders = new ObservableCollection<string>(dataTable.Columns.Cast<DataColumn>().Select(c => c.ColumnName));
                        DynamicData = new ObservableCollection<Dictionary<string, object>>(
                            dataTable.AsEnumerable().Select(row =>
                                dataTable.Columns.Cast<DataColumn>().ToDictionary(col => col.ColumnName, col => row[col])));

                        GenerateColumns();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void GenerateColumns()
        {
            if (ColumnHeaders == null || DynamicData == null || ExcelDataGrid == null)
                return;

            ExcelDataGrid.Columns.Clear();
            foreach (var header in ColumnHeaders)
            {
                var column = new DataGridTextColumn
                {
                    Header = header,
                    Binding = new Binding($"[{header}]")
                };
                ExcelDataGrid.Columns.Add(column);
            }
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
