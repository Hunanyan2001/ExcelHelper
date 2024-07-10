using ExcelHelper.Services;
using System;
using System.Windows.Input;
using System.Collections.ObjectModel;
using System.Collections.Generic;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows;
using System.Data;
using Microsoft.Win32;
using System.Linq;
using ClosedXML.Excel;
using ExcelHelper.Constants;

namespace ExcelHelper.ViewModel
{
    public class ExcelViewModel : ViewModelBase
    {
        public DataGrid ExcelDataGrid { get; set; }

        public ExcelViewModel()
        {
            ExportCommand = new CommandService(OnExport, CanExport);
            ImportCommand = new CommandService(OnImport, CanImport);
        }

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

        #region Command
        public ICommand? ExportCommand { get; set; }

        public ICommand? ImportCommand { get; set; }

        public ICommand? DeleteCommand { get; set; }
        #endregion

        private bool CanExport(object parameter)
        {
            return !string.IsNullOrEmpty(FilePath);
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

        private bool CanImport(object parameter)
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
    }
}
