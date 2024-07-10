using ExcelHelper.ViewModel;

using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace ExcelHelper.View
{
    /// <summary>
    /// Interaction logic for Page1.xaml
    /// </summary>
    public partial class ExcelView : Page
    {
        public ExcelViewModel ViewModel { get; }
        public ExcelView(ExcelViewModel viewModel)
        {
            InitializeComponent();
            ViewModel = viewModel;
            DataContext = ViewModel;
            Loaded += ExcelView_Loaded;
        }

        private void ExcelView_Loaded(object sender, RoutedEventArgs e)
        {
            ViewModel.ExcelDataGrid = ExcelDataGrid;
        }
    }
}
