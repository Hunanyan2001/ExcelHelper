using ExcelHelper.ViewModel;

using System.Windows;
using System.Windows.Controls;

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
            Loaded += MainView_Loaded;
        }
        private async void MainView_Loaded(object sender, RoutedEventArgs e)
        {
            await ViewModel.Loaded().ConfigureAwait(false);
        }
    }
}
