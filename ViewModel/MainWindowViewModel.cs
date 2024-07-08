using ExcelHelper.Interfaces;
using ExcelHelper.View;

using System.Windows.Controls;

namespace ExcelHelper.ViewModel
{
    public class MainWindowViewModel : ViewModelBase
    {
        private readonly INavigationService _navigationService;

        public MainWindowViewModel(INavigationService navigationService)
        {
            _navigationService = navigationService;
        }

        public void Loaded(Frame mainFrame)
        {
            _navigationService.SetFrame(mainFrame);
            _navigationService.NavigateTo<ExcelView>();
        }
    }
}
