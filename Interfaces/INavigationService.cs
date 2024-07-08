using System;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace ExcelHelper.Interfaces
{
    public interface INavigationService
    {
        void SetFrame(Frame frame);

        void NavigateTo(Type pageType, object? parameter = null);

        void NavigateTo<T>(object? parameter = null) where T : Page;

        Task NavigateTo<TParameter>(Type pageType, TParameter param, object? parameter = null);

        Task NavigateTo<TView, TParameter>(TParameter param, object? parameter = null)
            where TView : Page;

        void GoBack();
    }
}
