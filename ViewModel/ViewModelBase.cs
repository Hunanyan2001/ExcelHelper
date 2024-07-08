using ExcelHelper.Interfaces;
using ExcelHelper.Models;

using System.Threading.Tasks;

namespace ExcelHelper.ViewModel
{
    public class ViewModelBase : BindableObject
    {
    }
    public class NavigatableViewModelBase<TModel> : ViewModelBase, INavigatable<TModel>
    {
        public virtual Task NavigateFrom()
        {
            return Task.CompletedTask;
        }

        public virtual Task OnNavigate(TModel model)
        {
            return Task.CompletedTask;
        }
    }
}
