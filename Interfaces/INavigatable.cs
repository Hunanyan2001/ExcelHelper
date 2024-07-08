using System.Threading.Tasks;

namespace ExcelHelper.Interfaces
{
    public interface INavigatable<TModel>
    {
        Task OnNavigate(TModel model);

        Task NavigateFrom();
    }
}
