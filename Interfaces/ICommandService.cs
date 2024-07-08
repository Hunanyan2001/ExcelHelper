using System;

namespace ExcelHelper.Interfaces
{
    public interface ICommandService
    {
        event EventHandler CanExecuteChanged;
        void Execute(object parameter);
        bool CanExecute(object parameter);
    }
}
