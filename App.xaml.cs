using ExcelHelper.Interfaces;
using ExcelHelper.Services;
using ExcelHelper.View;
using ExcelHelper.ViewModel;

using Microsoft.Extensions.DependencyInjection;

using System.Windows;
using System.Windows.Input;

namespace ExcelHelper
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public IServiceScope ApplicationScope { get; private set; }
        public App()
        {
            var serviceCollection = new ServiceCollection();
            ConfigureServices(serviceCollection);
            ApplicationScope = serviceCollection.BuildServiceProvider().CreateScope();
            ApplicationScope.ServiceProvider.GetService<ScopeManager>().ApplicationScope = ApplicationScope;
        }

        public void ConfigureServices(IServiceCollection serviceCollection)
        {
            serviceCollection.AddSingleton<ScopeManager>();
            serviceCollection.AddScoped<INavigationService, NavigationService>();
            serviceCollection.AddScoped<ICommand, CommandService>();
            serviceCollection.AddScoped<MainWindow>();
            serviceCollection.AddScoped<MainWindowViewModel>();
            serviceCollection.AddScoped<ExcelViewModel>();
            serviceCollection.AddScoped<ExcelView>();
            serviceCollection.AddScoped(serviceProvider => Dispatcher);
        }

        protected override void OnStartup(StartupEventArgs e)
        {
            var mainWindowViewModel = ApplicationScope.ServiceProvider.GetService<MainWindowViewModel>();
            var mainWindow = new MainWindow(mainWindowViewModel);
            mainWindow.Show();
            base.OnStartup(e);
        }
    }
    public class ScopeManager
    {
        public ScopeManager() { }

        public IServiceScope ApplicationScope { get; set; }
    }
}
