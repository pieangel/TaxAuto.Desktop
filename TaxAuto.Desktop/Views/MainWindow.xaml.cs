using System.Windows;
using TaxAuto.Desktop.Services;
using TaxAuto.Desktop.ViewModels;

namespace TaxAuto.Desktop.Views
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            DataContext = new MainWindowViewModel();
        }
    }
}