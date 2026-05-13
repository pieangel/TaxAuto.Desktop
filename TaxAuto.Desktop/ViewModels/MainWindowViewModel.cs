using Microsoft.Win32;
using System.IO;
using System.Windows;
using System.Windows.Input;
using TaxAuto.Desktop.Models;
using TaxAuto.Desktop.Services;

namespace TaxAuto.Desktop.ViewModels
{
    public class MainWindowViewModel
    {
        public DocumentOcrTabViewModel PurchaseTab { get; }
        public DocumentOcrTabViewModel SalesTab { get; }
        public DocumentOcrTabViewModel WorkOrderTab { get; }

        public MainWindowViewModel()
        {
            PurchaseTab = new DocumentOcrTabViewModel(DocumentKind.Purchase);
            SalesTab = new DocumentOcrTabViewModel(DocumentKind.Sales);
            WorkOrderTab = new DocumentOcrTabViewModel(DocumentKind.WorkOrder);
        }
    }
}
