using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using InsERT.Moria.Sfera;
using InsERT.Moria.Rozszerzanie;
using HrAppka_Import_Pracowników.UI.ViewModels;

namespace HrAppka_Import_Pracowników.UI
{
    /// <summary>
    /// Logika interakcji dla okna ImportPreviewWindow.xaml.
    /// Realizuje Wymaganie R6 (Modalne okno WPF z wirtualizacją DataGrid i asynchronicznym importem).
    /// </summary>
    public partial class ImportPreviewWindow : Window
    {
        public ImportPreviewViewModel ViewModel { get; }

        public ImportPreviewWindow(
            IUchwyt? uchwyt, 
            List<PracownikImportRow> rows,
            Func<List<PracownikImportRow>, IProgress<ImportProgressReport>, CancellationToken, Task>? onExecuteImportAsync = null)
        {
            InitializeComponent();
            ViewModel = new ImportPreviewViewModel(uchwyt, rows, onExecuteImportAsync);
            DataContext = ViewModel;
        }

        public ImportPreviewWindow(IUchwyt? uchwyt, List<PracownikImportRow> rows)
            : this(uchwyt, rows, null)
        {
        }

        private void OnCloseClicked(object sender, RoutedEventArgs e)
        {
            if (ViewModel.IsImporting)
            {
                var result = MessageBox.Show(
                    "Trwa proces importu. Czy na pewno chcesz przerwać i zamknąć okno?",
                    "Potwierdzenie zamknięcia",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

                if (result == MessageBoxResult.No)
                    return;

                ViewModel.CancelImportCommand.Execute(null);
            }

            DialogResult = true;
            Close();
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            if (ViewModel.IsImporting)
            {
                var result = MessageBox.Show(
                    "Trwa import danych. Zamknięcie okna przerwie operację. Kontynuować?",
                    "Import w toku",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);

                if (result == MessageBoxResult.No)
                {
                    e.Cancel = true;
                    return;
                }

                ViewModel.CancelImportCommand.Execute(null);
            }

            base.OnClosing(e);
        }
    }
}
