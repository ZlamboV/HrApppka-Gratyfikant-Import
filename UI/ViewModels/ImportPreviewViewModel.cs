using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using InsERT.Moria.ModelDanych;
using InsERT.Moria.Rozszerzanie;
using InsERT.Moria.Sfera;
using HrAppka_Import_Pracowników.Services;

namespace HrAppka_Import_Pracowników.UI.ViewModels
{
    public class DuplicateResolutionOption
    {
        public DecyzjaIstniejacyPracownik Value { get; set; }
        public string DisplayName { get; set; } = string.Empty;
    }

    public class ImportPreviewViewModel : INotifyPropertyChanged
    {
        private readonly IUchwyt? _uchwyt;
        private readonly Func<List<PracownikImportRow>, IProgress<ImportProgressReport>, CancellationToken, Task>? _onExecuteImportAsync;
        private CancellationTokenSource? _cancellationTokenSource;

        private int _totalCount;
        private int _readyCount;
        private int _warningCount;
        private int _errorCount;

        private int _currentProgress;
        private int _totalProgress;
        private double _progressPercentage;
        private string _currentEmployeeName = string.Empty;
        private string _currentStatusMessage = "Gotowy do importu";

        private bool _isImporting;
        private bool _canCancel;
        private string _filterText = string.Empty;
        private string _selectedStatusFilter = "Wszystkie";

        public ObservableCollection<PracownikImportRow> Rows { get; }
        public ICollectionView RowsView { get; }

        public List<SzablonUmowyItem> AvailableTemplates { get; }
        public List<OrganPodatkowy> AvailableTaxOffices { get; }
        public List<DuplicateResolutionOption> DuplicateResolutions { get; }

        public event PropertyChangedEventHandler? PropertyChanged;

        public ICommand StartImportCommand { get; }
        public ICommand CancelImportCommand { get; }
        public ICommand SelectAllCommand { get; }
        public ICommand DeselectAllCommand { get; }

        public ImportPreviewViewModel(
            IUchwyt? uchwyt, 
            List<PracownikImportRow> rows,
            Func<List<PracownikImportRow>, IProgress<ImportProgressReport>, CancellationToken, Task>? onExecuteImportAsync = null)
        {
            _uchwyt = uchwyt;
            _onExecuteImportAsync = onExecuteImportAsync;

            // Inicjalizacja słowników dla kontrolek ComboBox
            if (_uchwyt != null)
            {
                AvailableTemplates = ContractCreationService.Instance.GetAvailableTemplates(_uchwyt);
            }
            else
            {
                AvailableTemplates = new List<SzablonUmowyItem>();
            }

            var defaultTemplate = _uchwyt != null ? ContractCreationService.Instance.GetDefaultTemplate(_uchwyt) : null;

            AvailableTaxOffices = (SferaDictionaryCache.Instance.AllTaxOffices ?? new List<OrganPodatkowy>())
                .OrderBy(o => o.Firma?.Nazwa ?? "")
                .ToList();

            DuplicateResolutions = new List<DuplicateResolutionOption>
            {
                new() { Value = DecyzjaIstniejacyPracownik.AktualizujKartotekeIDodajUmowe, DisplayName = "Aktualizuj kartotekę i dodaj umowę" },
                new() { Value = DecyzjaIstniejacyPracownik.TylkoDodajUmowe, DisplayName = "Tylko dodaj umowę" },
                new() { Value = DecyzjaIstniejacyPracownik.NadpiszUmowe, DisplayName = "Nadpisz umowę" },
                new() { Value = DecyzjaIstniejacyPracownik.PominWiersz, DisplayName = "Pomiń wiersz" }
            };

            // Domyślne przypisanie szablonu do wierszy, które go nie mają
            foreach (var r in rows)
            {
                if (r.SelectedSzablonUmowy == null && defaultTemplate != null)
                {
                    r.SelectedSzablonUmowy = defaultTemplate;
                }

                // Subskrypcja zmian do odświeżania liczników
                r.PropertyChanged += (s, e) =>
                {
                    if (e.PropertyName == nameof(PracownikImportRow.Status) || 
                        e.PropertyName == nameof(PracownikImportRow.StatusLevel))
                    {
                        UpdateSummaryCounts();
                    }
                };
            }

            Rows = new ObservableCollection<PracownikImportRow>(rows);
            RowsView = CollectionViewSource.GetDefaultView(Rows);
            RowsView.Filter = FilterRow;

            StartImportCommand = new RelayCommand(async () => await StartImportAsync(), () => !IsImporting);
            CancelImportCommand = new RelayCommand(CancelImport, () => IsImporting && CanCancel);
            SelectAllCommand = new RelayCommand(() => SetSelection(true), () => !IsImporting);
            DeselectAllCommand = new RelayCommand(() => SetSelection(false), () => !IsImporting);

            UpdateSummaryCounts();
        }

        public int TotalCount { get => _totalCount; private set => SetProperty(ref _totalCount, value); }
        public int ReadyCount { get => _readyCount; private set => SetProperty(ref _readyCount, value); }
        public int WarningCount { get => _warningCount; private set => SetProperty(ref _warningCount, value); }
        public int ErrorCount { get => _errorCount; private set => SetProperty(ref _errorCount, value); }

        public int CurrentProgress { get => _currentProgress; private set => SetProperty(ref _currentProgress, value); }
        public int TotalProgress { get => _totalProgress; private set => SetProperty(ref _totalProgress, value); }
        public double ProgressPercentage { get => _progressPercentage; private set => SetProperty(ref _progressPercentage, value); }
        public string CurrentEmployeeName { get => _currentEmployeeName; private set => SetProperty(ref _currentEmployeeName, value); }
        public string CurrentStatusMessage { get => _currentStatusMessage; private set => SetProperty(ref _currentStatusMessage, value); }

        public bool CanClose => !_isImporting;

        public bool IsImporting
        {
            get => _isImporting;
            private set
            {
                if (SetProperty(ref _isImporting, value))
                {
                    OnPropertyChanged(nameof(CanClose));
                    CommandManager.InvalidateRequerySuggested();
                }
            }
        }

        public bool CanCancel
        {
            get => _canCancel;
            private set => SetProperty(ref _canCancel, value);
        }

        public string FilterText
        {
            get => _filterText;
            set
            {
                if (SetProperty(ref _filterText, value))
                {
                    RowsView.Refresh();
                }
            }
        }

        public string SelectedStatusFilter
        {
            get => _selectedStatusFilter;
            set
            {
                if (SetProperty(ref _selectedStatusFilter, value))
                {
                    RowsView.Refresh();
                }
            }
        }

        private bool FilterRow(object item)
        {
            if (item is not PracownikImportRow row) return false;

            if (SelectedStatusFilter == "Błędy" && row.StatusLevel != RowStatusLevel.CriticalError)
                return false;
            if (SelectedStatusFilter == "Ostrzeżenia" && row.StatusLevel != RowStatusLevel.Warning)
                return false;
            if (SelectedStatusFilter == "Gotowe" && row.StatusLevel != RowStatusLevel.Ready)
                return false;

            if (!string.IsNullOrWhiteSpace(FilterText))
            {
                string query = FilterText.Trim().ToLowerInvariant();
                bool matches = (row.PracownikNazwa?.ToLowerInvariant().Contains(query) ?? false) ||
                                (row.Pesel?.Contains(query) ?? false) ||
                                (row.Paszport?.ToLowerInvariant().Contains(query) ?? false) ||
                                (row.DzialNazwa?.ToLowerInvariant().Contains(query) ?? false);
                if (!matches) return false;
            }

            return true;
        }

        private void SetSelection(bool isSelected)
        {
            foreach (var r in RowsView.OfType<PracownikImportRow>())
            {
                r.IsSelected = isSelected;
            }
        }

        public void UpdateSummaryCounts()
        {
            TotalCount = Rows.Count;
            ReadyCount = Rows.Count(r => r.StatusLevel == RowStatusLevel.Ready);
            WarningCount = Rows.Count(r => r.StatusLevel == RowStatusLevel.Warning);
            ErrorCount = Rows.Count(r => r.StatusLevel == RowStatusLevel.CriticalError);
        }

        private async Task StartImportAsync()
        {
            var selectedRows = Rows.Where(r => r.IsSelected && r.StatusLevel != RowStatusLevel.CriticalError).ToList();
            if (!selectedRows.Any())
            {
                MessageBox.Show(
                    "Nie wybrano żadnych poprawnych wierszy do zaimportowania (wiersze z błędami krytycznymi są pomijane).",
                    "Import Pracowników",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            IsImporting = true;
            CanCancel = true;
            _cancellationTokenSource = new CancellationTokenSource();
            var token = _cancellationTokenSource.Token;

            TotalProgress = selectedRows.Count;
            CurrentProgress = 0;
            ProgressPercentage = 0;

            var progress = new Progress<ImportProgressReport>(report =>
            {
                CurrentProgress = report.Current;
                TotalProgress = report.Total;
                ProgressPercentage = report.Total > 0 ? ((double)report.Current / report.Total) * 100 : 0;
                CurrentEmployeeName = report.CurrentEmployee;
                CurrentStatusMessage = report.StatusMessage;
            });

            try
            {
                if (_onExecuteImportAsync != null)
                {
                    await Task.Run(async () =>
                    {
                        await _onExecuteImportAsync(selectedRows, progress, token);
                    }, token);
                }
                else
                {
                    await Task.Run(() =>
                    {
                        var progressHandler = (IProgress<ImportProgressReport>)progress;
                        int count = selectedRows.Count;
                        var pendingZusBatch = new List<ZusRegistrationRequest>();

                        for (int i = 0; i < count; i++)
                        {
                            if (token.IsCancellationRequested)
                            {
                                progressHandler.Report(new ImportProgressReport(i, count, "Przerwano", "Operacja importu została przerwana przez użytkownika."));
                                break;
                            }

                            var row = selectedRows[i];
                            progressHandler.Report(new ImportProgressReport(i + 1, count, $"{row.Nazwisko} {row.Imie}".Trim(), "Importowanie kartoteki i umowy..."));

                            try
                            {
                                var wynik = OperacjaImportuPracownikow20.ZaimportujRow20(
                                    kontekstOperacji: null,
                                    item: row,
                                    fallbackUchwyt: _uchwyt,
                                    pendingZusBatch: pendingZusBatch);

                                bool success = (wynik != null && wynik.Status == InsERT.Moria.Rozszerzanie.Operacje.StatusOperacji.ZakonczonaPowodzeniem);
                                string statusMsg = success ? (wynik?.Komunikat ?? "Pomyślnie zaimportowano") : (wynik?.Komunikat ?? "Wystąpił błąd");
                                progressHandler.Report(new ImportProgressReport(i + 1, count, $"{row.Nazwisko} {row.Imie}".Trim(), statusMsg));
                            }
                            catch (Exception ex)
                            {
                                row.Status = "Błąd";
                                row.Ostrzezenia += ex.Message + " ";
                                progressHandler.Report(new ImportProgressReport(i + 1, count, $"{row.Nazwisko} {row.Imie}".Trim(), $"Błąd: {ex.Message}"));
                            }
                        }

                        // Zbiorczy zapis deklaracji ZUS (Wymaganie R7)
                        if (pendingZusBatch.Count > 0 && !token.IsCancellationRequested && _uchwyt != null)
                        {
                            progressHandler.Report(new ImportProgressReport(count, count, "Deklaracje ZUS", "Zbiorczy zapis deklaracji ZUS..."));
                            ZusDeclarationBatchService.ProcessAndSaveBatch(_uchwyt, pendingZusBatch);
                        }
                    }, token);
                }

                CurrentStatusMessage = token.IsCancellationRequested 
                    ? "Import został przerwany przez użytkownika." 
                    : "Import został zakończony pomyślnie.";
            }
            catch (OperationCanceledException)
            {
                CurrentStatusMessage = "Import został anulowany.";
            }
            catch (Exception ex)
            {
                CurrentStatusMessage = $"Błąd podczas importu: {ex.Message}";
                MessageBox.Show($"Wystąpił błąd podczas importu: {ex.Message}", "Błąd importu", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsImporting = false;
                CanCancel = false;
                UpdateSummaryCounts();
            }
        }

        private void CancelImport()
        {
            if (_cancellationTokenSource != null && !_cancellationTokenSource.IsCancellationRequested)
            {
                CurrentStatusMessage = "Przerywanie operacji...";
                CanCancel = false;
                _cancellationTokenSource.Cancel();
            }
        }

        protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        protected bool SetProperty<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value)) return false;
            field = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            return true;
        }
    }
}
