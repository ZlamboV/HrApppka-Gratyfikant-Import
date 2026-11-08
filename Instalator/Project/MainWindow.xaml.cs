using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;

namespace InsERT.Mox.QuickInstallerCore
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window, INotifyPropertyChanged
	{
		private const int MaxIconsForExpandedView = 2;

		private readonly ICollectionView _databaseCollectionView;
		private bool _autoScroll = true;
		private bool _checkedAll = true;
		private ObservableCollection<DatabaseInfoWrapper> _databaseCollection;
		private Visibility _databaseSelectionVisibility = Visibility.Collapsed;
		private Etap _etap;
		private bool _flagaUstawianieUtworzSkrotNaPulpicie = false;
		private bool _flagaUstawianieUtworzSkrotWMenuStart = false;
		private string _informacja;
		private InstallationConfig _konfiguracja;
		private Visibility _licencjaVisibility = Visibility.Collapsed;
		private bool _moznaDalej;
		private Visibility _restartInfoVisibility = Visibility.Collapsed;
		private string[] _scenarios;
		private ObservableCollection<ShortcutCfg> _shortcutCfgCollection;
		private bool _uninstallMode;
		private bool? _utworzSkrotNaPulpicie = true;
		private Visibility _utworzSkrotVisibility = Visibility.Collapsed;
		private bool? _utworzSkrotWMenuStart = true;
		private Installer _installer;

		public MainWindow(bool uninstall)
		{
			_uninstallMode = uninstall;
			InitializeComponent();
			DataContext = this;

			if (_uninstallMode)
			{
				OdinstalujLink.HideAndDisable();
			}
			Application.Current.Resources.Add("ObjectEditWorkspaceBackground", new SolidColorBrush(System.Windows.Media.Color.FromRgb(238, 238, 238)));

			_databaseCollection = [];
			_databaseCollectionView = CollectionViewSource.GetDefaultView(_databaseCollection);
			_databaseCollectionView.Filter = ShouldShowDatabase;
		}

		public event PropertyChangedEventHandler PropertyChanged;

		private enum Etap
		{
			Etap0_Start,
			Etap1_Licencja,
			Etap2_PrzedInstalacja,
			Etap3_Informacja,
			Etap4_WyborBaz,
			Etap5_Instalacja,
			Etap6_Podsumowanie
		};

		public bool CheckedAll
		{
			get => _checkedAll;
			set
			{
				_checkedAll = value;
				foreach (var database in _databaseCollectionView.OfType<DatabaseInfoWrapper>())
				{
					if (database.IsEnabled && database.IsChecked != _checkedAll)
					{
						database.IsChecked = _checkedAll;
					}
				}
				OnPropertyChanged(nameof(CheckedAll));
			}
		}

		public ObservableCollection<DatabaseInfoWrapper> DatabaseCollection
		{
			get => _databaseCollection;
			set
			{
				_databaseCollection = value;
				OnPropertyChanged(nameof(DatabaseCollection));
			}
		}

		public string DatabaseSelectionDescription => _uninstallMode ?
			InsERT.Mox.QuickInstallerCore.Resources.TekstWyboruBazPrzyOdinstalowywaniu :
			InsERT.Mox.QuickInstallerCore.Resources.TekstWyboruBazPrzyInstalowaniu;

		public Visibility DatabaseSelectionVisibility
		{
			get => _databaseSelectionVisibility;
			set
			{
				_databaseSelectionVisibility = value;
				OnPropertyChanged(nameof(DatabaseSelectionVisibility));
			}
		}

		public string Informacja
		{
			get => _informacja;
			set
			{
				_informacja = value;
				OnPropertyChanged(nameof(Informacja));
				RefreshControl(MainTextBox);
			}
		}

		public bool IsExpandedUtworzSkrot => ShortcutCfgCollection?.Count() <= MaxIconsForExpandedView;

		public InstallationConfig Konfiguracja
		{
			get => _konfiguracja;
			set
			{
				_konfiguracja = value;
				_installer = new Installer(_konfiguracja, _uninstallMode, this);
				_installer.Info += OnInstallerInfo;

				Informacja = _konfiguracja.WelcomeText;
				UtworzSkrotNaPulpicie = _konfiguracja.CreateDesktopShortcut;
				UtworzSkrotWMenuStart = _konfiguracja.CreateMenuStartShortcut;

				if (Konfiguracja.ContainsPackage())
				{
					EtapInstalacji = Etap.Etap0_Start;
				}
				else
				{
					OdinstalujLink.HideAndDisable();
					EtapInstalacji = Etap.Etap6_Podsumowanie;
					Informacja += Environment.NewLine + Environment.NewLine;
				}

				MoznaDalej = true;
				OnPropertyChanged(nameof(Konfiguracja));
			}
		}

		public Visibility LicencjaVisibility
		{
			get => _licencjaVisibility;
			set
			{
				_licencjaVisibility = value;
				OnPropertyChanged(nameof(LicencjaVisibility));
			}
		}

		public bool MoznaDalej
		{
			get => _moznaDalej;
			set
			{
				_moznaDalej = value;
				OnPropertyChanged(nameof(MoznaDalej));
			}
		}

		public Visibility RestartInfoVisibility
		{
			get => _restartInfoVisibility;
			set
			{
				_restartInfoVisibility = value;
				OnPropertyChanged(nameof(RestartInfoVisibility));
			}
		}

		public string[] Scenarios
		{
			get => _scenarios;
			set
			{
				_scenarios = value;
				RefreshAfterSettingScenarios();
			}
		}

		public ObservableCollection<ShortcutCfg> ShortcutCfgCollection
		{
			get
			{
				if (_shortcutCfgCollection == null && Scenarios != null && Scenarios.Count() > 0)
				{
					_shortcutCfgCollection = new ObservableCollection<ShortcutCfg>(Scenarios.Select(sc =>
						new ShortcutCfg(sc,
								Konfiguracja?.Shortcuts?.FirstOrDefault(s => string.Equals(sc, s.NazwaScenariusza)),
								UtworzSkrotNaPulpicie ?? false,
								UtworzSkrotWMenuStart ?? false)
						));

					foreach (var shortcutCfg in ShortcutCfgCollection)
					{
						(shortcutCfg as INotifyPropertyChanged).PropertyChanged += ShortcutCfg_PropertyChanged;
					}
				}
				return _shortcutCfgCollection;
			}
		}

		public string TekstKlawisza
		{
			get
			{
				if (_etap == Etap.Etap6_Podsumowanie)
				{
					return "Zakończ";
				}
				else
				{
					return _etap == Etap.Etap1_Licencja ? "Zgadzam się" : "Dalej >";
				}
			}
		}

		public bool? UtworzSkrotNaPulpicie
		{
			get => _utworzSkrotNaPulpicie;
			set
			{
				_utworzSkrotNaPulpicie = value ?? !_utworzSkrotNaPulpicie;
				UstawWSkrotachUtworzNaPulpicie(_utworzSkrotNaPulpicie.Value);
				OnPropertyChanged(nameof(UtworzSkrotNaPulpicie));
			}
		}

		public string UtworzSkrotNaPulpicieTekst
		{
			get
			{
				if (ShortcutCfgCollection?.Count() == 1)
				{
					return _uninstallMode ? "Usuń skrót z pulpitu" : "Utwórz skrót na pulpicie";
				}
				else
				{
					return _uninstallMode ? "Usuń skróty z pulpitu" : "Utwórz skróty na pulpicie";
				}
			}
		}

		public Visibility UtworzSkrotVisibility
		{
			get => _utworzSkrotVisibility;
			set
			{
				_utworzSkrotVisibility = value;
				OnPropertyChanged(nameof(UtworzSkrotVisibility));
			}
		}

		public bool? UtworzSkrotWMenuStart
		{
			get => _utworzSkrotWMenuStart;
			set
			{
				_utworzSkrotWMenuStart = value ?? !_utworzSkrotWMenuStart;
				UstawWSkrotachUtworzWMenuStart(_utworzSkrotWMenuStart.Value);
				OnPropertyChanged(nameof(UtworzSkrotWMenuStart));
			}
		}

		public string UtworzSkrotWMenuStartTekst
		{
			get
			{
				if (ShortcutCfgCollection?.Count() == 1)
				{
					return _uninstallMode ? "Usuń skrót z menu start" : "Utwórz skrót w menu start";
				}
				else
				{
					return _uninstallMode ? "Usuń skróty z menu start" : "Utwórz skróty w menu start";
				}
			}
		}

		private Etap EtapInstalacji
		{
			get => _etap;
			set
			{
				_etap = value;
				OnPropertyChanged(nameof(Informacja));
				OnPropertyChanged(nameof(TekstKlawisza));
				OnPropertyChanged(nameof(MoznaDalej));
			}
		}

		/// <summary>
		/// Oto jest pytanie :)
		/// </summary>
		/// <returns>True - udana instalacja/dezinstalacji, False - wpp</returns>
		public bool DeployOrUndeploy()
		{
			var wybraneBazy = DatabaseCollection.Where(d => d.IsChecked).Select(d => d.Name);

			bool odinstalujZLaunchera = !DatabaseCollection.Any(d => !d.IsChecked && d.HasPackage);

			bool result = _installer.DeployOrUndeploy(wybraneBazy, odinstalujZLaunchera, ShortcutCfgCollection);

			return result;
		}

		public (string Package, string[] Scenarios) GetPackageInfo()
		{
			return _installer.GetPackageInfo();
		}

		private void CheckAllCommand_CanExecute(object sender, CanExecuteRoutedEventArgs e)
		{
			e.CanExecute = EtapInstalacji == Etap.Etap4_WyborBaz;
		}

		private void CheckAllCommand_Executed(object sender, ExecutedRoutedEventArgs e)
		{
			CheckedAll = !CheckedAll;
		}

		private void Dalej_Click(object sender, RoutedEventArgs e)
		{
			if (EtapInstalacji == Etap.Etap6_Podsumowanie)
			{
				Close();
				return;
			}
			else if (EtapInstalacji == Etap.Etap0_Start)
			{
				OdinstalujLink.HideAndDisable();
				Action action = delegate ()
				{
					using (var c = new WaitCursorInScope(this))
					{
						if (!_uninstallMode)
						{
							if (!string.IsNullOrWhiteSpace(Konfiguracja.LicencePath))
							{
								if (Konfiguracja.LicencePath.EndsWith(".rtf", StringComparison.CurrentCultureIgnoreCase) || Konfiguracja.LicenceText.StartsWith(@"{\rtf"))
								{
									byte[] documentBytes = Encoding.UTF8.GetBytes(Konfiguracja.LicenceText);
									using (var reader = new MemoryStream(documentBytes))
									{
										reader.Position = 0;
										Licencja.SelectAll();
										Licencja.Selection.Load(reader, DataFormats.Rtf);
									}
								}
								else
								{
									Licencja.AppendText(Konfiguracja.LicenceText.Replace("\r\n", "\r"));
								}

								Opcje.Visibility = Visibility.Collapsed;
								LicencjaVisibility = Visibility.Visible;
								EtapInstalacji++;
							}
							else
							{
								EtapInstalacji += 2;
								Dalej_Click(sender, e);
							}
						}
					}
				};
				_ = Dispatcher.BeginInvoke(action);
			}
			else if (EtapInstalacji == Etap.Etap1_Licencja)
			{
				LicencjaVisibility = Visibility.Collapsed;

				EtapInstalacji++;
				Dalej_Click(sender, e);
			}
			else if (EtapInstalacji == Etap.Etap2_PrzedInstalacja)
			{
				OdinstalujLink.HideAndDisable();

				Action action = delegate ()
				{
					if (_uninstallMode)
					{
						Informacja += "\n\nOdczytywanie informacji o pakiecie do odinstalowania...";
					}
					else
					{
						Informacja += "\n\nOdczytywanie informacji o pakiecie do zainstalowania...";
					}

					using (var c = new WaitCursorInScope(this))
					{
						var (package, scenarios) = GetPackageInfo();
						bool error = package == Installer.GetPackageInfoErrorString;

						if (!string.IsNullOrWhiteSpace(package) && !error)
						{
							Scenarios = scenarios;

							if (_uninstallMode)
							{
								Informacja = string.Format(
@"Konieczne będzie teraz zalogowanie do serwera SQL.
Następnie zostanie odinstalowany pakiet:
{0}."
									, package);
							}
							else
							{
								Informacja = string.Format(
@"Konieczne będzie teraz zalogowanie do serwera SQL.
Następnie zostanie zainstalowany pakiet:
{0}."
									, package);
							}

							if (Scenarios != null && Scenarios.Any())
							{
								Opcje.Visibility = Visibility.Visible;
								UtworzSkrotVisibility = Visibility.Visible;
							}
							EtapInstalacji++;
						}
						else
						{
							if (error)
							{
								Informacja = "Błąd pakietu.";
							}

							if (_uninstallMode)
							{
								Informacja += "\n\nOdinstalowywanie nie zostało ukończona.";
							}
							else
							{
								Informacja += "\n\nInstalacja nie została ukończona.";
							}

							EtapInstalacji = Etap.Etap6_Podsumowanie;
						}
					}
				};

				_ = Dispatcher.BeginInvoke(action);
			}
			else if (EtapInstalacji == Etap.Etap3_Informacja)
			{
				MoznaDalej = false;

				Action action = delegate ()
				{
					WczytajBazy();
				};
				EtapInstalacji = Etap.Etap4_WyborBaz;

				_ = Dispatcher.BeginInvoke(action);
			}
			else if (EtapInstalacji == Etap.Etap4_WyborBaz)
			{
				if (!DatabaseCollection.Any(d => d.IsChecked))
				{
					_ = MessageBox.Show(this, _uninstallMode ? InsERT.Mox.QuickInstallerCore.Resources.TekstBleduNieWybranoZadnejBazyDoOdinstalowaniaPakietu : InsERT.Mox.QuickInstallerCore.Resources.TekstBleduNieWybranoZadnejBazyDoInstalacjiPakietu,
						this.Title, MessageBoxButton.OK, MessageBoxImage.Error);
					return;
				}

				MoznaDalej = false;

				Action action = delegate ()
				{
					UtworzSkrotVisibility = Visibility.Hidden;
					DatabaseSelectionVisibility = Visibility.Collapsed;

					using (var c = new WaitCursorInScope(this))
					{
						bool result = DeployOrUndeploy();
						if (result)
						{
							RestartInfoVisibility = Konfiguracja.RequireRestart ? Visibility.Visible : Visibility.Hidden;
						}

						MoznaDalej = true;
						EtapInstalacji = Etap.Etap6_Podsumowanie;
					}
				};
				EtapInstalacji = Etap.Etap5_Instalacja;

				_ = Dispatcher.BeginInvoke(action);
			}
		}

		private void DatabaseFilter_TextChanged(object sender, TextChangedEventArgs e)
		{
			if (string.IsNullOrEmpty(DatabaseFilter.Text))
			{
				DatabaseFilterWatermark.Visibility = Visibility.Visible;
			}
			else
			{
				DatabaseFilterWatermark.Visibility = Visibility.Collapsed;
			}

			_databaseCollectionView.Refresh();
		}

		private void Odinstaluj_Click(object sender, RoutedEventArgs e)
		{
			_uninstallMode = true;
			_installer.ToggleUninstallMode(_uninstallMode);
			EtapInstalacji += 2;
			OnPropertyChanged(nameof(DatabaseSelectionDescription));
			Dalej_Click(sender, e);
		}

		private void Odrzucam_Click(object sender, RoutedEventArgs e)
		{
			var result = MessageBox.Show(this, "Czy na pewno przerwać instalację?", Title, MessageBoxButton.YesNo);
			if (result == MessageBoxResult.Yes)
			{
				LicencjaVisibility = Visibility.Collapsed;
				Informacja += "\n\nPrzerwano na życzenie użytkownika.\nInstalacja nie została ukończona.";
				EtapInstalacji = Etap.Etap6_Podsumowanie;
			}
		}

		private void OnInstallerInfo(object sender, MessageEventArgs e)
		{
			Informacja += "\n" + e.Message;
		}

		private void OnPropertyChanged(string name)
		{
			PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
		}

		private void RefreshAfterSettingScenarios()
		{
			_shortcutCfgCollection = null;
			OnPropertyChanged(nameof(ShortcutCfgCollection));
			OnPropertyChanged(nameof(IsExpandedUtworzSkrot));
			OnPropertyChanged(nameof(UtworzSkrotNaPulpicieTekst));
			OnPropertyChanged(nameof(UtworzSkrotWMenuStartTekst));
		}

		private void RefreshControl(DispatcherObject o)
		{
			o.Dispatcher.Invoke(new Action(() => { }), DispatcherPriority.Render);
		}

		private void RemoveFilterButton_Click(object sender, RoutedEventArgs e)
		{
			DatabaseFilter.Text = string.Empty;
		}

		private void ScrollViewer_ScrollChanged(object sender, ScrollChangedEventArgs e)
		{
			// User scroll event : set or unset autoscroll mode
			if (e.ExtentHeightChange == 0)
			{   // Content unchanged : user scroll event
				if (ScrollViewer.VerticalOffset == ScrollViewer.ScrollableHeight)
				{   // Scroll bar is in bottom
					// Set autoscroll mode
					_autoScroll = true;
				}
				else
				{   // Scroll bar isn't in bottom
					// Unset autoscroll mode
					_autoScroll = false;
				}
			}

			// Content scroll event : autoscroll eventually
			if (_autoScroll && e.ExtentHeightChange != 0)
			{   // Content changed and autoscroll mode set
				// Autoscroll
				ScrollViewer.ScrollToVerticalOffset(ScrollViewer.ExtentHeight);
			}
		}

		private void ShortcutCfg_PropertyChanged(object sender, PropertyChangedEventArgs e)
		{
			if (e.PropertyName == nameof(ShortcutCfg.CreateOnDesktop))
			{
				if (_flagaUstawianieUtworzSkrotNaPulpicie)
				{
					return;
				}

				if (ShortcutCfgCollection.All(s => s.CreateOnDesktop))
				{
					_utworzSkrotNaPulpicie = true;
				}
				else if (ShortcutCfgCollection.All(s => s.CreateOnDesktop == false))
				{
					_utworzSkrotNaPulpicie = false;
				}
				else
				{
					_utworzSkrotNaPulpicie = null;
				}

				OnPropertyChanged(nameof(UtworzSkrotNaPulpicie));
			}
			else if (e.PropertyName == nameof(ShortcutCfg.CreateInStartMenu))
			{
				if (_flagaUstawianieUtworzSkrotWMenuStart)
				{
					return;
				}

				if (ShortcutCfgCollection.All(s => s.CreateInStartMenu))
				{
					_utworzSkrotWMenuStart = true;
				}
				else if (ShortcutCfgCollection.All(s => s.CreateInStartMenu == false))
				{
					_utworzSkrotWMenuStart = false;
				}
				else
				{
					_utworzSkrotWMenuStart = null;
				}

				OnPropertyChanged(nameof(UtworzSkrotWMenuStart));
			}
		}

		private bool ShouldShowDatabase(object input)
		{
			string filterText = DatabaseFilter.Text;
			if (input is DatabaseInfoWrapper database && !string.IsNullOrEmpty(filterText))
			{
				return database.Name.IndexOf(filterText, StringComparison.InvariantCultureIgnoreCase) >= 0
					|| database.Version.ToString().IndexOf(filterText, StringComparison.InvariantCultureIgnoreCase) >= 0;
			}
			return true;
		}

		private void UstawWSkrotachUtworzNaPulpicie(bool utworzSkrotNaPulpicie)
		{
			try
			{
				_flagaUstawianieUtworzSkrotNaPulpicie = true;

				if (ShortcutCfgCollection?.Any() == true)
				{
					foreach (var shortcutCfg in ShortcutCfgCollection)
					{
						shortcutCfg.CreateOnDesktop = utworzSkrotNaPulpicie;
					}
				}
			}
			finally
			{
				_flagaUstawianieUtworzSkrotNaPulpicie = false;
			}
		}

		private void UstawWSkrotachUtworzWMenuStart(bool utworzSkrotWMenuStart)
		{
			try
			{
				_flagaUstawianieUtworzSkrotWMenuStart = true;

				if (ShortcutCfgCollection?.Any() == true)
				{
					foreach (var shortcutCfg in ShortcutCfgCollection)
					{
						shortcutCfg.CreateInStartMenu = utworzSkrotWMenuStart;
					}
				}
			}
			finally
			{
				_flagaUstawianieUtworzSkrotWMenuStart = false;
			}
		}

		private void WczytajBazy()
		{
			Informacja = _uninstallMode ? $"Odinstalowywanie {Konfiguracja.PackagePath} - wybór baz\n" : $"Instalacja {Konfiguracja.PackagePath} - wybór baz\n";

			UtworzSkrotVisibility = Visibility.Hidden;

			using (var c = new WaitCursorInScope(this))
			{
				var databases = _installer.WczytajBazy();

				if (databases != null && databases.Any())
				{
					foreach (var database in databases)
					{
						DatabaseCollection.Add(database);
					}

					Opcje.Visibility = Visibility.Collapsed;
					DatabaseSelectionVisibility = Visibility.Visible;
					_ = DatabaseList.Focus();
				}
				else
				{
					Informacja += string.Concat("\n\nBrak dostępnych baz lub użytkownik SQL nie posiada do nich uprawnień.\n",
						_uninstallMode ? "Odinstalowywanie nie zostało ukończone." : "Instalacja nie została ukończona.");
					EtapInstalacji = Etap.Etap6_Podsumowanie;
				}

				MoznaDalej = true;
			}
		}
	}
}