using InsERT.Mox.QuickInstallerCore.Xml;
using System.ComponentModel;
using System.Windows.Media.Imaging;

namespace InsERT.Mox.QuickInstallerCore
{
	public class InstallationConfig : INotifyPropertyChanged
	{
		private bool _alfa = false;
		private Xml.AssemblyInfo _assemblyInfo;
		private bool _createDesktopShortcut = true;
		private bool _createStartMenuShortcut = true;
		private bool _kip2 = false;
		private string _licencePath = string.Empty;
		private string _logoPath;
		private string _outputFilePath;
		private string _packagePath;
		private string _providerInfo;
		private bool _requireRestart = false;
		private string _title;
		private string _welcomeText;

		internal InstallationConfig()
		{
			PackagePath = null;
			Title = Resources.Tytul;
			WelcomeText = Resources.TekstPowitalny;
			ProviderInfo = Resources.DostawcaInfo;
			Logo = null;
			AssemblyInfo = new Xml.AssemblyInfo();
		}

		public event PropertyChangedEventHandler PropertyChanged;

		public bool Alfa
		{
			get => _alfa;
			set
			{
				_alfa = value;
				OnPropertyChanged(nameof(Alfa));
			}
		}

		public string AssemblyCompany
		{
			get => _assemblyInfo.Company;
			set
			{
				_assemblyInfo.Company = value;
				OnPropertyChanged(nameof(AssemblyCompany));
				OnPropertyChanged(nameof(HasAssemblyInfo));
			}
		}

		public string AssemblyCopyright
		{
			get => _assemblyInfo.Copyright;
			set
			{
				_assemblyInfo.Copyright = value;
				OnPropertyChanged(nameof(AssemblyCopyright));
				OnPropertyChanged(nameof(HasAssemblyInfo));
			}
		}

		public string AssemblyFileVersion
		{
			get => _assemblyInfo.FileVersion;
			set
			{
				_assemblyInfo.FileVersion = value;
				OnPropertyChanged(nameof(AssemblyFileVersion));
				OnPropertyChanged(nameof(HasAssemblyInfo));
			}
		}

		internal Xml.AssemblyInfo AssemblyInfo
		{
			get => _assemblyInfo;
			set
			{
				_assemblyInfo = value;
				OnPropertyChanged(nameof(AssemblyInfo));
			}
		}

		public string AssemblyProduct
		{
			get => _assemblyInfo.Product;
			set
			{
				_assemblyInfo.Product = value;
				OnPropertyChanged(nameof(AssemblyProduct));
				OnPropertyChanged(nameof(HasAssemblyInfo));
			}
		}

		public string AssemblyVersion
		{
			get => _assemblyInfo.Version;
			set
			{
				_assemblyInfo.Version = value;
				OnPropertyChanged(nameof(AssemblyVersion));
				OnPropertyChanged(nameof(HasAssemblyInfo));
			}
		}

		public bool CreateDesktopShortcut
		{
			get => _createDesktopShortcut;
			set
			{
				_createDesktopShortcut = value;
				OnPropertyChanged(nameof(CreateDesktopShortcut));
			}
		}

		public bool CreateMenuStartShortcut
		{
			get => _createStartMenuShortcut;
			set
			{
				_createStartMenuShortcut = value;
				OnPropertyChanged(nameof(CreateMenuStartShortcut));
			}
		}

		public bool HasAssemblyInfo => !string.IsNullOrWhiteSpace(AssemblyCompany)
			|| !string.IsNullOrWhiteSpace(AssemblyCopyright)
			|| !string.IsNullOrWhiteSpace(AssemblyFileVersion)
			|| !string.IsNullOrWhiteSpace(AssemblyProduct)
			|| !string.IsNullOrWhiteSpace(AssemblyVersion);

		public bool KiP2
		{
			get => _kip2;
			set
			{
				_kip2 = value;
				OnPropertyChanged(nameof(KiP2));
			}
		}

		public string LicencePath
		{
			get => _licencePath;
			set
			{
				_licencePath = value;
				OnPropertyChanged(nameof(LicencePath));
			}
		}

		public string LicenceText { get; private set; }

		public BitmapImage Logo { get; private set; }

		public string LogoPath
		{
			get => _logoPath;
			set
			{
				_logoPath = value;
				OnPropertyChanged(nameof(LogoPath));
			}
		}

		public string PackagePath
		{
			get => _packagePath;
			set
			{
				_packagePath = value;
				OnPropertyChanged(nameof(PackagePath));
			}
		}

		public string PlikWynikowy
		{
			get => _outputFilePath;
			set
			{
				_outputFilePath = value;
				OnPropertyChanged(nameof(PlikWynikowy));
			}
		}

		public string ProviderInfo
		{
			get => _providerInfo;
			set
			{
				_providerInfo = value;
				OnPropertyChanged(nameof(ProviderInfo));
			}
		}

		public bool RequireRestart
		{
			get => _requireRestart;
			set
			{
				_requireRestart = value;
				OnPropertyChanged(nameof(RequireRestart));
			}
		}

		public Skrot[] Shortcuts { get; set; }

		public string Title
		{
			get => _title;
			set
			{
				_title = value;
				OnPropertyChanged(nameof(Title));
			}
		}

		public string WelcomeText
		{
			get => _welcomeText;
			set
			{
				_welcomeText = value;
				OnPropertyChanged(nameof(WelcomeText));
			}
		}

		public static InstallationConfig LoadOrCreate(string path, bool uninstall)
		{
			var cfg = new InstallationConfig();
			var cfgStream = Payload.GetItemStream(path);
			if (cfgStream != null)
			{
				InstallationCfgXmlSerialization.LoadXml(cfgStream, cfg);

				cfg.LoadLicenceFile(cfg.LicencePath);
				cfg.LoadLogoFile(cfg.LogoPath);
				cfgStream.Dispose();
			}

			if (uninstall)
			{
				cfg.WelcomeText = Resources.TekstPowitalnyUninstall;
			}

			if (!cfg.ContainsPackage())
			{
				cfg.WelcomeText = Resources.TekstPowitalnyBrakPakietu;
			}

			return cfg;
		}

		public bool ContainsPackage()
		{
			return !string.IsNullOrWhiteSpace(PackagePath) && Payload.HasItem(PackagePath);
		}

		public void LoadLicenceFile()
		{
			LoadLicenceText(LicencePath);
		}

		public void LoadLogoFile()
		{
			LoadBitmapImage(LogoPath);
		}

		internal static BitmapImage LoadBitmapImage(string plik)
		{
			if (Payload.HasItem(plik))
			{
				var stream = Payload.GetItemStream(plik);
				if (stream != null)
				{
					var image = new BitmapImage();
					image.BeginInit();
					image.CacheOption = BitmapCacheOption.OnLoad;
					image.StreamSource = stream;
					image.EndInit();
					return image;
				}
			}
			return null;
		}

		private static string LoadLicenceText(string plik)
		{
			return Payload.GetItemAsText(plik);
		}

		private void LoadLicenceFile(string newLicenceFile)
		{
			if (Payload.HasItem(newLicenceFile))
			{
				string licenceText = LoadLicenceText(newLicenceFile);
				LicenceText = licenceText;
			}
		}

		private void LoadLogoFile(string newLogoFile)
		{
			if (Payload.HasItem(newLogoFile))
			{
				var logo = LoadBitmapImage(newLogoFile);
				Logo = logo;
			}
		}

		private void OnPropertyChanged(string name)
		{
			PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
		}
	}
}