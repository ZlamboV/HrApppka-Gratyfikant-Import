using System;
using System.ComponentModel;

namespace InsERT.Mox.QuickInstallerCore
{
	public class DatabaseInfoWrapper : INotifyPropertyChanged
	{
		private bool _isChecked;

		public DatabaseInfoWrapper(string name, bool hasPackage, bool hasPackageInDifferentVersion, string packageVersion, bool hasRequiredPacakges, Version version, bool uninstallMode)
		{
			Name = name;
			HasPackage = hasPackage;
			HasPackageInDifferentVersion = hasPackageInDifferentVersion;
			PackageVersion = packageVersion;
			HasRequiredPackage = hasRequiredPacakges;
			Version = version;

			if (uninstallMode)
			{
				_isChecked = IsEnabled = HasPackage;
				Description = HasPackage ? string.Empty : 
					HasPackageInDifferentVersion ? 
						string.Format(Resources.TekstZawieraRozwiazanieWWersji, PackageVersion) : 
						Resources.TekstNieZawieraRozwiazania;
			}
			else
			{
				_isChecked = IsEnabled = HasRequiredPackage;
				if (HasPackage)
				{
					Description = Resources.TekstZawieraRozwiazanie;
				}
				else if(HasPackageInDifferentVersion)
				{
					Description = string.Format(Resources.TekstZawieraRozwiazanieWWersji, PackageVersion);
                }
                else if (!HasRequiredPackage)
				{
					Description = Resources.TekstNieZawieraWymaganychPakietow;
				}
				else
				{
					Description = string.Empty;
				}
			}
		}

		public event PropertyChangedEventHandler PropertyChanged;

		public string Description { get; private set; }

		public bool HasPackage { get; private set; }

		public bool HasPackageInDifferentVersion { get; private set; }

        public string PackageVersion { get; private set; }

        public bool HasRequiredPackage { get; private set; }

		public bool IsChecked
		{
			get => _isChecked;
			set
			{
				_isChecked = value;
				OnPropertyChanged(nameof(IsChecked));
			}
		}

		public bool IsEnabled { get; private set; }
		public string Name { get; private set; }
		public Version Version { get; private set; }

		protected void OnPropertyChanged(string name = null)
		{
			PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
		}
	}
}