using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Loader;
using System.Windows;

namespace InsERT.Mox.QuickInstallerCore
{
	internal class PackageDeployerAdapter
	{
        private const int DatabaseInfoPartsCount = 6;

        private readonly string _nexoInstallationPath;
		private readonly string _nexoAlfaInstallationPath;
		private readonly Product _product;
		private readonly Window _window;
		private object _deployer;
		private Assembly _packageDeployer = null;
		private Type _typeDeployer;

		private bool _uninstallMode;

		internal PackageDeployerAdapter(Window window, Product product, bool uninstallMode)
		{
			_window = window;
			_product = product;
			_uninstallMode = uninstallMode;
			_nexoAlfaInstallationPath = new LauncherSupport(Product.NexoAlfa).GetInstallPath();
			_nexoInstallationPath = new LauncherSupport(Product.Nexo).GetInstallPath();
		}

		public event EventHandler<MessageEventArgs> DeployerStatusChanged;

		private object Deployer
		{
			get
			{
				if (_deployer == null)
				{
					var constructorInfo = TypeDeployer.GetConstructors().Single();					
					_deployer = constructorInfo.Invoke([_window, _product.ToString()]); 
				}
				return _deployer;
			}
		}

		private Assembly PackageDeployer
		{
			get
			{
				if (_packageDeployer == null)
				{
					AppDomain.CurrentDomain.AssemblyResolve += LoadAssembly;
					_packageDeployer = Assembly.Load(Resources.PackageDeployer);
				}
				return _packageDeployer;
			}
		}

		private Type TypeDeployer
		{
			get
			{
				if (_typeDeployer == null)
				{
					_typeDeployer = PackageDeployer.GetType("PackageDeployer.Deployer");
				}
				return _typeDeployer;
			}
		}

		internal bool Deploy(string packageFilePath, IEnumerable<string> selectedDatabases)
		{
			var eventSource = (INotifyPropertyChanged)Deployer;
			eventSource.PropertyChanged += Deployer_PropertyChanged;
			bool result = (bool)TypeDeployer.InvokeMember("Deploy",
					BindingFlags.Public | BindingFlags.Instance | BindingFlags.InvokeMethod,
					null,
					Deployer,
					[packageFilePath, selectedDatabases]);
			eventSource.PropertyChanged -= Deployer_PropertyChanged;
			return result;
		}

		internal (string Package, string[] Scenarios) GetPackageInfo(string packageFilePath)
		{
			string pakiet = TemporaryPayload.GetItem(packageFilePath);
			var result = (Tuple<string, string[]>)PackageDeployer.GetType("PackageDeployer.Deployer")
					.InvokeMember("GetPackageInfo",
									BindingFlags.Public |
									BindingFlags.Static |
									BindingFlags.InvokeMethod,
									null,
									null,
									new object[] { pakiet });
			return (Package: result.Item1, Scenarios: result.Item2);
		}

		internal void SetServerConnectionParameters(DeployParameters deployParameters)
		{
			ArgumentNullException.ThrowIfNull(deployParameters);

			var eventSource = (INotifyPropertyChanged)Deployer;
			eventSource.PropertyChanged += Deployer_PropertyChanged;

			TypeDeployer.InvokeMember("SetServerConnectionParameters",
				BindingFlags.Public | BindingFlags.Instance | BindingFlags.InvokeMethod,
				null,
				Deployer,
				[
					deployParameters.SqlServer,
					deployParameters.WindowsAuthentication,
					deployParameters.User,
					deployParameters.Password
				]);

			eventSource.PropertyChanged -= Deployer_PropertyChanged;
		}

		internal void ToggleUninstallMode(bool newValue)
		{
			_uninstallMode = newValue;
		}

		internal IEnumerable<DatabaseInfoWrapper> TryGetDatabaseInfos(string packageFilePath)
		{
			var eventSource = (INotifyPropertyChanged)Deployer;
			eventSource.PropertyChanged += Deployer_PropertyChanged;
			var result = (IEnumerable<string>)TypeDeployer.InvokeMember("TryGetDatabaseInfos",
				BindingFlags.Public | BindingFlags.Instance | BindingFlags.InvokeMethod,
				null,
				Deployer,
				[packageFilePath, _uninstallMode]);
			eventSource.PropertyChanged -= Deployer_PropertyChanged;

			var databases = new List<DatabaseInfoWrapper>();
			if (result != null && result.Any())
			{
				foreach (string db in result)
				{
					var wrappedDatabase = ParseDatabase(db);
					if (wrappedDatabase != null)
					{
						databases.Add(wrappedDatabase);
					}
				}
			}
			return databases;
		}

		internal bool Undeploy(string packageFilePath, IEnumerable<string> selectedDatabases, bool removeFromDistributionDatabase)
		{
			var eventSource = (INotifyPropertyChanged)Deployer;
			eventSource.PropertyChanged += Deployer_PropertyChanged;
			bool result = (bool)TypeDeployer.InvokeMember(
					"Undeploy",
					BindingFlags.Public | BindingFlags.Instance | BindingFlags.InvokeMethod,
					null,
					Deployer,
					[packageFilePath, selectedDatabases, removeFromDistributionDatabase]);
			eventSource.PropertyChanged -= Deployer_PropertyChanged;
			return result;
		}

		private void Deployer_PropertyChanged(object sender, PropertyChangedEventArgs e)
		{
			if (e.PropertyName == "Status")
			{
				OnDeployerStatusChanged((string)sender.GetType().GetProperty("Status").GetValue(sender, null));
			}
		}

		private Assembly LoadAssembly(object sender, ResolveEventArgs args)
		{
			/* Kiedyś mieliśmy tu ładowanie InsLauncher.exe, ale teraz już nie ma sensu tego robić.
			 * 
			 * Jeśli InsLauncher.exe targetuje .NET Framework, to załadowanie go do procesu net8 skończy się błędem:
			 * System.TypeLoadException: Could not load type 'System.Security.Principal.WindowsImpersonationContext' 
			 * from assembly 'mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'. 
			 * 
			 * Jeśli InsLauncher.exe targetuje .NET 8, to jest tylko app hostem i załadowanie go jako assembly skończy się błędem. 
			 * Należy ładować InsLauncher.dll. Znajdziemy go w folderze instalacji NexoAlfa, a w przyszłości także w folderze Nexo. 
			 *  
			 */

			if (args.Name.Contains(".resources, "))
			{
				return null;
			}

			var installationPaths = new string[] { _nexoAlfaInstallationPath, _nexoInstallationPath };

			foreach (var installationPath in installationPaths)
			{
				if (!string.IsNullOrEmpty(installationPath))
				{
					string assemblyName = new AssemblyName(args.Name).Name + ".dll";
					string assemblyPathDll = Path.Combine(installationPath, assemblyName);
					if (File.Exists(assemblyPathDll))
					{
						return AssemblyLoadContext.Default.LoadFromAssemblyPath(assemblyPathDll);
					}
				}
			}

			return null;
		}

		private void OnDeployerStatusChanged(string message)
		{
			DeployerStatusChanged?.Invoke(this, new MessageEventArgs() { Message = message });
		}

		private DatabaseInfoWrapper ParseDatabase(string databaseInfo)
		{
			if (!string.IsNullOrEmpty(databaseInfo))
			{
				string[] parts = databaseInfo.Split(';');
				if (parts.Length == DatabaseInfoPartsCount
                    && !string.IsNullOrEmpty(parts[0])
					&& bool.TryParse(parts[1], out bool hasPackage)
					&& bool.TryParse(parts[2], out bool hasPackageInDifferentVersion)
                    && parts[3] is not null
                    && bool.TryParse(parts[4], out bool hasRequiredPackages)
					&& Version.TryParse(parts[5], out var databaseVersion))
				{
					return new DatabaseInfoWrapper(parts[0], hasPackage, hasPackageInDifferentVersion, parts[3], hasRequiredPackages, databaseVersion, _uninstallMode);
				}
			}
			return null;
		}
	}
}