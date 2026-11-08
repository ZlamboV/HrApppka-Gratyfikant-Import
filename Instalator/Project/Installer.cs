using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;

namespace InsERT.Mox.QuickInstallerCore
{
	internal sealed class Installer
	{
		internal const string GetPackageInfoErrorString = "ERROR";

		private readonly InstallationConfig _konfiguracja;
		private bool _uninstallMode;
		private readonly PackageDeployerAdapter _adapter;

		internal Installer(InstallationConfig cfg, bool uninstall, Window window)
		{
			_konfiguracja = cfg ?? throw new ArgumentNullException(nameof(cfg));
			_uninstallMode = uninstall;
			_adapter = new PackageDeployerAdapter(window, _konfiguracja.DetermineProduct(), _uninstallMode);
		}

		public void ToggleUninstallMode(bool newValue)
		{
			_uninstallMode = newValue;
			_adapter.ToggleUninstallMode(newValue);
		}

		public event EventHandler<MessageEventArgs> Info;

		internal bool DeployOrUndeploy(IEnumerable<string> selectedDatabases, bool removeFromDistributionDatabase, IEnumerable<ShortcutCfg> shortcutsToCreateOrRemove)
		{
			if (selectedDatabases == null || !selectedDatabases.Any())
			{
				throw new InvalidOperationException("Nie wybrano żadnej bazy");
			}

			_adapter.DeployerStatusChanged += Accessor_DeployerStatusChanged;

			bool result;
			string pakiet = TemporaryPayload.GetItem(_konfiguracja.PackagePath);

			var launcherSupport = new LauncherSupport(_konfiguracja.DetermineProduct());
			if (_uninstallMode)
			{
				result = _adapter.Undeploy(pakiet, selectedDatabases, removeFromDistributionDatabase);

				//usuń skróty
				if (result && shortcutsToCreateOrRemove?.Any(skrot => skrot.CreateOnDesktop || skrot.CreateInStartMenu) == true)
				{
					string shortcutMessage = launcherSupport.RemoveShortcuts(shortcutsToCreateOrRemove.Where(skrot => skrot.CreateOnDesktop || skrot.CreateInStartMenu),
						Path.GetFileNameWithoutExtension(_konfiguracja.PackagePath));
					OnInfo(shortcutMessage);
				}
			}
			else
			{
				result = _adapter.Deploy(pakiet, selectedDatabases);

				//utwórz skróty
				if (result && shortcutsToCreateOrRemove?.Any(skrot => skrot.CreateOnDesktop || skrot.CreateInStartMenu) == true)
				{
					string shortcutMessage = launcherSupport.CreateShortcuts(shortcutsToCreateOrRemove.Where(skrot => skrot.CreateOnDesktop || skrot.CreateInStartMenu),
						Path.GetFileNameWithoutExtension(_konfiguracja.PackagePath));
					OnInfo(shortcutMessage);
				}
			}

			_adapter.DeployerStatusChanged -= Accessor_DeployerStatusChanged;

			if (result)
			{
				OnInfo(_uninstallMode ? "Odinstalowywanie zakończone." : "Instalacja zakończona.");
			}
			else
			{
				OnInfo(_uninstallMode ? "Odinstalowywanie nie zostało ukończone." : "Instalacja nie została ukończona.");
			}

			return result;
		}

		internal (string Package, string[] Scenarios) GetPackageInfo()
		{
			string pakiet = TemporaryPayload.GetItem(_konfiguracja.PackagePath);
			return _adapter.GetPackageInfo(pakiet);
		}

		internal void SetDeployParameters(DeployParameters deployParameters)
		{
			_adapter.DeployerStatusChanged += Accessor_DeployerStatusChanged;
			_adapter.SetServerConnectionParameters(deployParameters);
			_adapter.DeployerStatusChanged -= Accessor_DeployerStatusChanged;
		}

		internal IEnumerable<DatabaseInfoWrapper> WczytajBazy()
		{
			string pakiet = TemporaryPayload.GetItem(_konfiguracja.PackagePath);

			_adapter.DeployerStatusChanged += Accessor_DeployerStatusChanged;

			var databases = _adapter.TryGetDatabaseInfos(pakiet);

			_adapter.DeployerStatusChanged -= Accessor_DeployerStatusChanged;

			return databases;
		}

		private void Accessor_DeployerStatusChanged(object sender, MessageEventArgs e)
		{
			OnInfo(e.Message);
		}

		private void OnInfo(string message)
		{
			Info?.Invoke(this, new MessageEventArgs() { Message = message });
		}
	}
}