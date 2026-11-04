using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace InsERT.Mox.QuickInstallerCore
{
	internal static class WindowlessInstallation
	{
		internal static void Execute(string[] args, InstallationConfig cfg, bool uninstallMode)
		{
			var deploymentParameters = GetDeployParameters(args);
			if (deploymentParameters == null)
			{
				return;
			}

			var installer = new Installer(cfg, uninstallMode, window: null);
			installer.Info += OnInstallerInfo_ConsoleOnly;

			var (package, scenarios) = installer.GetPackageInfo();
			if (package != Installer.GetPackageInfoErrorString)
			{
				var shortcuts = scenarios.Length == 0
					? scenarios.Select(sc =>
						new ShortcutCfg(sc,
								cfg.Shortcuts?.FirstOrDefault(s => string.Equals(sc, s.NazwaScenariusza)),
								cfg.CreateDesktopShortcut,
								cfg.CreateMenuStartShortcut))
					: Enumerable.Empty<ShortcutCfg>();

				installer.SetDeployParameters(deploymentParameters);
				bool removeFromDistributionDatabase = ShouldRemovePackageFromDistributionDatabase(deploymentParameters.Databases, installer);
				installer.DeployOrUndeploy(deploymentParameters.Databases, removeFromDistributionDatabase, shortcuts);
			}

			installer.Info -= OnInstallerInfo_ConsoleOnly;
		}

		private static string ExtractArgValue(string[] args, string name)
		{
			string argValue = null;
			string argFullTag = "/" + name + ":";
			string argFull = Array.Find(args, a => a.StartsWith(argFullTag, StringComparison.CurrentCultureIgnoreCase));
			if (argFull != null)
			{
				argValue = argFull.Substring(argFullTag.Length).Trim('\"');
			}

			return argValue;
		}

		private static DeployParameters GetDeployParameters(string[] args)
		{
			string deploymentParametersPath = ExtractArgValue(args, "parametryInstalacji");
			if (string.IsNullOrWhiteSpace(deploymentParametersPath))
			{
				ShowErrorMessage("Nie podano ścieżki do pliku zawierającego parametry instalacji.");
				return null;
			}

			if (!File.Exists(deploymentParametersPath))
			{
				ShowErrorMessage($@"Plik o podanej ścieżce nie istnieje: ""{deploymentParametersPath}"".");
				return null;
			}

			string[] parameterLines = File.ReadAllLines(deploymentParametersPath);
			var deployParameters = DeployParametersParser.Parse(parameterLines);

			if (!deployParameters.Validate(out string errorMessage))
			{
				ShowErrorMessage(errorMessage);
				return null;
			}

			return deployParameters;
		}

		private static void OnInstallerInfo_ConsoleOnly(object sender, MessageEventArgs e)
		{
			Console.WriteLine(e.Message);
		}

		private static bool ShouldRemovePackageFromDistributionDatabase(List<string> selectedDatabases, Installer installer)
		{
			// jeśli istnieje jakaś baza, która ma ten pakiet, ale nie jest jedną z wybranych do deinstalacji,
			// to znaczy, że nie możemy usunąć pakietu z bazy dystrybucyjnej
			return !installer.WczytajBazy().Any(x => x.HasPackage && !selectedDatabases.Contains(x.Name));
		}

		private static void ShowErrorMessage(string message)
		{
			Console.WriteLine(message);
		}
	}
}