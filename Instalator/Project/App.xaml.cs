using System;
using System.IO;
using System.Linq;
using System.Windows;

namespace InsERT.Mox.QuickInstallerCore
{
	public partial class App : Application
	{
		private static bool _consoleMode = false;
		private static bool _noConsolePause = false;

		private static void AttachOrAllocConsole()
		{
			if (!NativeMethods.AttachConsole(NativeMethods.ATTACH_PARENT_PROCESS))
			{
				_ = NativeMethods.AllocConsole();
			}
		}

		private static void FreeConsole()
		{
			NativeMethods.FreeConsole();
		}

		private static void InitializeLogger()
		{
			Logger.ConfigureLog(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "InsERT", "QuickInstaller", "QuickInstaller.log"));
		}

		[STAThread]
		private static void Main(string[] args)
		{
			InitializeLogger();

			_consoleMode = args.Contains("/bezUI");
			_noConsolePause = args.Contains("/bezPauzy");

			try
			{
				if (_consoleMode)
				{
					AttachOrAllocConsole();
				}

				string configurationFileName = "konfiguracja.cfg";
				if (!Payload.HasItem(configurationFileName))
				{
					string komunikat = "Nie znaleziono pliku z konfiguracją instalatora.";
					ShowErrorMessage(komunikat);
					return;
				}

				bool uninstall = args.Contains("/odinstaluj");
				var cfg = InstallationConfig.LoadOrCreate(configurationFileName, uninstall);

				if (cfg.ContainsPackage()) // zawiera pakiet dystrybucyjny
				{
					if (!new LauncherSupport(cfg.DetermineProduct()).IsProductInstalled()) //działamy zgodnie z docelową platformą
					{
						string komunikat = cfg.KiP2
							? "Program InsERT Kadry i Płace 2 musi być zainstalowany."
							: "InsERT nexo musi być zainstalowane.";
						ShowErrorMessage(komunikat);
						return;
					}
				}
				else // nie zawiera pakietu dystrybucyjnego
				{
					string komunikat = "Nie znaleziono pakietu dystrybucyjnego z rozwiązaniem własnym.";
					ShowErrorMessage(komunikat);
					return;
				}

				if (_consoleMode)
				{
					ZainstalujBezUI(args, cfg, uninstall);
				}
				else
				{
					var app = new App();
					UnhandledExceptions.HandleInCurrentDomain();
					UnhandledExceptions.HandleInWpfApp(app);

					var window = new MainWindow(uninstall)
					{
						Konfiguracja = cfg
					};
					UruchomInstalator(app, window);
				}
			}
			catch (Exception e)
			{				
				ShowErrorMessage("Wystąpił błąd: " + e.Message);
				Logger.LogException(e);
			}
			finally
			{
				TemporaryPayload.Delete();
				if (_consoleMode)
				{
					if (!_noConsolePause)
					{
						PauseConsole();
					}
					FreeConsole();
				}
			}
		}

		private static void PauseConsole()
		{
			Console.Write("Press Enter to continue . . . ");
			_ = Console.ReadLine();
		}

		private static void ShowErrorMessage(string komunikat)
		{
			if (_consoleMode)
			{
				Console.WriteLine(komunikat);
			}
			else
			{
				_ = MessageBox.Show(komunikat, "QuickInstaller", MessageBoxButton.OK, MessageBoxImage.Error);
			}

			Logger.LogInformation(komunikat);
		}

		private static void UruchomInstalator(App app, MainWindow window)
		{
			app.MainWindow = window;
			app.MainWindow.Show();
			_ = app.Run();
		}

		private static void ZainstalujBezUI(string[] args, InstallationConfig cfg, bool uninstallMode)
		{
			WindowlessInstallation.Execute(args, cfg, uninstallMode);
		}
	}
}