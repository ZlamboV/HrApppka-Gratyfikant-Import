using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace InsERT.Mox.QuickInstallerCore
{
	internal class LauncherSupport
	{
		private const string InsLauncherExe = "InsLauncher.exe";
				
		private readonly ProductInfo _productInfo;
		private string _icoFolderPath;

		public LauncherSupport(Product product)
		{
			_productInfo = KnownProducts.GetProductInfo(product);
		}

		public string CreateShortcuts(IEnumerable<ShortcutCfg> shortcutCfgs, string packageName)
		{
			if (!IsProductInstalled())
			{
				return $"Musi być zainstalowany program {GetProductName()}. Skrót nie został utworzony.";
			}
			if (string.IsNullOrWhiteSpace(packageName))
			{
				return "Nie podano nazwy pakietu potrzebnej do utworzeniu folderu zawierającego ikony rozszerzenia.";
			}

			var sb = new StringBuilder();

			foreach (var shortcutCfg in shortcutCfgs)
			{
				string target = GetInsLauncherFullPath();
				string arg = $"\"{GetShortcutProductName()}/{shortcutCfg.ScenarioName}\"";

				string plikIkony;
				if (string.IsNullOrWhiteSpace(shortcutCfg.Shortcut?.PlikIkony) || !Payload.HasItem(shortcutCfg.Shortcut.PlikIkony))
				{
					plikIkony = target; //jeżeli nie podano ikony to bierzemy ikonę insLaunchera
				}
				else
				{
					string unpackedIconPath = TemporaryPayload.GetItem(shortcutCfg.Shortcut.PlikIkony);
					plikIkony = Path.Combine(GetIcoFolderPath(packageName), Path.GetFileName(unpackedIconPath));
					if (!File.Exists(plikIkony))
					{
						File.Copy(unpackedIconPath, plikIkony);
					}
				}

				if (shortcutCfg.CreateOnDesktop && Shortcut.CreateDesktopShortcut(shortcutCfg.ScenarioName,
						shortcutCfg.Shortcut?.Opis ?? shortcutCfg.ScenarioName, target, plikIkony, 0, arg))
				{
					sb.AppendLine($"Skrót dla \"{shortcutCfg.ScenarioName}\" został utworzony na pulpicie.");
				}

				if (shortcutCfg.CreateInStartMenu && Shortcut.CreateStartMenuShortcut(GetShortcutProductName(),
						shortcutCfg.ScenarioName, shortcutCfg.Shortcut?.Opis ?? shortcutCfg.ScenarioName, target, plikIkony, 0, arg))
				{
					sb.AppendLine($"Skrót dla \"{shortcutCfg.ScenarioName}\" został utworzony w menu Start.");
				}
			}

			return sb.ToString();
		}

		public string GetInsLauncherFullPath()
		{
			return Path.Combine(GetInstallPath(), InsLauncherExe);
		}

		public string GetInstallPath()
		{
			string componentCode = GetComponentCodeInsLauncher();
			return GetInstallPath(componentCode);
		}

		public string GetProductName()
		{
			return _productInfo.Name;
		}

		public string GetShortcutProductName()
		{
			return _productInfo.ShortcutProductName;
		}

		public bool IsProductInstalled()
		{
			string componentCode = GetComponentCodeInsLauncher();
			return IsProductInstalled(componentCode);
		}

		public string RemoveShortcuts(IEnumerable<ShortcutCfg> shortcutCfgs, string packageName)
		{
			string insLauncherPath = GetInstallPath();
			if (string.IsNullOrWhiteSpace(insLauncherPath))
			{
				return $"Musi być zainstalowany program {GetProductName()}. Skrót nie został usunięty.";
			}
			if (string.IsNullOrWhiteSpace(packageName))
			{
				return "Nie podano nazwy pakietu potrzebnej do usunięcia folderu zawierającego ikony rozszerzenia.";
			}

			var sb = new StringBuilder();

			foreach (var shortcutCfg in shortcutCfgs)
			{
				string scenario = shortcutCfg.ScenarioName;
				string target = Path.Combine(insLauncherPath, InsLauncherExe);
				string arg = string.Join("/", GetShortcutProductName(), scenario);
				if (shortcutCfg.CreateOnDesktop && Shortcut.RemoveDesktopShortcut(target, arg))
				{
					_ = sb.AppendLine($"Skrót dla \"{scenario}\" został usunięty z pulpitu.");
				}

				if (shortcutCfg.CreateInStartMenu && Shortcut.RemoveStartMenuShortcut(GetShortcutProductName(), target, arg))
				{
					_ = sb.AppendLine($"Skrót dla \"{scenario}\" został usunięty z menu Start.");
				}
			}
			return sb.ToString();
		}
			
		private static string GetInstallPath(string componentCode)
		{
			var sbProductBuf = new StringBuilder(39);
			Msi.MsiGetProductCode(componentCode, sbProductBuf);
			long size = 2048;
			var sbPathBuf = new StringBuilder((int)size);
			if (Msi.MSI_INSTALLSTATE.INSTALLSTATE_LOCAL == (Msi.MSI_INSTALLSTATE)Msi.MsiGetComponentPath(sbProductBuf.ToString(), componentCode, sbPathBuf, ref size))
			{
				return Path.GetDirectoryName(sbPathBuf.ToString());
			}

			return null;
		}

		private static bool IsProductInstalled(string componentCode)
		{
			string productInstallPath = GetInstallPath(componentCode);
			if (!string.IsNullOrEmpty(productInstallPath) && Directory.Exists(productInstallPath))
			{
				return true;
			}
			return false;
		}

		private string GetComponentCodeInsLauncher()
		{
			return _productInfo.ComponentCode;
		}

		private string GetIcoFolderPath(string packageName)
		{
			if (string.IsNullOrEmpty(_icoFolderPath))
			{
				_icoFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "InsERT", "Rozszerzenia", packageName, "ico");
				if (!Directory.Exists(_icoFolderPath))
				{
					Directory.CreateDirectory(_icoFolderPath);
				}
			}
			return _icoFolderPath;
		}
	}
}