using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

namespace InsERT.Mox.QuickInstallerCore
{
	internal class Shortcut
	{
		public static bool CreateDesktopShortcut(string name, string description, string target, string iconPath, int icon, params string[] args)
		{
			string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);

			if (ShortcutExists(desktopPath, name, target, args))
			{
				return true;
			}

            return CreateShortcut(desktopPath, name, description, target, iconPath, icon, args);
		}

		public static bool CreateStartMenuShortcut(string product, string name, string description, string target, string iconPath, int icon, params string[] args)
		{
			string appStartMenuPath = GetStartMenuFolder(product);

			if (ShortcutExists(appStartMenuPath, name, target, args))
			{
				return true;
			}

            if (!Directory.Exists(appStartMenuPath))
			{
				Directory.CreateDirectory(appStartMenuPath);
			}

			return CreateShortcut(appStartMenuPath, name, description, target, iconPath, icon, args);
		}

		private static bool ShortcutExists(string path, string name, string target, string[] args)
		{
			string filename = Path.Combine(path, name + extLink);

			if (!File.Exists(filename))
			{
				return false;
			}

			string expectedArgs = args != null && args.Length > 0 ? string.Join(" ", args) : string.Empty;
			var sbPath = new StringBuilder(1024);
			var sbArgs = new StringBuilder(1024);

			try
			{
				var link = (IShellLink)new ShellLink();
				var file = (IPersistFile)link;

				file.Load(filename, 0);
				var data = new WIN32_FIND_DATAW();

				link.GetPath(sbPath, sbPath.Capacity, out data, 0);
				link.GetArguments(sbArgs, sbArgs.Capacity);

				return
					string.Equals(target, sbPath.ToString(), StringComparison.OrdinalIgnoreCase) &&
					string.Equals(expectedArgs, sbArgs.ToString(), StringComparison.OrdinalIgnoreCase);
			}
			catch
			{
				return false;
			}
		}

        private static string GetStartMenuFolder(string product)
		{
			string startMenuPath = Environment.GetFolderPath(Environment.SpecialFolder.StartMenu);
			string menuName = product == "nexo" ? "InsERT nexo" : "Navireo";
			return Path.Combine(startMenuPath, "Programs", "InsERT", menuName);
		}

		private static bool CreateShortcut(string locationPath, string name, string description, string target, string iconPath, int icon, params string[] args)
		{
			try
			{
				var link = (IShellLink)new ShellLink();

				link.SetPath(target);
				link.SetWorkingDirectory(Path.GetDirectoryName(target));

				if (args.Length > 0)
				{
					link.SetArguments(string.Join(" ", args));
				}

				if (!string.IsNullOrWhiteSpace(description))
				{
					link.SetDescription(description);
				}

				if (icon >= 0)
				{
					link.SetIconLocation(iconPath ?? target, icon);
				}

				var file = (IPersistFile)link;
				string filename = GetLinkUniqueFilename(name, locationPath);
				file.Save(filename, false);

				return true;
			}
			catch
			{
				return false;
			}
		}

		private const string extLink = ".lnk";
		private static string GetLinkUniqueFilename(string baseName, string path)
		{
			string filename = null;
			string fullname = null;
			for (uint i = 0; ; i++)
			{
				string diff = i == 0 ? string.Empty : string.Format(" ({0})", i);
				filename = string.Join(string.Empty, baseName, diff, extLink);
				fullname = Path.Combine(path, filename);
				if (!File.Exists(fullname))
				{
					try
					{
						using (var f = File.Open(fullname, FileMode.CreateNew, FileAccess.Write, FileShare.None))
						{
							break;
						}
					}
					catch
					{
					}
				}
			}
			return fullname;
		}

		internal static bool RemoveStartMenuShortcut(string product, string target, string arg)
		{
			string appStartMenuPath = GetStartMenuFolder(product);

			if (!Directory.Exists(appStartMenuPath))
			{
				return false;
			}

			return RemoveShortcut(appStartMenuPath, target, arg);
		}

		internal static bool RemoveDesktopShortcut(string target, string arg)
		{
			string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
			return RemoveShortcut(desktopPath, target, arg);
		}

		private static bool RemoveShortcut(string path, string target, string arg)
		{
			string[] filenames = Directory.GetFiles(path, "*" + extLink);
			foreach (string filename in filenames)
			{
				var sbPath = new StringBuilder(1024);
				var sbArgs = new StringBuilder(1024);
				try
				{
					var link = (IShellLink)new ShellLink();
					var file = (IPersistFile)link;

					file.Load(filename, 0);
					file = null;
					var data = new WIN32_FIND_DATAW();

					link.GetPath(sbPath, sbPath.Capacity, out data, 0);
					link.GetArguments(sbArgs, sbArgs.Capacity);
					if (string.Compare(target, sbPath.ToString(), true) == 0 && sbArgs.ToString().ToLower().Contains(arg.ToLower()))
					{
						File.Delete(filename);
					}
				}
				catch
				{ }
			}
			return true;
		}

		[ComImport]
		[Guid("00021401-0000-0000-C000-000000000046")]
		private class ShellLink
		{
		}

		[ComImport]
		[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
		[Guid("000214F9-0000-0000-C000-000000000046")]
		private interface IShellLink
		{
			void GetPath([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszFile, int cchMaxPath, out WIN32_FIND_DATAW pfd, int fFlags);
			void GetIDList(out IntPtr ppidl);
			void SetIDList(IntPtr pidl);
			void GetDescription([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszName, int cchMaxName);
			void SetDescription([MarshalAs(UnmanagedType.LPWStr)] string pszName);
			void GetWorkingDirectory([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszDir, int cchMaxPath);
			void SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string pszDir);
			void GetArguments([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszArgs, int cchMaxPath);
			void SetArguments([MarshalAs(UnmanagedType.LPWStr)] string pszArgs);
			void GetHotkey(out short pwHotkey);
			void SetHotkey(short wHotkey);
			void GetShowCmd(out int piShowCmd);
			void SetShowCmd(int iShowCmd);
			void GetIconLocation([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszIconPath, int cchIconPath, out int piIcon);
			void SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string pszIconPath, int iIcon);
			void SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string pszPathRel, int dwReserved);
			void Resolve(IntPtr hwnd, int fFlags);
			void SetPath([MarshalAs(UnmanagedType.LPWStr)] string pszFile);
		}

		[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
		private struct WIN32_FIND_DATAW
		{
			public uint dwFileAttributes;
			public long ftCreationTime;
			public long ftLastAccessTime;
			public long ftLastWriteTime;
			public uint nFileSizeHigh;
			public uint nFileSizeLow;
			public uint dwReserved0;
			public uint dwReserved1;
			[MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
			public string cFileName;
			[MarshalAs(UnmanagedType.ByValTStr, SizeConst = 14)]
			public string cAlternateFileName;
		}
	}
}
