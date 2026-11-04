using InsERT.Mox.QuickInstallerCore.Xml;
using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace InsERT.Mox.QuickInstallerCore
{
	public class ShortcutCfg : INotifyPropertyChanged
	{
		private bool _createInStartMenu;
		private bool _createOnDesktop;
		private ImageSource _icon;
		private ImageSource _nexoIcon;
		private string _scenarioName;

		public ShortcutCfg(string scenarioName, Skrot shortcut, bool createOnDesktop, bool createInStartMenu)
		{
			ScenarioName = scenarioName;
			Shortcut = shortcut;
			CreateOnDesktop = createOnDesktop;
			CreateInStartMenu = createInStartMenu;
		}

		public event PropertyChangedEventHandler PropertyChanged;

		public bool CreateInStartMenu
		{
			get => _createInStartMenu;
			set
			{
				_createInStartMenu = value;
				OnPropertyChanged(nameof(CreateInStartMenu));
			}
		}

		public bool CreateOnDesktop
		{
			get => _createOnDesktop;
			set
			{
				_createOnDesktop = value;
				OnPropertyChanged(nameof(CreateOnDesktop));
			}
		}

		public ImageSource Icon
		{
			get
			{
				if (_icon == null)
				{
					if (!string.IsNullOrWhiteSpace(Shortcut?.PlikIkony))
					{
						_icon = InstallationConfig.LoadBitmapImage(Shortcut.PlikIkony);
					}

                    _icon ??= NexoIcon;
				}
				return _icon;
			}
		}

		public string ScenarioName
		{
			get => _scenarioName;
			set
			{
				_scenarioName = value;
				OnPropertyChanged(nameof(ScenarioName));
			}
		}

		public Skrot Shortcut { get; }

		private ImageSource NexoIcon
		{
			get
			{
				if (_nexoIcon == null)
				{
					var icon = Resources.NexoIcon;
					var bitmap = icon.ToBitmap();
					IntPtr hBitmap = bitmap.GetHbitmap();

					_nexoIcon = Imaging.CreateBitmapSourceFromHBitmap(
							  hBitmap, IntPtr.Zero, Int32Rect.Empty,
							  BitmapSizeOptions.FromEmptyOptions());
				}
				return _nexoIcon;
			}
		}

		private void OnPropertyChanged(string name)
		{
			PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
		}
	}
}