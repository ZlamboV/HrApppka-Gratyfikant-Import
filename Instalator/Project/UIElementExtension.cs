using System.Windows;

namespace InsERT.Mox.QuickInstallerCore
{
	internal static class UIElementExtension
	{
		public static void HideAndDisable(this UIElement element)
		{
			element.Visibility = Visibility.Hidden;
			element.IsEnabled = false;
		}

		public static void ShowAndEnable(this UIElement element)
		{
			element.Visibility = Visibility.Visible;
			element.IsEnabled = true;
		}
	}
}
