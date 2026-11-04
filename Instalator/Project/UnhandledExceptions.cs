using System;
using System.Windows;

namespace InsERT.Mox.QuickInstallerCore
{
	internal static class UnhandledExceptions
	{
		internal static void HandleInCurrentDomain()
		{
			AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
		}

		internal static void HandleInWpfApp(Application app)
		{
			app.Dispatcher.UnhandledException += Dispatcher_UnhandledException;
		}

		private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
		{
			if (e.ExceptionObject is Exception exception)
			{
				Logger.LogException(exception);
			}
		}

		private static void Dispatcher_UnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
		{
			Logger.LogException(e.Exception);
			e.Handled = true;
		}
	}
}