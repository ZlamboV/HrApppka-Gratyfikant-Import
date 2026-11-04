using System;
using System.Diagnostics;

namespace InsERT.Mox.QuickInstallerCore
{
	internal static class Logger
	{
		internal static void ConfigureLog(string logLocation)
		{
			_ = Trace.Listeners.Add(new TextWriterTraceListener(logLocation));
			Trace.AutoFlush = true;
		}

		internal static void LogException(Exception exception)
		{
			Trace.TraceError(FormatMessage(exception.ToString()));
		}

		internal static void LogInformation(string message)
		{
			Trace.TraceInformation(FormatMessage(message));
		}

		private static string FormatMessage(string message)
		{
			return $"{DateTime.Now:yyyy-MM-dd HH:mm:ss}\t{message}";
		}
	}
}