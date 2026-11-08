using System;

namespace InsERT.Mox.QuickInstallerCore
{
	public class QuickInstallerException : Exception
	{
		public QuickInstallerException()
		{
		}

		public QuickInstallerException(string message) : base(message)
		{
		}

		public QuickInstallerException(string message, Exception innerException) : base(message, innerException)
		{
		}
	}
}