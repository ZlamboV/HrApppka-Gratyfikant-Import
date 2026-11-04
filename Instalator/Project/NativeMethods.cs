using System.Runtime.InteropServices;

namespace InsERT.Mox.QuickInstallerCore
{
	internal static class NativeMethods
	{
		public const int ATTACH_PARENT_PROCESS = -1;

		[DllImport("kernel32.dll")]
		public static extern bool AllocConsole();

		[DllImport("kernel32.dll")]
		public static extern bool AttachConsole(int dwProcessId);

		[DllImport("Kernel32.dll")]
		public static extern bool FreeConsole();
	}
}