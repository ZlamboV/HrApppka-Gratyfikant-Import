using System.Collections.Generic;

namespace InsERT.Mox.QuickInstallerCore
{
	public class DeployParameters
	{
		public string SqlServer { get; set; }
		public bool WindowsAuthentication { get; set; }
		public string User { get; set; }
		public string Password { get; set; }
		public List<string> Databases { get; set; }
	}
}