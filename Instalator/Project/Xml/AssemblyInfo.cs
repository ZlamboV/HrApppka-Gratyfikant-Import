using System;
using System.Xml.Serialization;

namespace InsERT.Mox.QuickInstallerCore.Xml
{
	[Serializable]
	[XmlRoot(Namespace = "", IsNullable = false)]
	public class AssemblyInfo
	{
		[XmlElement]
		public string Company { get; set; }
		[XmlElement]
		public string Copyright { get; set; }
		[XmlElement]
		public string FileVersion { get; set; }
		[XmlElement]
		public string Product { get; set; }
		[XmlElement]
		public string Version { get; set; }
	}
}