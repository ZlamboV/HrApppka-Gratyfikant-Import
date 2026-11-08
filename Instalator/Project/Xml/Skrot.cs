using System;
using System.Xml.Serialization;

namespace InsERT.Mox.QuickInstallerCore.Xml
{
	[Serializable]
	[XmlRoot(Namespace = "", IsNullable = false)]
	public class Skrot
	{
		/// <summary>
		/// nazwa scenariusza uruchomienia z manifestu pakietu (LaunchScenario)
		/// </summary>
		[XmlAttribute]
		public string NazwaScenariusza { get; set; }
		/// <summary>
		/// opis skrótu
		/// </summary>
		[XmlAttribute]
		public string Opis { get; set; }
		/// <summary>
		/// plik z ikoną
		/// </summary>
		[XmlElement]
		public string PlikIkony { get; set; }
	}
}