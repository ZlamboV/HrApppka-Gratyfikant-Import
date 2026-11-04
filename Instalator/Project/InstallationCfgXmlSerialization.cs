using InsERT.Mox.QuickInstallerCore.Xml;
using System.IO;
using System.Xml.Serialization;

namespace InsERT.Mox.QuickInstallerCore
{
	internal static class InstallationCfgXmlSerialization
	{
		private static class XmlElementNames
		{
			internal const string Tytul = "Tytul";
			internal const string TekstPowitalny = "TekstPowitalny";
			internal const string DostawcaInfo = "DostawcaInfo";
			internal const string KiP2 = "KiP2";
			internal const string Alfa = "Alfa";
			internal const string WymaganyRestart = "WymaganyRestart";
			internal const string PlikLogo = "PlikLogo";
			internal const string PlikLicencji = "PlikLicencji";
			internal const string Pakiet = "Pakiet";
			internal const string AssemblyInfo = "AssemblyInfo";
			internal const string Skroty = "Skroty";
			internal const string UtworzSkrotNaPulpicie = "UtworzSkrotNaPulpicie";
			internal const string UtworzSkrotWMenuStart = "UtworzSkrotWMenuStart";
		}

		internal static void LoadXml(Stream xmlStream, InstallationConfig cfg)
		{
			using (var xmlReader = new System.Xml.XmlTextReader(xmlStream))
			{
				while (xmlReader.Read())
				{
					if (xmlReader.IsStartElement())
					{
						if (xmlReader.Name == XmlElementNames.TekstPowitalny)
						{
							cfg.WelcomeText = xmlReader.ReadString();
						}
						else if (xmlReader.Name == XmlElementNames.UtworzSkrotNaPulpicie)
						{
							cfg.CreateDesktopShortcut = xmlReader.ReadElementContentAsBoolean();
						}
						else if (xmlReader.Name == XmlElementNames.UtworzSkrotWMenuStart)
						{
							cfg.CreateMenuStartShortcut = xmlReader.ReadElementContentAsBoolean();
						}
						else if (xmlReader.Name == XmlElementNames.DostawcaInfo)
						{
							cfg.ProviderInfo = xmlReader.ReadString();
						}
						else if (xmlReader.Name == XmlElementNames.Pakiet)
						{
							cfg.PackagePath = xmlReader.ReadString();
						}
						else if (xmlReader.Name == XmlElementNames.Tytul)
						{
							cfg.Title = xmlReader.ReadString();
						}
						else if (xmlReader.Name == XmlElementNames.KiP2)
						{
							cfg.KiP2 = xmlReader.ReadElementContentAsBoolean();
						}
						else if (xmlReader.Name == XmlElementNames.Alfa)
						{
							cfg.Alfa = xmlReader.ReadElementContentAsBoolean();
						}
						else if (xmlReader.Name == XmlElementNames.WymaganyRestart)
						{
							cfg.RequireRestart = xmlReader.ReadElementContentAsBoolean();
						}
						else if (xmlReader.Name == XmlElementNames.PlikLogo)
						{
							cfg.LogoPath = xmlReader.ReadString();
						}
						else if (xmlReader.Name == XmlElementNames.PlikLicencji)
						{
							cfg.LicencePath = xmlReader.ReadString();
						}
						else if (xmlReader.Name == XmlElementNames.Skroty)
						{
							Skrot[] skroty = null;
							using (var subTreeReader = xmlReader.ReadSubtree())
							{
								var serializer = new XmlSerializer(typeof(Skrot[]), new XmlRootAttribute(XmlElementNames.Skroty));
								skroty = serializer.Deserialize(subTreeReader) as Skrot[];
							}
							cfg.Shortcuts = skroty;
						}
						else if (xmlReader.Name == XmlElementNames.AssemblyInfo)
						{
							using (var subTreeReader = xmlReader.ReadSubtree())
							{
								var serializer = new XmlSerializer(typeof(AssemblyInfo), new XmlRootAttribute(XmlElementNames.AssemblyInfo));
								var assemblyInfo = serializer.Deserialize(subTreeReader) as AssemblyInfo;
								if (assemblyInfo != null)
								{
									cfg.AssemblyInfo = assemblyInfo;
								}
							}
						}
					}
				}
			}
		}
	}
}