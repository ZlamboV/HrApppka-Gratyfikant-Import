using System.IO;
using System.Linq;
using System.Reflection;

namespace InsERT.Mox.QuickInstallerCore
{
	internal static class ResourceReader
	{
		internal static string GetResourceAsText(string name)
		{
			var stream = GetResourceStream(name);
			try
			{
				using (var reader = new StreamReader(stream))
				{
					stream = null;
					return reader.ReadToEnd();
				}
			}
			finally
			{
				stream?.Dispose();
			}
		}

		internal static Stream GetResourceStream(string name)
		{
			var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(GetExpectedResourceName(name))
				?? throw new QuickInstallerException($"Nie znaleziono zasobu o nazwie '{name}'.");
			return stream;
		}

		internal static bool HasResource(string name)
		{
			string expectedName = GetExpectedResourceName(name);
			return Assembly.GetExecutingAssembly().GetManifestResourceNames().Any(x => string.Equals(expectedName, x));
		}

		private static string GetExpectedResourceName(string name)
		{
			return "InsERT.Mox.QuickInstallerCore." + name;
		}
	}
}