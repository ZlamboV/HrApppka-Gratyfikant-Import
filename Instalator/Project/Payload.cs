using System.IO;

namespace InsERT.Mox.QuickInstallerCore
{
	internal static class Payload
	{
		internal static string GetItemAsText(string payloadItemName)
		{
			return ResourceReader.GetResourceAsText(FormatPayloadItemName(payloadItemName));
		}

		internal static Stream GetItemStream(string payloadItemName)
		{
			return ResourceReader.GetResourceStream(FormatPayloadItemName(payloadItemName));
		}

		internal static bool HasItem(string payloadItemName)
		{
			return ResourceReader.HasResource(FormatPayloadItemName(payloadItemName));
		}

		private static string FormatPayloadItemName(string payloadItemName)
		{
			return $"Payload.{payloadItemName}";
		}
	}
}