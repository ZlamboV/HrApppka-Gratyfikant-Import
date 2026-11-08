using System;
using System.IO;

namespace InsERT.Mox.QuickInstallerCore
{
	internal static class TemporaryPayload
	{
		private static readonly string _tempPath = Path.Combine(System.IO.Path.GetTempPath(), $"QuickInstaller_{Guid.NewGuid().ToString()}");

		static TemporaryPayload()
		{
			Directory.CreateDirectory(_tempPath);
		}

		internal static void Delete()
		{
			if (!string.IsNullOrWhiteSpace(_tempPath) && Directory.Exists(_tempPath))
			{
				Directory.Delete(_tempPath, true);
			}
		}

		internal static string GetItem(string payloadItemName)
		{
			string unpackedItemPath = Path.Combine(_tempPath, payloadItemName);

			if (File.Exists(unpackedItemPath))
			{
				return unpackedItemPath;
			}

			var inputStream = Payload.GetItemStream(payloadItemName);
			if (inputStream != null)
			{
				using (var outputStream = File.Create(unpackedItemPath))
				{
					inputStream.CopyTo(outputStream);
				}
				inputStream.Dispose();
				return unpackedItemPath;
			}

			return string.Empty;
		}
	}
}