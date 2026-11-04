using System;

namespace InsERT.Mox.QuickInstallerCore
{
	internal static class InstallationConfigExtensions
	{
		internal static Product DetermineProduct(this InstallationConfig config)
		{
			ArgumentNullException.ThrowIfNull(config);

			if (config.KiP2)
			{
				return Product.Navireo;
			}
			else if (config.Alfa)
			{
				return Product.NexoAlfa;
			}
			else
			{
				return Product.Nexo;
			}
		}
	}
}