using System;
using System.Collections.Generic;

namespace InsERT.Mox.QuickInstallerCore
{
	internal static class KnownProducts
	{
		/* ComponentCode to guid przypisany do InsLauncher.exe w Nexo.ism albo NexoNav.ism
		 * */
		internal const string ComponentCodeInsLauncherKiP2 = "{EE7B408F-581F-4AF7-B1ED-BCBE41F6ED70}";
		internal const string ComponentCodeInsLauncherNexo = "{F448FB2E-8EB5-4A69-B42E-3F779DA8B212}";
		internal const string ComponentCodeInsLauncherNexoAlfa = "{AB4037F2-EF36-4B34-B305-599FEA21C9E8}";

		private static Dictionary<Product, ProductInfo> _products = new()
		{
			[Product.Nexo] = new ProductInfo
			{
				Name = "InsERT nexo",
				ComponentCode = ComponentCodeInsLauncherNexo,
				ShortcutProductName = "nexo"
			},
			[Product.Navireo] = new ProductInfo
			{
				Name = "InsERT Kadry i Płace 2",
				ComponentCode = ComponentCodeInsLauncherKiP2,
				ShortcutProductName = "Navireo"
			},
			[Product.NexoAlfa] = new ProductInfo
			{
				Name = "InsERT nexo ALFA",
				ComponentCode = ComponentCodeInsLauncherNexoAlfa,
				ShortcutProductName = "NexoAlfa"
			}
		};

		internal static ProductInfo GetProductInfo(Product product)
		{
			if (_products.TryGetValue(product, out var value))
			{
				return value;
			}
			throw new ArgumentException($"Nieznany produkt: {product}");
		}
	}
}