using System;
using System.Runtime.InteropServices;
using System.Text;

namespace InsERT.Mox.QuickInstallerCore
{
	internal class Msi
    {
        [DllImport("msi.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern long MsiGetProductCode(string strComponent, StringBuilder sbProductBuf);

        [DllImport("msi.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern uint MsiGetComponentPath(string szProduct, string szComponent, StringBuilder sbPath, ref long sizeBuf);

		public enum MSI_INSTALLSTATE : long
		{
			INSTALLSTATE_NOTUSED = -7,  // component disabled
			INSTALLSTATE_BADCONFIG = -6,  // configuration data corrupt
			INSTALLSTATE_INCOMPLETE = -5,  // installation suspended or in progress
			INSTALLSTATE_SOURCEABSENT = -4,  // run from source, source is unavailable
			INSTALLSTATE_MOREDATA = -3,  // return buffer overflow
			INSTALLSTATE_INVALIDARG = -2,  // invalid function argument
			INSTALLSTATE_UNKNOWN = -1,  // unrecognized product or feature
			INSTALLSTATE_BROKEN = 0,  // broken
			INSTALLSTATE_ADVERTISED = 1,  // advertised feature
			INSTALLSTATE_REMOVED = 1,  // component being removed (action state, not settable)
			INSTALLSTATE_ABSENT = 2,  // uninstalled (or action state absent but clients remain)
			INSTALLSTATE_LOCAL = 3,  // installed on local drive
			INSTALLSTATE_SOURCE = 4,  // run from source, CD or net
			INSTALLSTATE_DEFAULT = 5,  // use default, local or source
		}
	}
}
