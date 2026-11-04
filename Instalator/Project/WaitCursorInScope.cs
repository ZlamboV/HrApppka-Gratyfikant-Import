using System;
using System.Windows;
using System.Windows.Input;

namespace InsERT.Mox.QuickInstallerCore
{
	internal sealed class WaitCursorInScope : IDisposable
	{
		private readonly bool _force;
		private FrameworkElement _elem;
		private Cursor _originalCursor;

		public WaitCursorInScope(FrameworkElement elem)
		{
			_force = elem.ForceCursor;
			_elem = elem;
			elem.ForceCursor = true;
			_originalCursor = elem.Cursor;
			elem.Cursor = Cursors.Wait;
		}

		public void Dispose()
		{
			_elem.Cursor = _originalCursor;
			_elem.ForceCursor = _force;
			_elem = null;
			_originalCursor = null;
		}
	}
}