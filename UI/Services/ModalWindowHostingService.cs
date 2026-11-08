using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Interop;

namespace HrAppka_Import_Pracowników.UI.Services
{
    /// <summary>
    /// Serwis zarządzający modalnym wyświetlaniem okien WPF wewnątrz Gratyfikant Nexo Pro.
    /// Zapewnia poprawne dowiązanie Owner HWND (Application.Current / Process MainWindowHandle)
    /// oraz bezpieczne wywołanie na wątku STA/Dispatcherze.
    /// </summary>
    public static class ModalWindowHostingService
    {
        private static class NativeMethods
        {
            [DllImport("user32.dll")]
            public static extern IntPtr GetActiveWindow();

            [DllImport("user32.dll")]
            public static extern IntPtr GetForegroundWindow();
        }

        /// <summary>
        /// Wyświetla okno modalnie z gwarancją poprawnego wątku STA oraz dowiązania do okna głównego Nexo.
        /// </summary>
        /// <typeparam name="TWindow">Typ okna dziedziczącego po Window</typeparam>
        /// <param name="windowFactory">Fabryka tworząca instancję okna</param>
        /// <returns>Wynik ShowDialog()</returns>
        public static bool? ShowModal<TWindow>(Func<TWindow> windowFactory) where TWindow : Window
        {
            if (windowFactory == null) throw new ArgumentNullException(nameof(windowFactory));

            // Scenariusz 1: Dostępny globalny kontekst Application.Current (Standard w Nexo)
            if (Application.Current != null && Application.Current.Dispatcher != null)
            {
                if (Application.Current.Dispatcher.CheckAccess())
                {
                    return ExecuteShowDialog(windowFactory);
                }
                else
                {
                    return Application.Current.Dispatcher.Invoke(() => ExecuteShowDialog(windowFactory));
                }
            }

            // Scenariusz 2: Brak Application.Current (np. zewnętrzny test harness / unmanaged shell)
            if (Thread.CurrentThread.GetApartmentState() == ApartmentState.STA)
            {
                return ExecuteShowDialog(windowFactory);
            }
            else
            {
                bool? dialogResult = null;
                var staThread = new Thread(() =>
                {
                    dialogResult = ExecuteShowDialog(windowFactory);
                });
                staThread.SetApartmentState(ApartmentState.STA);
                staThread.Start();
                staThread.Join();
                return dialogResult;
            }
        }

        private static bool? ExecuteShowDialog<TWindow>(Func<TWindow> windowFactory) where TWindow : Window
        {
            var window = windowFactory();
            AttachNexoOwner(window);
            return window.ShowDialog();
        }

        /// <summary>
        /// Przypisuje okno nadrzędne (Owner) do okna WPF, priorytetyzując Application.Current.MainWindow,
        /// z fallbackiem do Win32 HWND procesu Gratyfikanta Nexo.
        /// </summary>
        public static void AttachNexoOwner(Window window)
        {
            if (window == null) return;

            // Krok 1: Próba podpięcia przez WPF Application.Current.MainWindow
            try
            {
                var app = Application.Current;
                if (app?.MainWindow != null &&
                    app.MainWindow != window &&
                    app.MainWindow.Dispatcher.CheckAccess() &&
                    app.MainWindow.IsVisible)
                {
                    window.Owner = app.MainWindow;
                    return;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[UI Hosting] Nie udało się przypisać Owner przez Application.Current: {ex.Message}");
            }

            // Krok 2: Fallback natywny Win32 przez WindowInteropHelper i Process.MainWindowHandle
            try
            {
                IntPtr hwnd = IntPtr.Zero;
                using (var currentProcess = Process.GetCurrentProcess())
                {
                    hwnd = currentProcess.MainWindowHandle;
                }

                if (hwnd == IntPtr.Zero)
                {
                    hwnd = NativeMethods.GetActiveWindow();
                }

                if (hwnd == IntPtr.Zero)
                {
                    hwnd = NativeMethods.GetForegroundWindow();
                }

                if (hwnd != IntPtr.Zero)
                {
                    var helper = new WindowInteropHelper(window);
                    helper.Owner = hwnd;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[UI Hosting] Błąd przypisywania Win32 Owner HWND: {ex.Message}");
            }
        }
    }
}
