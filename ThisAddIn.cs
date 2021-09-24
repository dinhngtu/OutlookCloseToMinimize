using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;

namespace OutlookCloseToMinimize {
    [ComImport]
    [Guid("00000114-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IOleWindow {
        int GetWindow(out IntPtr phwnd);
        int ContextSensitiveHelp([In, MarshalAs(UnmanagedType.Bool)] bool fEnterMode);
    }

    public partial class ThisAddIn {
        [DllImport("comctl32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetWindowSubclass(IntPtr hWnd, IntPtr pfnSubclass, UIntPtr uIdSubclass, UIntPtr dwRefData);
        [DllImport("comctl32.dll")]
        static extern IntPtr DefSubclassProc(IntPtr hWnd, uint uMsg, UIntPtr wParam, IntPtr lParam);
        [DllImport("comctl32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool RemoveWindowSubclass(IntPtr hWnd, IntPtr pfnSubclass, UIntPtr uIdSubclass);
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        delegate IntPtr SubclassProc(IntPtr hWnd, uint uMsg, UIntPtr wParam, IntPtr lParam, UIntPtr uIdSubclass, UIntPtr dwRefData);

        const int S_OK = 0;
        const uint WM_CLOSE = 0x0010;
        const int SW_MINIMIZE = 6;

        IntPtr _hWnd = IntPtr.Zero, _scProc = IntPtr.Zero;
        GCHandle _gch;

        IntPtr MySubclassProc(IntPtr hWnd, uint uMsg, UIntPtr wParam, IntPtr lParam, UIntPtr uIdSubclass, UIntPtr dwRefData) {
            if (uMsg == WM_CLOSE) {
                ShowWindow(hWnd, SW_MINIMIZE);
                return new IntPtr(1);
            }
            return DefSubclassProc(hWnd, uMsg, wParam, lParam);
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e) {
            // Outlook doesn't send Shutdown events on close any more,
            // so we need to hook the Quit event to remove our subclass callback
            // before our module is unloaded to avoid crashes
            ((Outlook.ApplicationEvents_11_Event)Application).Quit += ThisAddIn_Quit;
            if (((IOleWindow)Application.ActiveExplorer()).GetWindow(out _hWnd) == S_OK && _hWnd != IntPtr.Zero) {
                _gch = GCHandle.Alloc(this);
                _scProc = Marshal.GetFunctionPointerForDelegate((SubclassProc)MySubclassProc);
                if (SetWindowSubclass(_hWnd, _scProc, UIntPtr.Zero, UIntPtr.Zero)) {
                    return;
                }
                _gch.Free();
            }
            _hWnd = IntPtr.Zero;
        }

        private void ThisAddIn_Quit() {
            if (_hWnd != IntPtr.Zero) {
                RemoveWindowSubclass(_hWnd, _scProc, UIntPtr.Zero);
                _hWnd = IntPtr.Zero;
                _gch.Free();
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e) {
            ThisAddIn_Quit();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup() {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
