using System;
using System.Runtime.InteropServices;
using OfficeAgent.Core.Diagnostics;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace OfficeAgent.ExcelAddIn.Excel
{
    internal sealed class ExcelFocusCoordinator
    {
        private readonly Action activateActiveWindow;
        private readonly Func<IntPtr> getApplicationWindowHandle;
        private readonly Func<IntPtr, IntPtr> setActiveWindow;
        private readonly Func<IntPtr, IntPtr> setFocusWindow;
        private readonly Func<IntPtr, bool> setForegroundWindow;

        public ExcelFocusCoordinator(ExcelInterop.Application application)
            : this(
                  () => application?.ActiveWindow?.Activate(),
                  () => application == null ? IntPtr.Zero : new IntPtr(application.Hwnd),
                  NativeMethods.SetActiveWindow,
                  NativeMethods.SetFocus,
                  NativeMethods.SetForegroundWindow)
        {
        }

        internal ExcelFocusCoordinator(
            Action activateActiveWindow,
            Func<IntPtr> getApplicationWindowHandle,
            Func<IntPtr, IntPtr> setActiveWindow,
            Func<IntPtr, IntPtr> setFocusWindow,
            Func<IntPtr, bool> setForegroundWindow)
        {
            this.activateActiveWindow = activateActiveWindow ?? throw new ArgumentNullException(nameof(activateActiveWindow));
            this.getApplicationWindowHandle = getApplicationWindowHandle ?? throw new ArgumentNullException(nameof(getApplicationWindowHandle));
            this.setActiveWindow = setActiveWindow ?? throw new ArgumentNullException(nameof(setActiveWindow));
            this.setFocusWindow = setFocusWindow ?? throw new ArgumentNullException(nameof(setFocusWindow));
            this.setForegroundWindow = setForegroundWindow ?? throw new ArgumentNullException(nameof(setForegroundWindow));
        }

        public void RestoreWorksheetFocus(Action activateSelection)
        {
            TryInvoke("window.activate", activateActiveWindow);
            TryInvoke("selection.activate", activateSelection);

            IntPtr windowHandle;
            try
            {
                windowHandle = getApplicationWindowHandle();
            }
            catch (Exception error)
            {
                OfficeAgentLog.Warn("focus", "window.handle.failed", $"Failed to resolve Excel window handle: {error.Message}");
                return;
            }

            if (windowHandle == IntPtr.Zero)
            {
                return;
            }

            TryInvoke(
                "window.active",
                () => setActiveWindow(windowHandle));
            TryInvoke(
                "window.focus",
                () => setFocusWindow(windowHandle));
            TryInvoke(
                "window.foreground",
                () => setForegroundWindow(windowHandle));
        }

        private static void TryInvoke(string operation, Action action)
        {
            if (action == null)
            {
                return;
            }

            try
            {
                action();
            }
            catch (Exception error)
            {
                OfficeAgentLog.Warn("focus", operation, $"Focus restoration step failed: {error.Message}");
            }
        }

        private static class NativeMethods
        {
            [DllImport("user32.dll")]
            public static extern IntPtr SetActiveWindow(IntPtr hWnd);

            [DllImport("user32.dll")]
            public static extern IntPtr SetFocus(IntPtr hWnd);

            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static extern bool SetForegroundWindow(IntPtr hWnd);
        }
    }
}
