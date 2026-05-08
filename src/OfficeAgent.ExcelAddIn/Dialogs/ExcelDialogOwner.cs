using System;
using System.Windows.Forms;

namespace OfficeAgent.ExcelAddIn.Dialogs
{
    internal sealed class ExcelDialogOwner : IWin32Window
    {
        private ExcelDialogOwner(IntPtr handle)
        {
            Handle = handle;
        }

        public IntPtr Handle { get; }

        public static IWin32Window FromCurrentApplication()
        {
            try
            {
                var hwnd = Globals.ThisAddIn?.Application?.Hwnd ?? 0;
                return hwnd > 0 ? new ExcelDialogOwner(new IntPtr(hwnd)) : null;
            }
            catch
            {
                return null;
            }
        }
    }
}
