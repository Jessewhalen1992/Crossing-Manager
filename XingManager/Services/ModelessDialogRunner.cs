using System;
using System.Threading;
using System.Windows.Forms;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace XingManager.Services
{
    /// <summary>
    /// Helper that shows a WinForms dialog modelessly while waiting synchronously
    /// for the user to close it. This allows interaction with the DWG (panning, zooming, etc.)
    /// while the dialog is visible.
    /// </summary>
    internal static class ModelessDialogRunner
    {
        private sealed class WindowHandleWrapper : IWin32Window
        {
            private readonly IntPtr _handle;

            public WindowHandleWrapper(IntPtr handle)
            {
                _handle = handle;
            }

            public IntPtr Handle => _handle;
        }

        public static DialogResult ShowDialog(Form dialog)
        {
            if (dialog == null)
                throw new ArgumentNullException(nameof(dialog));

            var completed = false;
            var result = DialogResult.None;

            void OnClosed(object sender, FormClosedEventArgs args)
            {
                dialog.FormClosed -= OnClosed;
                completed = true;
                result = dialog.DialogResult;
            }

            dialog.FormClosed += OnClosed;

            var mainWindow = AcadApp.MainWindow;
            if (mainWindow != null)
            {
                var owner = new WindowHandleWrapper(mainWindow.Handle);
                dialog.Show(owner);
            }
            else
            {
                dialog.Show();
            }

            while (!completed)
            {
                System.Windows.Forms.Application.DoEvents();
                Thread.Sleep(25);
            }

            return result;
        }
    }
}

/////////////////////////////////////////////////////////////////////

