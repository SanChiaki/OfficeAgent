using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

internal static class RenderBatchUploadProgressDialog
{
    [STAThread]
    private static void Main(string[] args)
    {
        var outputDirectory = args.Length > 0 ? args[0] : Environment.CurrentDirectory;
        Directory.CreateDirectory(outputDirectory);

        foreach (var percent in new[] { 50, 75, 100, 125, 150, 175, 200, 225, 250, 300 })
        {
            Render(outputDirectory, "batch-upload-dialog-" + percent + ".png", percent / 100f);
        }
    }

    private static void Render(string outputDirectory, string fileName, float fontScale)
    {
        using (var dialog = OfficeAgent.ExcelAddIn.Dialogs.BatchUploadProgressDialog.CreateSample())
        using (var scaledFont = new Font(dialog.Font.FontFamily, dialog.Font.Size * fontScale, dialog.Font.Style))
        {
            ApplyFont(dialog, scaledFont);
            dialog.StartPosition = FormStartPosition.Manual;
            dialog.Location = new Point(-32000, -32000);
            dialog.Show();
            Application.DoEvents();
            dialog.PerformLayout();
            Application.DoEvents();

            using (var bitmap = new Bitmap(dialog.ClientSize.Width, dialog.ClientSize.Height))
            {
                dialog.DrawToBitmap(bitmap, new Rectangle(Point.Empty, dialog.ClientSize));
                bitmap.Save(Path.Combine(outputDirectory, fileName), System.Drawing.Imaging.ImageFormat.Png);
            }

            dialog.Close();
        }
    }

    private static void ApplyFont(Control root, Font font)
    {
        root.Font = font;
        foreach (Control child in root.Controls)
        {
            ApplyFont(child, font);
        }
    }
}
