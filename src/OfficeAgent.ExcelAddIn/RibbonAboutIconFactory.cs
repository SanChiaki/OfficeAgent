using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;

namespace OfficeAgent.ExcelAddIn
{
    internal static class RibbonAboutIconFactory
    {
        public static Image CreateAboutIcon(bool hasUpdate)
        {
            var bitmap = new Bitmap(32, 32);
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                graphics.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
                graphics.Clear(Color.Transparent);

                using (var fillBrush = new SolidBrush(Color.FromArgb(255, 45, 112, 179)))
                using (var borderPen = new Pen(Color.FromArgb(255, 28, 80, 135), 2f))
                {
                    graphics.FillEllipse(fillBrush, 5, 5, 22, 22);
                    graphics.DrawEllipse(borderPen, 5, 5, 22, 22);
                }

                using (var font = new Font("Segoe UI", 18f, FontStyle.Bold, GraphicsUnit.Pixel))
                using (var textBrush = new SolidBrush(Color.White))
                using (var format = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                {
                    graphics.DrawString("i", font, textBrush, new RectangleF(5, 4, 22, 24), format);
                }

                if (hasUpdate)
                {
                    using (var shadowBrush = new SolidBrush(Color.White))
                    using (var dotBrush = new SolidBrush(Color.FromArgb(255, 220, 38, 38)))
                    using (var dotPen = new Pen(Color.FromArgb(255, 153, 27, 27), 1f))
                    {
                        graphics.FillEllipse(shadowBrush, 19, 2, 11, 11);
                        graphics.FillEllipse(dotBrush, 20, 3, 9, 9);
                        graphics.DrawEllipse(dotPen, 20, 3, 9, 9);
                    }
                }
            }

            return bitmap;
        }
    }
}
