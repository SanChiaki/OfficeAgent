using System;
using System.Drawing;
using System.IO;

namespace OfficeAgent.ExcelAddIn
{
    internal static class RibbonAboutIconFactory
    {
        private const string ResourcesDirectoryName = "Resources";
        private const string AboutIconFileName = "about_32.png";
        private const string AboutIconWithUpdateFileName = "about_update_32.png";

        public static Image LoadAboutIcon(bool hasUpdate)
        {
            var iconPath = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                ResourcesDirectoryName,
                hasUpdate ? AboutIconWithUpdateFileName : AboutIconFileName);

            if (File.Exists(iconPath))
            {
                using (var stream = File.OpenRead(iconPath))
                using (var image = Image.FromStream(stream))
                {
                    return new Bitmap(image);
                }
            }

            return (Image)Properties.Resources.Logo.Clone();
        }
    }
}
