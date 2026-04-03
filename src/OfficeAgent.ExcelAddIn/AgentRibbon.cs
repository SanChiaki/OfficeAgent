using System;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace OfficeAgent.ExcelAddIn
{
    public partial class AgentRibbon
    {
        private void AgentRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            try
            {
                toggleTaskPaneButton.Image = Properties.Resources.Logo;
            }
            catch
            {
                // Logo is optional; skip if not found
            }
        }

        private void ToggleTaskPaneButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.TaskPaneController?.Toggle();
        }

        private async void LoginButton_Click(object sender, RibbonControlEventArgs e)
        {
            var settings = Globals.ThisAddIn.SettingsStore.Load();
            var ssoUrl = settings.SsoUrl;

            if (string.IsNullOrWhiteSpace(ssoUrl))
            {
                MessageBox.Show("\u8BF7\u5148\u5728\u8BBE\u7F6E\u4E2D\u914D\u7F6E SSO \u5730\u5740\u3002", "Resy AI", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            loginButton.Label = "\u767B\u5F55\u4E2D...";
            loginButton.Enabled = false;

            try
            {
                var popup = new SsoLoginPopup(ssoUrl, Globals.ThisAddIn.SharedCookies, Globals.ThisAddIn.CookieStore);
                await popup.InitializeAsync();
                popup.ShowDialog();
            }
            finally
            {
                loginButton.Label = "\u767B\u5F55";
                loginButton.Enabled = true;
            }
        }
    }
}
