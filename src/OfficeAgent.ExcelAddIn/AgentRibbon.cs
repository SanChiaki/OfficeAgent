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
    }
}
