using System;
using System.IO;
using OfficeAgent.Core.Services;
using OfficeAgent.ExcelAddIn.Excel;
using OfficeAgent.ExcelAddIn.TaskPane;
using OfficeAgent.Infrastructure.Security;
using OfficeAgent.Infrastructure.Storage;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace OfficeAgent.ExcelAddIn
{
    public partial class ThisAddIn
    {
        internal TaskPaneController TaskPaneController { get; private set; }
        internal FileSessionStore SessionStore { get; private set; }
        internal FileSettingsStore SettingsStore { get; private set; }
        internal IExcelContextService ExcelContextService { get; private set; }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            var appDataDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OfficeAgent");
            SessionStore = new FileSessionStore(Path.Combine(appDataDirectory, "sessions"));
            SettingsStore = new FileSettingsStore(
                Path.Combine(appDataDirectory, "settings.json"),
                new DpapiSecretProtector());
            ExcelContextService = new ExcelSelectionContextService(Application);
            TaskPaneController = new TaskPaneController(this, SessionStore, SettingsStore, ExcelContextService);
            Application.SheetSelectionChange += Application_SheetSelectionChange;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Application.SheetSelectionChange -= Application_SheetSelectionChange;
        }

        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        private void Application_SheetSelectionChange(object sh, ExcelInterop.Range target)
        {
            TaskPaneController?.PublishSelectionContext(ExcelContextService.GetCurrentSelectionContext());
        }
    }
}
