using Microsoft.Office.Core;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;
using OfficeAgent.Infrastructure.Storage;

namespace OfficeAgent.ExcelAddIn.TaskPane
{
    internal sealed class TaskPaneController
    {
        private readonly ThisAddIn addIn;
        private readonly IExcelContextService excelContextService;
        private readonly FileSessionStore sessionStore;
        private readonly FileSettingsStore settingsStore;
        private Microsoft.Office.Tools.CustomTaskPane taskPane;
        private TaskPaneHostControl hostControl;

        public TaskPaneController(
            ThisAddIn addIn,
            FileSessionStore sessionStore,
            FileSettingsStore settingsStore,
            IExcelContextService excelContextService)
        {
            this.addIn = addIn;
            this.sessionStore = sessionStore;
            this.settingsStore = settingsStore;
            this.excelContextService = excelContextService;
        }

        public void Toggle()
        {
            EnsureCreated();
            taskPane.Visible = !taskPane.Visible;
            if (taskPane.Visible)
            {
                PublishCurrentSelectionContext();
            }
        }

        public void Show()
        {
            EnsureCreated();
            taskPane.Visible = true;
            PublishCurrentSelectionContext();
        }

        public void PublishSelectionContext(SelectionContext selectionContext)
        {
            hostControl?.PublishSelectionContext(selectionContext);
        }

        private void EnsureCreated()
        {
            if (taskPane != null)
            {
                return;
            }

            hostControl = new TaskPaneHostControl(sessionStore, settingsStore, excelContextService);
            taskPane = addIn.CustomTaskPanes.Add(hostControl, "OfficeAgent");
            taskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
            taskPane.Width = 420;
            taskPane.Visible = false;
        }

        private void PublishCurrentSelectionContext()
        {
            PublishSelectionContext(excelContextService.GetCurrentSelectionContext());
        }
    }
}
