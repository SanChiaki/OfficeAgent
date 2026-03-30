using Microsoft.Office.Core;
using OfficeAgent.Core.Diagnostics;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;
using OfficeAgent.Infrastructure.Storage;

namespace OfficeAgent.ExcelAddIn.TaskPane
{
    internal sealed class TaskPaneController
    {
        private readonly ThisAddIn addIn;
        private readonly IExcelContextService excelContextService;
        private readonly IExcelCommandExecutor excelCommandExecutor;
        private readonly IAgentOrchestrator agentOrchestrator;
        private readonly FileSessionStore sessionStore;
        private readonly FileSettingsStore settingsStore;
        private Microsoft.Office.Tools.CustomTaskPane taskPane;
        private TaskPaneHostControl hostControl;

        public TaskPaneController(
            ThisAddIn addIn,
            FileSessionStore sessionStore,
            FileSettingsStore settingsStore,
            IExcelContextService excelContextService,
            IExcelCommandExecutor excelCommandExecutor,
            IAgentOrchestrator agentOrchestrator)
        {
            this.addIn = addIn;
            this.sessionStore = sessionStore;
            this.settingsStore = settingsStore;
            this.excelContextService = excelContextService;
            this.excelCommandExecutor = excelCommandExecutor;
            this.agentOrchestrator = agentOrchestrator;
        }

        public void Toggle()
        {
            EnsureCreated();
            taskPane.Visible = !taskPane.Visible;
            OfficeAgentLog.Info("taskpane", "visibility.toggled", $"Task pane visible: {taskPane.Visible}.");
            if (taskPane.Visible)
            {
                PublishCurrentSelectionContext();
            }
        }

        public void Show()
        {
            EnsureCreated();
            taskPane.Visible = true;
            OfficeAgentLog.Info("taskpane", "visibility.shown", "Task pane shown.");
            PublishCurrentSelectionContext();
        }

        public void PublishSelectionContext(SelectionContext selectionContext)
        {
            hostControl?.PublishSelectionContext(selectionContext);
        }

        internal bool IsVisible => taskPane?.Visible == true;

        private void EnsureCreated()
        {
            if (taskPane != null)
            {
                return;
            }

            hostControl = new TaskPaneHostControl(sessionStore, settingsStore, excelContextService, excelCommandExecutor, agentOrchestrator);
            taskPane = addIn.CustomTaskPanes.Add(hostControl, "OfficeAgent");
            taskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
            taskPane.Width = 800;
            taskPane.Visible = false;
            OfficeAgentLog.Info("taskpane", "created", "Custom task pane created.");
        }

        private void PublishCurrentSelectionContext()
        {
            PublishSelectionContext(excelContextService.GetCurrentSelectionContext());
        }
    }
}
