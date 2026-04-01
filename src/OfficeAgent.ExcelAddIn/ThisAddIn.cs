using System;
using System.IO;
using OfficeAgent.Core.Diagnostics;
using OfficeAgent.Core.Orchestration;
using OfficeAgent.Core.Services;
using OfficeAgent.Core.Skills;
using OfficeAgent.ExcelAddIn.Excel;
using OfficeAgent.ExcelAddIn.TaskPane;
using OfficeAgent.Infrastructure.Diagnostics;
using OfficeAgent.Infrastructure.Http;
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
        internal IExcelCommandExecutor ExcelCommandExecutor { get; private set; }
        internal IAgentOrchestrator AgentOrchestrator { get; private set; }
        internal ExcelFocusCoordinator ExcelFocusCoordinator { get; private set; }

        private bool isRestoringWorksheetFocus;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            var appDataDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OfficeAgent");
            var logSink = new FileLogSink(Path.Combine(appDataDirectory, "logs", "officeagent.log"));
            OfficeAgentLog.Configure(logSink.Write);
            OfficeAgentLog.Info("host", "startup.begin", "Starting OfficeAgent Excel add-in.");
            SessionStore = new FileSessionStore(Path.Combine(appDataDirectory, "sessions"));
            SettingsStore = new FileSettingsStore(
                Path.Combine(appDataDirectory, "settings.json"),
                new DpapiSecretProtector());

            var sharedCookies = new SharedCookieContainer();
            var cookieStore = new FileCookieStore(
                Path.Combine(appDataDirectory, "cookies.json"),
                new DpapiSecretProtector());
            cookieStore.Load(sharedCookies.Container);

            // Set SSO domain from settings for login status checks.
            var initialSettings = SettingsStore.Load();
            if (!string.IsNullOrWhiteSpace(initialSettings.SsoUrl))
            {
                try
                {
                    sharedCookies.SsoDomain = new Uri(initialSettings.SsoUrl).Host;
                }
                catch (UriFormatException)
                {
                    sharedCookies.SsoDomain = string.Empty;
                }
            }

            ExcelContextService = new ExcelSelectionContextService(Application);
            ExcelCommandExecutor = new ExcelInteropAdapter(Application, ExcelContextService);
            ExcelFocusCoordinator = new ExcelFocusCoordinator(Application);
            var skillRegistry = new SkillRegistry(
                new UploadDataSkill(ExcelCommandExecutor, new BusinessApiClient(() => SettingsStore.Load(), cookieContainer: sharedCookies.Container)));
            var fetchClient = new AgentFetchClient(() => SettingsStore.Load(), cookieContainer: sharedCookies.Container);
            AgentOrchestrator = new AgentOrchestrator(
                skillRegistry,
                ExcelContextService,
                ExcelCommandExecutor,
                new LlmPlannerClient(SettingsStore),
                new PlanExecutor(ExcelCommandExecutor, skillRegistry),
                fetchClient,
                () => SettingsStore.Load());
            TaskPaneController = new TaskPaneController(this, SessionStore, SettingsStore, ExcelContextService, ExcelCommandExecutor, AgentOrchestrator, sharedCookies, cookieStore);
            Application.SheetSelectionChange += Application_SheetSelectionChange;
            OfficeAgentLog.Info("host", "startup.completed", "OfficeAgent Excel add-in started.");
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Application.SheetSelectionChange -= Application_SheetSelectionChange;
            OfficeAgentLog.Info("host", "shutdown", "OfficeAgent Excel add-in stopped.");
            OfficeAgentLog.Reset();
        }

        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        private void Application_SheetSelectionChange(object sh, ExcelInterop.Range target)
        {
            OfficeAgentLog.Info("excel", "selection.changed", "Excel selection changed.");
            TaskPaneController?.PublishSelectionContext(ExcelContextService.GetCurrentSelectionContext());
            RestoreWorksheetFocus(target);
        }

        private void RestoreWorksheetFocus(ExcelInterop.Range target)
        {
            if (isRestoringWorksheetFocus || TaskPaneController?.IsVisible != true || ExcelFocusCoordinator == null)
            {
                return;
            }

            try
            {
                isRestoringWorksheetFocus = true;
                ExcelFocusCoordinator.RestoreWorksheetFocus(() => target?.Activate());
            }
            finally
            {
                isRestoringWorksheetFocus = false;
            }
        }
    }
}
