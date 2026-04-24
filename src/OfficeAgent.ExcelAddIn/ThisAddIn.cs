using System;
using System.Globalization;
using System.IO;
using Microsoft.Office.Core;
using OfficeAgent.Core.Diagnostics;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Orchestration;
using OfficeAgent.Core.Services;
using OfficeAgent.Core.Skills;
using OfficeAgent.Core.Sync;
using OfficeAgent.Core.Templates;
using OfficeAgent.ExcelAddIn.Excel;
using OfficeAgent.ExcelAddIn.Localization;
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
        internal SharedCookieContainer SharedCookies { get; private set; }
        internal FileCookieStore CookieStore { get; private set; }
        internal ISystemConnector CurrentBusinessConnector { get; private set; }
        internal ISystemConnectorRegistry SystemConnectorRegistry { get; private set; }
        internal IWorksheetMetadataStore WorksheetMetadataStore { get; private set; }
        internal WorksheetSyncService WorksheetSyncService { get; private set; }
        internal WorksheetSyncExecutionService WorksheetSyncExecutionService { get; private set; }
        internal RibbonSyncController RibbonSyncController { get; private set; }
        internal ITemplateStore TemplateStore { get; private set; }
        internal ITemplateCatalog TemplateCatalog { get; private set; }
        internal RibbonTemplateController RibbonTemplateController { get; private set; }
        internal Func<AppSettings, string> GetResolvedUiLocale { get; private set; }
        internal Localization.HostLocalizedStrings HostLocalizedStrings => GetHostLocalizedStrings();

        private bool isRestoringWorksheetFocus;
        private string lastProjectRefreshSheetName = string.Empty;

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
            var uiLocaleResolver = new UiLocaleResolver(GetExcelUiLocale);
            GetResolvedUiLocale = settings => uiLocaleResolver.Resolve(settings ?? SettingsStore.Load());

            SharedCookies = new SharedCookieContainer();
            CookieStore = new FileCookieStore(
                Path.Combine(appDataDirectory, "cookies.json"),
                new DpapiSecretProtector());
            CookieStore.Load(SharedCookies.Container);

            // Set SSO domain from settings for login status checks.
            var initialSettings = SettingsStore.Load();
            if (!string.IsNullOrWhiteSpace(initialSettings.SsoUrl))
            {
                try
                {
                    SharedCookies.SsoDomain = new Uri(initialSettings.SsoUrl).Host;
                }
                catch (UriFormatException)
                {
                    SharedCookies.SsoDomain = string.Empty;
                }
            }

            ExcelContextService = new ExcelSelectionContextService(Application);
            ExcelCommandExecutor = new ExcelInteropAdapter(Application, ExcelContextService);
            ExcelFocusCoordinator = new ExcelFocusCoordinator(Application);
            var skillRegistry = new SkillRegistry(
                new UploadDataSkill(ExcelCommandExecutor, new BusinessApiClient(() => SettingsStore.Load(), cookieContainer: SharedCookies.Container)));
            var fetchClient = new AgentFetchClient(() => SettingsStore.Load(), cookieContainer: SharedCookies.Container);
            AgentOrchestrator = new AgentOrchestrator(
                skillRegistry,
                ExcelContextService,
                ExcelCommandExecutor,
                new LlmPlannerClient(SettingsStore),
                new PlanExecutor(ExcelCommandExecutor, skillRegistry),
                fetchClient,
                () => SettingsStore.Load());
            CurrentBusinessConnector = new CurrentBusinessSystemConnector(() => SettingsStore.Load(), cookieContainer: SharedCookies.Container);
            SystemConnectorRegistry = new SystemConnectorRegistry(new[] { CurrentBusinessConnector });
            WorksheetMetadataStore = new WorksheetMetadataStore(new ExcelWorkbookMetadataAdapter(Application));
            WorksheetSyncService = new WorksheetSyncService(
                SystemConnectorRegistry,
                WorksheetMetadataStore,
                new WorksheetChangeTracker(),
                new SyncOperationPreviewFactory());
            WorksheetSyncExecutionService = new WorksheetSyncExecutionService(
                WorksheetSyncService,
                WorksheetMetadataStore,
                new ExcelVisibleSelectionReader(Application),
                new ExcelWorksheetGridAdapter(Application),
                new SyncOperationPreviewFactory());
            RibbonSyncController = new RibbonSyncController(
                WorksheetMetadataStore,
                WorksheetSyncService,
                GetActiveWorksheetName,
                WorksheetSyncExecutionService,
                new Dialogs.RibbonSyncDialogService(),
                () => Globals.Ribbons.AgentRibbon?.BeginLoginFlow(refreshProjectsAfterSuccess: false));
            TemplateStore = new LocalJsonTemplateStore(Path.Combine(appDataDirectory, "templates"));
            TemplateCatalog = new WorksheetTemplateCatalog(
                SystemConnectorRegistry,
                WorksheetMetadataStore,
                (IWorksheetTemplateBindingStore)WorksheetMetadataStore,
                TemplateStore);
            RibbonTemplateController = new RibbonTemplateController(
                TemplateCatalog,
                GetActiveWorksheetName);
            RibbonSyncController.RefreshActiveProjectFromSheetMetadata();
            RibbonTemplateController.RefreshActiveTemplateStateFromSheetMetadata();
            Globals.Ribbons.AgentRibbon?.BindToControllersAndRefresh();
            lastProjectRefreshSheetName = GetActiveWorksheetName();
            TaskPaneController = new TaskPaneController(this, SessionStore, SettingsStore, ExcelContextService, ExcelCommandExecutor, AgentOrchestrator, SharedCookies, CookieStore, GetResolvedUiLocale);
            Application.WorkbookActivate += Application_WorkbookActivate;
            Application.SheetActivate += Application_SheetActivate;
            Application.SheetSelectionChange += Application_SheetSelectionChange;
            Application.SheetChange += Application_SheetChange;
            OfficeAgentLog.Info("host", "startup.completed", "OfficeAgent Excel add-in started.");
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Application.WorkbookActivate -= Application_WorkbookActivate;
            Application.SheetActivate -= Application_SheetActivate;
            Application.SheetSelectionChange -= Application_SheetSelectionChange;
            Application.SheetChange -= Application_SheetChange;
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
            var sheetName = GetWorksheetName(sh);
            var activeSheetName = GetActiveWorksheetName();
            if (!string.Equals(sheetName, activeSheetName, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            OfficeAgentLog.Info("excel", "selection.changed", "Excel selection changed.");

            if (!string.Equals(lastProjectRefreshSheetName, sheetName, StringComparison.OrdinalIgnoreCase))
            {
                RibbonSyncController?.RefreshProjectFromSheetMetadata(sheetName);
                RibbonTemplateController?.RefreshTemplateState(sheetName);
                lastProjectRefreshSheetName = sheetName;
            }

            TaskPaneController?.PublishSelectionContext(ExcelContextService.GetCurrentSelectionContext());
            RestoreWorksheetFocus(target);
        }

        private void Application_SheetActivate(object sh)
        {
            var sheetName = GetWorksheetName(sh);
            RibbonSyncController?.RefreshProjectFromSheetMetadata(sheetName);
            RibbonTemplateController?.RefreshTemplateState(sheetName);
            lastProjectRefreshSheetName = sheetName;
        }

        private void Application_WorkbookActivate(ExcelInterop.Workbook wb)
        {
            RibbonSyncController?.InvalidateRefreshState();
            RibbonTemplateController?.InvalidateRefreshState();
            RibbonSyncController?.RefreshActiveProjectFromSheetMetadata();
            RibbonTemplateController?.RefreshActiveTemplateStateFromSheetMetadata();
            lastProjectRefreshSheetName = GetActiveWorksheetName();
        }

        private void Application_SheetChange(object sh, ExcelInterop.Range target)
        {
            var sheetName = GetWorksheetName(sh);
            if (!string.Equals(sheetName, "ISDP_Setting", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            var metadataStore = WorksheetMetadataStore as OfficeAgent.ExcelAddIn.Excel.WorksheetMetadataStore;
            metadataStore.InvalidateCache();
            RibbonSyncController?.InvalidateRefreshState();
            RibbonTemplateController?.InvalidateRefreshState();
            lastProjectRefreshSheetName = string.Empty;
        }

        private string GetActiveWorksheetName()
        {
            try
            {
                var worksheet = Application?.ActiveSheet as ExcelInterop.Worksheet;
                return worksheet?.Name ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string GetWorksheetName(object sheet)
        {
            var worksheet = sheet as ExcelInterop.Worksheet;
            return worksheet?.Name ?? string.Empty;
        }

        private string GetExcelUiLocale()
        {
            try
            {
                var languageSettings = Application?.LanguageSettings;
                if (languageSettings == null)
                {
                    return string.Empty;
                }

                var languageId = languageSettings.LanguageID[MsoAppLanguageID.msoLanguageIDUI];
                if (languageId <= 0)
                {
                    return string.Empty;
                }

                return CultureInfo.GetCultureInfo(languageId).Name;
            }
            catch (CultureNotFoundException)
            {
                return string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        internal Localization.HostLocalizedStrings GetHostLocalizedStrings(AppSettings settings = null)
        {
            var resolvedLocale = GetResolvedUiLocale?.Invoke(settings ?? SettingsStore?.Load() ?? new AppSettings()) ?? "en";
            return Localization.HostLocalizedStrings.ForLocale(resolvedLocale);
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
