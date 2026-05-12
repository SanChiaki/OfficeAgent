using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Microsoft.Office.Core;
using OfficeAgent.Core.Analytics;
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
using OfficeAgent.Infrastructure.Analytics;
using OfficeAgent.Infrastructure.Diagnostics;
using OfficeAgent.Infrastructure.Http;
using OfficeAgent.Infrastructure.Security;
using OfficeAgent.Infrastructure.Storage;
using OfficeAgent.ExcelAddIn.Analytics;
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
        internal IAnalyticsService AnalyticsService { get; private set; }
        internal ExcelFocusCoordinator ExcelFocusCoordinator { get; private set; }
        internal SharedCookieContainer SharedCookies { get; private set; }
        internal FileCookieStore CookieStore { get; private set; }
        internal ISystemConnector CurrentBusinessConnector { get; private set; }
        internal ISystemConnectorRegistry SystemConnectorRegistry { get; private set; }
        internal IWorksheetMetadataStore WorksheetMetadataStore { get; private set; }
        internal WorksheetSyncService WorksheetSyncService { get; private set; }
        internal WorksheetSyncExecutionService WorksheetSyncExecutionService { get; private set; }
        internal WorksheetPendingEditTracker WorksheetPendingEditTracker { get; private set; }
        internal IWorksheetChangeLogStore WorksheetChangeLogStore { get; private set; }
        internal RibbonSyncController RibbonSyncController { get; private set; }
        internal ITemplateStore TemplateStore { get; private set; }
        internal ITemplateCatalog TemplateCatalog { get; private set; }
        internal RibbonTemplateController RibbonTemplateController { get; private set; }
        internal Func<AppSettings, string> GetResolvedUiLocale { get; private set; }
        internal Localization.HostLocalizedStrings HostLocalizedStrings => GetHostLocalizedStrings();

        private const int MaxTrackedRangeCellCount = 10000;

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
            var initialSettings = SettingsStore.Load();
            AnalyticsService = string.IsNullOrWhiteSpace(initialSettings.AnalyticsBaseUrl)
                ? NoopAnalyticsService.Instance
                : new OfficeAgent.Core.Analytics.AnalyticsService(new InsertLogAnalyticsSink(() => SettingsStore.Load()));
            var uiLocaleResolver = new UiLocaleResolver(GetExcelUiLocale);
            GetResolvedUiLocale = settings => uiLocaleResolver.Resolve(settings ?? SettingsStore.Load());

            SharedCookies = new SharedCookieContainer();
            CookieStore = new FileCookieStore(
                Path.Combine(appDataDirectory, "cookies.json"),
                new DpapiSecretProtector());
            CookieStore.Load(SharedCookies.Container);

            // Set SSO domain from settings for login status checks.
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
            CurrentBusinessConnector = new CurrentBusinessSystemConnector(() => SettingsStore.Load(), cookieContainer: SharedCookies.Container, analyticsService: AnalyticsService);
            SystemConnectorRegistry = new SystemConnectorRegistry(new[] { CurrentBusinessConnector });
            WorksheetMetadataStore = new WorksheetMetadataStore(new ExcelWorkbookMetadataAdapter(Application));
            WorksheetSyncService = new WorksheetSyncService(
                SystemConnectorRegistry,
                WorksheetMetadataStore,
                new WorksheetChangeTracker(),
                new SyncOperationPreviewFactory(), AnalyticsService);
            var worksheetGridAdapter = new ExcelWorksheetGridAdapter(Application);
            WorksheetChangeLogStore = new WorksheetChangeLogStore(worksheetGridAdapter);
            WorksheetPendingEditTracker = new WorksheetPendingEditTracker();
            WorksheetSyncExecutionService = new WorksheetSyncExecutionService(
                WorksheetSyncService,
                WorksheetMetadataStore,
                new ExcelVisibleSelectionReader(Application),
                worksheetGridAdapter,
                new SyncOperationPreviewFactory(),
                WorksheetChangeLogStore,
                WorksheetPendingEditTracker,
                new AiColumnMappingClient(SettingsStore));
            RibbonSyncController = new RibbonSyncController(
                WorksheetMetadataStore,
                WorksheetSyncService,
                GetActiveWorksheetName,
                WorksheetSyncExecutionService,
                new Dialogs.RibbonSyncDialogService(),
                () => Globals.Ribbons.AgentRibbon?.BeginLoginFlow(refreshProjectsAfterSuccess: false),
                AnalyticsService);
            TemplateStore = new LocalJsonTemplateStore(Path.Combine(appDataDirectory, "templates"));
            TemplateCatalog = new WorksheetTemplateCatalog(
                SystemConnectorRegistry,
                WorksheetMetadataStore,
                (IWorksheetTemplateBindingStore)WorksheetMetadataStore,
                TemplateStore);
            RibbonTemplateController = new RibbonTemplateController(
                TemplateCatalog,
                GetActiveWorksheetName,
                new Dialogs.RibbonTemplateDialogService(),
                AnalyticsService);
            RibbonSyncController.RefreshActiveProjectFromSheetMetadata();
            RibbonTemplateController.RefreshActiveTemplateStateFromSheetMetadata();
            Globals.Ribbons.AgentRibbon?.BindToControllersAndRefresh();
            lastProjectRefreshSheetName = GetActiveWorksheetName();
            TaskPaneController = new TaskPaneController(this, SessionStore, SettingsStore, ExcelContextService, ExcelCommandExecutor, AgentOrchestrator, SharedCookies, CookieStore, GetResolvedUiLocale, AnalyticsService);
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

            if (ShouldTrackBusinessSheetChange(sheetName) && WorksheetPendingEditTracker != null)
            {
                WorksheetPendingEditTracker.CaptureBeforeValues(sheetName, ReadWorksheetCellValues(target));
            }

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
            if (IsSettingsSheet(sheetName))
            {
                var metadataStore = WorksheetMetadataStore as OfficeAgent.ExcelAddIn.Excel.WorksheetMetadataStore;
                metadataStore.InvalidateCache();
                RibbonSyncController?.InvalidateRefreshState();
                RibbonTemplateController?.InvalidateRefreshState();
                lastProjectRefreshSheetName = string.Empty;
                return;
            }

            if (ShouldTrackBusinessSheetChange(sheetName) && WorksheetPendingEditTracker != null)
            {
                WorksheetPendingEditTracker.MarkChanged(sheetName, ReadWorksheetCellAddresses(target));
            }
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

        private string GetActiveWorkbookName()
        {
            try
            {
                return Application?.ActiveWorkbook?.Name ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        internal RibbonAnalyticsHelper CreateRibbonAnalyticsHelper()
        {
            return new RibbonAnalyticsHelper(
                AnalyticsService,
                () => RibbonSyncController?.ActiveBinding,
                GetActiveWorksheetName,
                GetActiveWorkbookName,
                () => HostLocalizedStrings);
        }

        private static string GetWorksheetName(object sheet)
        {
            var worksheet = sheet as ExcelInterop.Worksheet;
            return worksheet?.Name ?? string.Empty;
        }

        private static bool ShouldTrackBusinessSheetChange(string sheetName)
        {
            return !IsSettingsSheet(sheetName) && !IsSyncLogSheet(sheetName);
        }

        private static bool IsSettingsSheet(string sheetName)
        {
            return MetadataWorksheetNames.IsMetadataWorksheet(sheetName);
        }

        private static bool IsSyncLogSheet(string sheetName)
        {
            return string.Equals(sheetName, "xISDP_Log", StringComparison.OrdinalIgnoreCase);
        }

        private static IReadOnlyList<WorksheetCellValue> ReadWorksheetCellValues(ExcelInterop.Range target)
        {
            var result = new List<WorksheetCellValue>();
            if (target == null || IsRangeTooLarge(target))
            {
                return result;
            }

            foreach (ExcelInterop.Range cell in target.Cells)
            {
                if (cell == null)
                {
                    continue;
                }

                result.Add(new WorksheetCellValue
                {
                    Row = cell.Row,
                    Column = cell.Column,
                    Text = Convert.ToString(cell.Text) ?? string.Empty,
                });
            }

            return result;
        }

        private static IReadOnlyList<WorksheetCellAddress> ReadWorksheetCellAddresses(ExcelInterop.Range target)
        {
            var result = new List<WorksheetCellAddress>();
            if (target == null || IsRangeTooLarge(target))
            {
                return result;
            }

            foreach (ExcelInterop.Range cell in target.Cells)
            {
                if (cell == null)
                {
                    continue;
                }

                result.Add(new WorksheetCellAddress
                {
                    Row = cell.Row,
                    Column = cell.Column,
                });
            }

            return result;
        }

        private static bool IsRangeTooLarge(ExcelInterop.Range target)
        {
            try
            {
                return Convert.ToDouble(target.Cells.CountLarge) > MaxTrackedRangeCellCount;
            }
            catch
            {
                return true;
            }
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
