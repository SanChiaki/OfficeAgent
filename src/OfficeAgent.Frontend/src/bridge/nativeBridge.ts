import type {
  AppSettings,
  AnalyticsPayload,
  AnalyticsResult,
  BridgeErrorPayload,
  BridgeEventEnvelope,
  BridgeRequestEnvelope,
  BridgeResponseEnvelope,
  AgentPlan,
  AgentRequestEnvelope,
  AgentResult,
  ExcelCommand,
  ExcelCommandPreview,
  ExcelCommandResult,
  HostContext,
  LoginResult,
  LoginStatus,
  SelectionContext,
  SessionState,
  SkillRequestEnvelope,
  SkillResult,
  UploadPreview,
  PingPayload,
  WebViewHostLike,
  WebViewMessageEventLike,
} from '../types/bridge';
import { getUiStrings } from '../i18n/uiStrings';

const BRIDGE_TYPES = {
  ping: 'bridge.ping',
  getSettings: 'bridge.getSettings',
  getHostContext: 'bridge.getHostContext',
  getSelectionContext: 'bridge.getSelectionContext',
  selectionContextChanged: 'bridge.selectionContextChanged',
  getSessions: 'bridge.getSessions',
  saveSessions: 'bridge.saveSessions',
  saveSettings: 'bridge.saveSettings',
  executeExcelCommand: 'bridge.executeExcelCommand',
  runSkill: 'bridge.runSkill',
  runAgent: 'bridge.runAgent',
  login: 'bridge.login',
  logout: 'bridge.logout',
  getLoginStatus: 'bridge.getLoginStatus',
  trackAnalytics: 'bridge.trackAnalytics',
} as const;

const BROWSER_PREVIEW_PING: PingPayload = {
  host: 'browser-preview',
  version: 'dev',
};

const BROWSER_PREVIEW_SETTINGS: AppSettings = {
  apiKey: '',
  baseUrl: 'https://api.example.com',
  businessBaseUrl: '',
  model: 'gpt-5-mini',
  apiFormat: 'openai-compatible',
  ssoUrl: '',
  ssoLoginSuccessPath: '',
  uiLanguageOverride: 'system',
};

const BROWSER_PREVIEW_SELECTION_CONTEXT: SelectionContext = {
  hasSelection: true,
  workbookName: 'Browser Preview.xlsx',
  sheetName: 'Sheet1',
  address: 'A1:C4',
  rowCount: 4,
  columnCount: 3,
  isContiguous: true,
  headerPreview: ['Name', 'Region', 'Amount'],
  sampleRows: [
    ['Project A', 'CN', '42'],
    ['Project B', 'US', '36'],
  ],
  warningMessage: null,
};

const BROWSER_PREVIEW_SESSIONS: SessionState = {
  activeSessionId: 'browser-preview-session',
  sessions: [
    {
      id: 'browser-preview-session',
      title: 'Browser preview',
      isSystemUntitled: false,
      createdAtUtc: '2026-03-29T00:00:00.0000000Z',
      updatedAtUtc: '2026-03-29T00:00:00.0000000Z',
      messages: [],
    },
  ],
};

class NativeBridgeError extends Error {
  public readonly code: string;

  constructor(error: BridgeErrorPayload) {
    super(error.message);
    this.code = error.code;
    this.name = 'NativeBridgeError';
  }
}

type PendingRequest = {
  resolve: (value: unknown) => void;
  reject: (reason?: unknown) => void;
};

type SelectionContextListener = (payload: SelectionContext) => void;

export class NativeBridge {
  private readonly webView?: WebViewHostLike;
  private browserPreviewSettings: AppSettings = { ...BROWSER_PREVIEW_SETTINGS };
  private readonly pendingRequests = new Map<string, PendingRequest>();
  private readonly selectionContextListeners = new Set<SelectionContextListener>();
  private readonly handleMessage = (event: WebViewMessageEventLike) => {
    const response = event.data;
    if (isBridgeEventEnvelope(response) && response.type === BRIDGE_TYPES.selectionContextChanged) {
      const payload = response.payload as SelectionContext | undefined;
      if (payload) {
        this.selectionContextListeners.forEach((listener) => listener(payload));
      }

      return;
    }

    if (!isBridgeResponseEnvelope(response)) {
      return;
    }

    const pending = this.pendingRequests.get(response.requestId);
    if (!pending) {
      return;
    }

    this.pendingRequests.delete(response.requestId);

    if (response.ok) {
      pending.resolve(response.payload);
      return;
    }

    pending.reject(new NativeBridgeError(normalizeError(response.error)));
  };

  constructor(webView: WebViewHostLike | undefined = getWebViewHost()) {
    this.webView = webView;
    this.webView?.addEventListener('message', this.handleMessage);
  }

  dispose() {
    this.webView?.removeEventListener('message', this.handleMessage);
    this.pendingRequests.clear();
  }

  ping() {
    return this.invoke<void, PingPayload>(BRIDGE_TYPES.ping);
  }

  getSettings() {
    return this.invoke<void, AppSettings>(BRIDGE_TYPES.getSettings);
  }

  getHostContext() {
    return this.invoke<void, HostContext>(BRIDGE_TYPES.getHostContext);
  }

  getSelectionContext() {
    return this.invoke<void, SelectionContext>(BRIDGE_TYPES.getSelectionContext);
  }

  getSessions() {
    return this.invoke<void, SessionState>(BRIDGE_TYPES.getSessions);
  }

  saveSessions(payload: SessionState) {
    return this.invoke<SessionState, SessionState>(BRIDGE_TYPES.saveSessions, payload);
  }

  saveSettings(payload: AppSettings) {
    return this.invoke<AppSettings, AppSettings>(BRIDGE_TYPES.saveSettings, payload);
  }

  executeExcelCommand(payload: ExcelCommand) {
    return this.invoke<ExcelCommand, ExcelCommandResult>(BRIDGE_TYPES.executeExcelCommand, payload);
  }

  runSkill(payload: SkillRequestEnvelope) {
    return this.invoke<SkillRequestEnvelope, SkillResult>(BRIDGE_TYPES.runSkill, payload);
  }

  runAgent(payload: AgentRequestEnvelope) {
    return this.invoke<AgentRequestEnvelope, AgentResult>(BRIDGE_TYPES.runAgent, payload);
  }

  login(payload: { ssoUrl: string; ssoLoginSuccessPath?: string }) {
    return this.invoke<{ ssoUrl: string; ssoLoginSuccessPath?: string }, LoginResult>(BRIDGE_TYPES.login, payload);
  }

  logout() {
    return this.invoke<void, LoginResult>(BRIDGE_TYPES.logout);
  }

  getLoginStatus() {
    return this.invoke<void, LoginStatus>(BRIDGE_TYPES.getLoginStatus);
  }

  trackAnalytics(payload: AnalyticsPayload) {
    return this.invoke<AnalyticsPayload, AnalyticsResult>(BRIDGE_TYPES.trackAnalytics, payload);
  }

  onSelectionContextChanged(listener: SelectionContextListener) {
    this.selectionContextListeners.add(listener);
    return () => {
      this.selectionContextListeners.delete(listener);
    };
  }

  private invoke<TPayload, TResult>(type: string, payload?: TPayload): Promise<TResult> {
    if (!this.webView) {
      const previewLocale = resolveBrowserPreviewLocale(this.browserPreviewSettings.uiLanguageOverride);
      const previewStrings = getUiStrings(previewLocale);

      if (type === BRIDGE_TYPES.ping) {
        return Promise.resolve(BROWSER_PREVIEW_PING as TResult);
      }

      if (type === BRIDGE_TYPES.getSettings) {
        return Promise.resolve(this.browserPreviewSettings as TResult);
      }

      if (type === BRIDGE_TYPES.getHostContext) {
        return Promise.resolve({
          resolvedUiLocale: previewLocale,
          uiLanguageOverride: this.browserPreviewSettings.uiLanguageOverride ?? 'system',
        } as TResult);
      }

      if (type === BRIDGE_TYPES.getSelectionContext) {
        return Promise.resolve(BROWSER_PREVIEW_SELECTION_CONTEXT as TResult);
      }

      if (type === BRIDGE_TYPES.getSessions) {
        return Promise.resolve(BROWSER_PREVIEW_SESSIONS as TResult);
      }

      if (type === BRIDGE_TYPES.saveSessions) {
        return Promise.resolve((payload ?? BROWSER_PREVIEW_SESSIONS) as TResult);
      }

      if (type === BRIDGE_TYPES.saveSettings) {
        this.browserPreviewSettings = {
          apiKey: typeof (payload as AppSettings | undefined)?.apiKey === 'string' ? (payload as AppSettings).apiKey : '',
          baseUrl: typeof (payload as AppSettings | undefined)?.baseUrl === 'string'
            ? (payload as AppSettings).baseUrl
            : BROWSER_PREVIEW_SETTINGS.baseUrl,
          businessBaseUrl: typeof (payload as AppSettings | undefined)?.businessBaseUrl === 'string'
            ? (payload as AppSettings).businessBaseUrl
            : BROWSER_PREVIEW_SETTINGS.businessBaseUrl,
          model: typeof (payload as AppSettings | undefined)?.model === 'string'
            ? (payload as AppSettings).model
            : BROWSER_PREVIEW_SETTINGS.model,
          apiFormat: isValidApiFormat((payload as AppSettings | undefined)?.apiFormat)
            ? (payload as AppSettings).apiFormat
            : BROWSER_PREVIEW_SETTINGS.apiFormat,
          ssoUrl: typeof (payload as AppSettings | undefined)?.ssoUrl === 'string'
            ? (payload as AppSettings).ssoUrl
            : BROWSER_PREVIEW_SETTINGS.ssoUrl,
          ssoLoginSuccessPath: typeof (payload as AppSettings | undefined)?.ssoLoginSuccessPath === 'string'
            ? (payload as AppSettings).ssoLoginSuccessPath
            : BROWSER_PREVIEW_SETTINGS.ssoLoginSuccessPath,
          uiLanguageOverride: isValidUiLanguageOverride((payload as AppSettings | undefined)?.uiLanguageOverride)
            ? (payload as AppSettings).uiLanguageOverride
            : BROWSER_PREVIEW_SETTINGS.uiLanguageOverride,
        };

        return Promise.resolve(this.browserPreviewSettings as TResult);
      }

      if (type === BRIDGE_TYPES.executeExcelCommand) {
        try {
          return Promise.resolve(
            createBrowserPreviewCommandResult(
              validateBrowserPreviewCommand(payload as ExcelCommand),
              previewLocale,
            ) as TResult,
          );
        } catch (error) {
          return Promise.reject(error);
        }
      }

      if (type === BRIDGE_TYPES.runSkill) {
        try {
          return Promise.resolve(
            createBrowserPreviewSkillResult(
              validateBrowserPreviewSkill(payload as SkillRequestEnvelope),
              previewLocale,
            ) as TResult,
          );
        } catch (error) {
          return Promise.reject(error);
        }
      }

      if (type === BRIDGE_TYPES.runAgent) {
        try {
          return Promise.resolve(
            createBrowserPreviewAgentResult(
              validateBrowserPreviewAgent(payload as AgentRequestEnvelope),
              previewLocale,
            ) as TResult,
          );
        } catch (error) {
          return Promise.reject(error);
        }
      }

      if (type === BRIDGE_TYPES.getLoginStatus) {
        return Promise.resolve({ isLoggedIn: false, ssoUrl: '' } as TResult);
      }

      if (type === BRIDGE_TYPES.login) {
        return Promise.resolve({ success: false, error: previewStrings.browserPreviewLoginUnavailable } as TResult);
      }

      if (type === BRIDGE_TYPES.logout) {
        return Promise.resolve({ success: true } as TResult);
      }

      if (type === BRIDGE_TYPES.trackAnalytics) {
        return Promise.resolve({ tracked: false } as TResult);
      }

      return Promise.reject(
        new NativeBridgeError({
          code: 'bridge_unavailable',
          message: 'Native bridge is only available inside the Excel task pane.',
        }),
      );
    }

    const requestId = createRequestId();
    const request: BridgeRequestEnvelope<TPayload> = { type, requestId };
    if (payload !== undefined) {
      request.payload = payload;
    }

    return new Promise<TResult>((resolve, reject) => {
      this.pendingRequests.set(requestId, {
        resolve: (value) => resolve(value as TResult),
        reject,
      });
      this.webView?.postMessage(request);
    });
  }
}

export const nativeBridge = new NativeBridge();

function getWebViewHost() {
  return window.chrome?.webview;
}

function createRequestId() {
  if (globalThis.crypto?.randomUUID) {
    return globalThis.crypto.randomUUID();
  }

  return `req-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function isBridgeResponseEnvelope(value: unknown): value is BridgeResponseEnvelope {
  if (!value || typeof value !== 'object') {
    return false;
  }

  const candidate = value as Record<string, unknown>;
  return (
    typeof candidate.type === 'string' &&
    typeof candidate.requestId === 'string' &&
    typeof candidate.ok === 'boolean'
  );
}

function isBridgeEventEnvelope(value: unknown): value is BridgeEventEnvelope {
  if (!value || typeof value !== 'object') {
    return false;
  }

  const candidate = value as Record<string, unknown>;
  return typeof candidate.type === 'string' && !('requestId' in candidate) && !('ok' in candidate);
}

function normalizeError(error: BridgeErrorPayload | undefined): BridgeErrorPayload {
  if (error?.code && error.message) {
    return error;
  }

  return {
    code: 'bridge_error',
    message: 'The native host returned an invalid error payload.',
  };
}

function isValidUiLanguageOverride(value: unknown): value is NonNullable<AppSettings['uiLanguageOverride']> {
  return value === 'system' || value === 'zh' || value === 'en';
}

function isValidApiFormat(value: unknown): value is AppSettings['apiFormat'] {
  return value === 'openai-compatible' || value === 'anthropic-messages';
}

function resolveBrowserPreviewLocale(uiLanguageOverride: AppSettings['uiLanguageOverride']): HostContext['resolvedUiLocale'] {
  return uiLanguageOverride === 'zh' ? 'zh' : 'en';
}

function createBrowserPreviewCommandResult(command: ExcelCommand, locale: HostContext['resolvedUiLocale']): ExcelCommandResult {
  const strings = getUiStrings(locale);

  switch (command.commandType) {
    case 'excel.readSelectionTable':
      return {
        commandType: command.commandType,
        requiresConfirmation: false,
        status: 'completed',
        message: strings.browserPreviewReadSelectionMessage('Sheet1', 'A1:C4'),
        table: {
          sheetName: 'Sheet1',
          address: 'A1:C4',
          headers: ['Name', 'Region', 'Amount'],
          rows: [
            ['Project A', 'CN', '42'],
            ['Project B', 'US', '36'],
          ],
        },
        selectionContext: BROWSER_PREVIEW_SELECTION_CONTEXT,
      };
    case 'excel.addWorksheet':
      return createBrowserPreviewWriteResult(command, locale, {
        previewTitle: strings.excelAddWorksheetPreviewTitle,
        previewSummary: strings.formatExcelAddWorksheetPreviewSummary(command.newSheetName ?? 'New Sheet'),
        completedMessage: strings.browserPreviewWorksheetCreatedMessage(command.newSheetName ?? 'New Sheet'),
      });
    case 'excel.renameWorksheet':
      return createBrowserPreviewWriteResult(command, locale, {
        previewTitle: strings.excelRenameWorksheetPreviewTitle,
        previewSummary: strings.formatExcelRenameWorksheetPreviewSummary(
          command.sheetName ?? 'Sheet1',
          command.newSheetName ?? 'Renamed Sheet',
        ),
        completedMessage: strings.browserPreviewWorksheetRenamedMessage(
          command.sheetName ?? 'Sheet1',
          command.newSheetName ?? 'Renamed Sheet',
        ),
      });
    case 'excel.deleteWorksheet':
      return createBrowserPreviewWriteResult(command, locale, {
        previewTitle: strings.excelDeleteWorksheetPreviewTitle,
        previewSummary: strings.formatExcelDeleteWorksheetPreviewSummary(command.sheetName ?? 'Sheet1'),
        completedMessage: strings.browserPreviewWorksheetDeletedMessage(command.sheetName ?? 'Sheet1'),
      });
    case 'excel.writeRange':
      return createBrowserPreviewWriteResult(command, locale, {
        previewTitle: strings.excelWriteRangePreviewTitle,
        previewSummary: strings.formatExcelWriteRangePreviewSummary(
          (command.values ?? []).length,
          (command.values ?? [])[0]?.length ?? 0,
          command.targetAddress ?? 'A1',
        ),
        completedMessage: strings.browserPreviewWriteRangeCompletedMessage(
          (command.values ?? []).length,
          command.targetAddress ?? 'A1',
        ),
        details: (command.values ?? []).slice(0, 3).map((row) => row.join(' | ')),
      });
    default:
      throw new NativeBridgeError({
        code: 'bridge_unavailable',
        message: strings.browserPreviewUnsupportedCommandMessage(command.commandType),
      });
  }
}

function createBrowserPreviewWriteResult(
  command: ExcelCommand,
  locale: HostContext['resolvedUiLocale'],
  options: {
    previewTitle: string;
    previewSummary: string;
    completedMessage: string;
    details?: string[];
  },
): ExcelCommandResult {
  const strings = getUiStrings(locale);

  if (!command.confirmed) {
    return {
      commandType: command.commandType,
      requiresConfirmation: true,
      status: 'preview',
      message: strings.browserPreviewExcelConfirmMessage,
      preview: {
        title: options.previewTitle,
        summary: options.previewSummary,
        details: options.details ?? [strings.formatWorkbookDetail('Browser Preview.xlsx')],
      },
      selectionContext: BROWSER_PREVIEW_SELECTION_CONTEXT,
    };
  }

  return {
    commandType: command.commandType,
    requiresConfirmation: false,
    status: 'completed',
    message: options.completedMessage,
    selectionContext: BROWSER_PREVIEW_SELECTION_CONTEXT,
  };
}

function validateBrowserPreviewCommand(command: ExcelCommand): ExcelCommand {
  if (!command?.commandType) {
    throw new NativeBridgeError({
      code: 'invalid_command',
      message: 'Excel commands must include a commandType.',
    });
  }

  if (command.commandType === 'excel.writeRange') {
    if (!command.targetAddress?.trim()) {
      throw new NativeBridgeError({
        code: 'invalid_command',
        message: 'Write range commands require a target address.',
      });
    }

    const trimmedTargetAddress = command.targetAddress.trim();
    if (trimmedTargetAddress.includes('!')) {
      const [, targetRangeAddress = ''] = trimmedTargetAddress.split('!', 2);
      if (!targetRangeAddress.trim()) {
        throw new NativeBridgeError({
          code: 'invalid_command',
          message: 'Write range commands must include a cell reference in the target address.',
        });
      }
    }

    if (!command.values?.length || !command.values[0]?.length) {
      throw new NativeBridgeError({
        code: 'invalid_command',
        message: 'Write range commands require at least one row and one column of values.',
      });
    }

    const expectedColumnCount = command.values[0].length;
    if (command.values.some((row) => row.length !== expectedColumnCount)) {
      throw new NativeBridgeError({
        code: 'invalid_command',
        message: 'Write range commands require a rectangular values payload.',
      });
    }

    if (command.sheetName?.trim() && command.targetAddress.includes('!')) {
      const [qualifiedSheetName] = command.targetAddress.split('!', 1);
      if (
        qualifiedSheetName.trim() &&
        command.sheetName.trim().toLowerCase() !== qualifiedSheetName.trim().toLowerCase()
      ) {
        throw new NativeBridgeError({
          code: 'invalid_command',
          message: 'Write range commands cannot specify conflicting sheet names.',
        });
      }
    }
  }

  if (command.commandType === 'excel.addWorksheet' && !command.newSheetName?.trim()) {
    throw new NativeBridgeError({
      code: 'invalid_command',
      message: 'Add worksheet commands require a new sheet name.',
    });
  }

  if (command.commandType === 'excel.renameWorksheet') {
    if (!command.sheetName?.trim() || !command.newSheetName?.trim()) {
      throw new NativeBridgeError({
        code: 'invalid_command',
        message: 'Rename worksheet commands require both the current and new sheet names.',
      });
    }

    if (command.sheetName.trim().toLowerCase() === command.newSheetName.trim().toLowerCase()) {
      throw new NativeBridgeError({
        code: 'invalid_command',
        message: 'Rename worksheet commands must change the worksheet name.',
      });
    }
  }

  if (command.commandType === 'excel.deleteWorksheet' && !command.sheetName?.trim()) {
    throw new NativeBridgeError({
      code: 'invalid_command',
      message: 'Delete worksheet commands require a sheet name.',
    });
  }

  return command;
}

function validateBrowserPreviewSkill(payload: SkillRequestEnvelope): SkillRequestEnvelope {
  if (!payload?.userInput?.trim()) {
    throw new NativeBridgeError({
      code: 'invalid_command',
      message: 'Skill execution requires user input.',
    });
  }

  if (payload.confirmed && !payload.uploadPreview) {
    throw new NativeBridgeError({
      code: 'invalid_command',
      message: 'upload_data confirmation requires an upload preview payload.',
    });
  }

  if (payload.confirmed && !hasCompleteUploadPreview(payload.uploadPreview)) {
    throw new NativeBridgeError({
      code: 'invalid_command',
      message: 'upload_data confirmation requires a complete preview payload.',
    });
  }

  return {
    ...payload,
    userInput: payload.userInput.trim(),
  };
}

function hasCompleteUploadPreview(preview: UploadPreview | undefined): preview is UploadPreview {
  return Boolean(
    preview &&
    typeof preview.projectName === 'string' &&
    preview.projectName.trim() &&
    typeof preview.sheetName === 'string' &&
    preview.sheetName.trim() &&
    typeof preview.address === 'string' &&
    preview.address.trim() &&
    Array.isArray(preview.headers) &&
    preview.headers.length > 0 &&
    Array.isArray(preview.rows) &&
    Array.isArray(preview.records),
  );
}

function extractResolvedBrowserPreviewProjectName(userInput: string): string {
  const trimmedInput = userInput.trim().replace(/^\/upload_data\s*/i, '');
  const chineseUploadTo = '\u4E0A\u4F20\u5230';
  const chineseIndex = trimmedInput.lastIndexOf(chineseUploadTo);
  if (chineseIndex >= 0) {
    const projectName = trimmedInput.slice(chineseIndex + chineseUploadTo.length).trim();
    if (projectName) {
      return projectName;
    }
  }

  const englishMatch = trimmedInput.match(/\bto\s+(.+)$/i);
  if (englishMatch?.[1]) {
    return englishMatch[1].trim();
  }

  throw new NativeBridgeError({
    code: 'invalid_command',
    message: 'upload_data requires a target project name.',
  });
}

function matchesUploadDataSkillInput(userInput: string): boolean {
  const trimmedInput = userInput.trim();
  return (
    trimmedInput.startsWith('/upload_data') ||
    trimmedInput.includes('\u4E0A\u4F20\u5230') ||
    /\bupload\b.+\bto\s+.+$/i.test(trimmedInput)
  );
}

function createBrowserPreviewSkillResult(payload: SkillRequestEnvelope, locale: HostContext['resolvedUiLocale']): SkillResult {
  const strings = getUiStrings(locale);

  if (!isUploadDataSkillInput(payload.userInput)) {
    return {
      route: 'chat',
      requiresConfirmation: false,
      status: 'completed',
      message: strings.browserPreviewChatFallback,
    };
  }

  const uploadPreview = payload.uploadPreview ?? buildBrowserPreviewUpload(payload.userInput);
  const preview: ExcelCommandPreview = {
    title: strings.uploadPreviewTitle,
    summary: strings.formatUploadPreviewSummary(uploadPreview.records.length, uploadPreview.projectName),
    details: [
      strings.formatUploadPreviewSourceDetail(uploadPreview.sheetName, uploadPreview.address),
      strings.formatUploadPreviewFieldsDetail(uploadPreview.headers),
    ],
  };

  if (!payload.confirmed) {
    return {
      route: 'skill',
      skillName: 'upload_data',
      requiresConfirmation: true,
      status: 'preview',
      message: strings.browserPreviewUploadReviewMessage(uploadPreview.projectName),
      preview,
      uploadPreview,
    };
  }

  return {
    route: 'skill',
    skillName: 'upload_data',
    requiresConfirmation: false,
    status: 'completed',
    message: strings.browserPreviewUploadCompletedMessage(uploadPreview.projectName, uploadPreview.records.length),
    preview,
    uploadPreview,
  };
}

function validateBrowserPreviewAgent(payload: AgentRequestEnvelope): AgentRequestEnvelope {
  if (!payload?.userInput?.trim()) {
    throw new NativeBridgeError({
      code: 'invalid_command',
      message: 'Agent execution requires user input.',
    });
  }

  if (payload.confirmed && !payload.plan) {
    throw new NativeBridgeError({
      code: 'invalid_command',
      message: 'Agent confirmation requires a frozen plan payload.',
    });
  }

  return {
    ...payload,
    userInput: payload.userInput.trim(),
  };
}

function createBrowserPreviewAgentResult(payload: AgentRequestEnvelope, locale: HostContext['resolvedUiLocale']): AgentResult {
  const strings = getUiStrings(locale);

  if (payload.confirmed && payload.plan) {
    return {
      route: 'plan',
      requiresConfirmation: false,
      status: 'completed',
      message: strings.browserPreviewPlanExecutedMessage,
      journal: {
        hasFailures: false,
        errorMessage: '',
        steps: payload.plan.steps.map((step) => ({
          type: step.type,
          title: formatBrowserPreviewPlanStep(step, locale),
          status: 'completed',
          message: strings.browserPreviewPlanExecutedMessage,
          errorMessage: '',
        })),
      },
    };
  }

  if (/\bsummary\b|\bworksheet\b|\bsheet\b/i.test(payload.userInput)) {
    return {
      route: 'plan',
      requiresConfirmation: true,
      status: 'preview',
      message: strings.browserPreviewPlanPreparedMessage,
      planner: {
        mode: 'plan',
        assistantMessage: strings.browserPreviewPlanPreparedMessage,
        plan: createBrowserPreviewPlan(locale),
      },
    };
  }

  return {
    route: 'chat',
    requiresConfirmation: false,
    status: 'completed',
    message: strings.browserPreviewChatFallback,
  };
}

function createBrowserPreviewPlan(locale: HostContext['resolvedUiLocale']): AgentPlan {
  const strings = getUiStrings(locale);

  return {
    summary: strings.browserPreviewPlanSummary,
    steps: [
      {
        type: 'excel.addWorksheet',
        args: {
          newSheetName: 'Summary',
        },
      },
      {
        type: 'excel.writeRange',
        args: {
          targetAddress: 'Summary!A1:B3',
          values: [
            ['Name', 'Region'],
            ['Project A', 'CN'],
            ['Project B', 'US'],
          ],
        },
      },
    ],
  };
}

function formatBrowserPreviewPlanStep(step: { type: string; args?: Record<string, unknown> }, locale: HostContext['resolvedUiLocale']) {
  const strings = getUiStrings(locale);

  switch (step.type) {
    case 'excel.addWorksheet':
      return strings.formatPlanStepAddWorksheet(String(step.args?.newSheetName ?? '').trim());
    case 'excel.writeRange':
      return strings.formatPlanStepWriteRange(String(step.args?.targetAddress ?? '').trim());
    case 'excel.renameWorksheet':
      return strings.formatPlanStepRenameWorksheet(
        String(step.args?.sheetName ?? '').trim(),
        String(step.args?.newSheetName ?? '').trim(),
      );
    case 'excel.deleteWorksheet':
      return strings.formatPlanStepDeleteWorksheet(String(step.args?.sheetName ?? '').trim());
    case 'skill.upload_data':
      return strings.formatPlanStepUploadData;
    default:
      return step.type;
  }
}

function buildBrowserPreviewUpload(userInput: string): UploadPreview {
  const projectName = extractBrowserPreviewProjectName(userInput);
  const rows = [
    ['Project A', 'CN'],
    ['Project B', 'US'],
  ];

  return {
    projectName,
    sheetName: 'Sheet1',
    address: 'A1:C3',
    headers: ['Name', 'Region'],
    rows,
    records: rows.map((row) => ({
      Name: row[0],
      Region: row[1],
    })),
  };
}

function extractBrowserPreviewProjectName(userInput: string): string {
  return extractResolvedBrowserPreviewProjectName(userInput);
}

function isUploadDataSkillInput(userInput: string): boolean {
  return matchesUploadDataSkillInput(userInput);
}

