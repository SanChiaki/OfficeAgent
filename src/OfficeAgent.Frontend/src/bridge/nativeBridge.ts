import type {
  AppSettings,
  BridgeErrorPayload,
  BridgeEventEnvelope,
  BridgeRequestEnvelope,
  BridgeResponseEnvelope,
  SelectionContext,
  SessionState,
  PingPayload,
  WebViewHostLike,
  WebViewMessageEventLike,
} from '../types/bridge';

const BRIDGE_TYPES = {
  ping: 'bridge.ping',
  getSettings: 'bridge.getSettings',
  getSelectionContext: 'bridge.getSelectionContext',
  selectionContextChanged: 'bridge.selectionContextChanged',
  getSessions: 'bridge.getSessions',
  saveSettings: 'bridge.saveSettings',
  executeExcelCommand: 'bridge.executeExcelCommand',
  runSkill: 'bridge.runSkill',
} as const;

const BROWSER_PREVIEW_PING: PingPayload = {
  host: 'browser-preview',
  version: 'dev',
};

const BROWSER_PREVIEW_SETTINGS: AppSettings = {
  apiKey: '',
  baseUrl: 'https://api.example.com',
  model: 'gpt-5-mini',
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

  getSelectionContext() {
    return this.invoke<void, SelectionContext>(BRIDGE_TYPES.getSelectionContext);
  }

  getSessions() {
    return this.invoke<void, SessionState>(BRIDGE_TYPES.getSessions);
  }

  saveSettings(payload: AppSettings) {
    return this.invoke<AppSettings, AppSettings>(BRIDGE_TYPES.saveSettings, payload);
  }

  executeExcelCommand(payload: unknown) {
    return this.invoke(BRIDGE_TYPES.executeExcelCommand, payload);
  }

  runSkill(payload: unknown) {
    return this.invoke(BRIDGE_TYPES.runSkill, payload);
  }

  onSelectionContextChanged(listener: SelectionContextListener) {
    this.selectionContextListeners.add(listener);
    return () => {
      this.selectionContextListeners.delete(listener);
    };
  }

  private invoke<TPayload, TResult>(type: string, payload?: TPayload): Promise<TResult> {
    if (!this.webView) {
      if (type === BRIDGE_TYPES.ping) {
        return Promise.resolve(BROWSER_PREVIEW_PING as TResult);
      }

      if (type === BRIDGE_TYPES.getSettings) {
        return Promise.resolve(BROWSER_PREVIEW_SETTINGS as TResult);
      }

      if (type === BRIDGE_TYPES.getSelectionContext) {
        return Promise.resolve(BROWSER_PREVIEW_SELECTION_CONTEXT as TResult);
      }

      if (type === BRIDGE_TYPES.getSessions) {
        return Promise.resolve(BROWSER_PREVIEW_SESSIONS as TResult);
      }

      if (type === BRIDGE_TYPES.saveSettings) {
        return Promise.resolve({
          apiKey: typeof (payload as AppSettings | undefined)?.apiKey === 'string' ? (payload as AppSettings).apiKey : '',
          baseUrl: typeof (payload as AppSettings | undefined)?.baseUrl === 'string'
            ? (payload as AppSettings).baseUrl
            : BROWSER_PREVIEW_SETTINGS.baseUrl,
          model: typeof (payload as AppSettings | undefined)?.model === 'string'
            ? (payload as AppSettings).model
            : BROWSER_PREVIEW_SETTINGS.model,
        } as TResult);
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
