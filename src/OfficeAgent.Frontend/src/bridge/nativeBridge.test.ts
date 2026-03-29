import { describe, expect, it } from 'vitest';
import { NativeBridge } from './nativeBridge';

type WebMessageEventHandler = (event: { data: unknown }) => void;

interface MockWebView {
  postedMessages: unknown[];
  addEventListener: (type: 'message', listener: WebMessageEventHandler) => void;
  removeEventListener: (type: 'message', listener: WebMessageEventHandler) => void;
  postMessage: (message: unknown) => void;
  dispatch: (message: unknown) => void;
}

function createMockWebView(): MockWebView {
  const listeners = new Set<WebMessageEventHandler>();
  const postedMessages: unknown[] = [];

  return {
    postedMessages,
    addEventListener(_type, listener) {
      listeners.add(listener);
    },
    removeEventListener(_type, listener) {
      listeners.delete(listener);
    },
    postMessage(message) {
      postedMessages.push(message);
    },
    dispatch(message) {
      listeners.forEach((listener) => listener({ data: message }));
    },
  };
}

describe('NativeBridge', () => {
  it('correlates responses by request id', async () => {
    const webView = createMockWebView();
    const bridge = new NativeBridge(webView);

    const pending = bridge.ping();
    const [request] = webView.postedMessages as Array<{ requestId: string }>;
    let settled = false;
    pending.then(() => {
      settled = true;
    });

    webView.dispatch({
      type: 'bridge.ping',
      requestId: 'ignored-request',
      ok: true,
      payload: { host: 'ignored-host', version: '0.0.0' },
    });

    await Promise.resolve();
    expect(settled).toBe(false);

    webView.dispatch({
      type: 'bridge.ping',
      requestId: request.requestId,
      ok: true,
      payload: { host: 'native-host', version: '1.0.0' },
    });

    await expect(pending).resolves.toEqual({
      host: 'native-host',
      version: '1.0.0',
    });
  });

  it('falls back to browser preview mode when WebView2 is unavailable', async () => {
    const bridge = new NativeBridge(undefined);

    await expect(bridge.ping()).resolves.toEqual({
      host: 'browser-preview',
      version: 'dev',
    });
  });

  it('returns default settings and session state in browser preview mode', async () => {
    const bridge = new NativeBridge(undefined);

    await expect(bridge.getSettings()).resolves.toEqual({
      apiKey: '',
      baseUrl: 'https://api.example.com',
      model: 'gpt-5-mini',
    });

    await expect(bridge.getSessions()).resolves.toEqual({
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
    });
  });

  it('sends getSettings requests through the structured bridge contract', async () => {
    const webView = createMockWebView();
    const bridge = new NativeBridge(webView);

    const pending = bridge.getSettings();
    const [request] = webView.postedMessages as Array<{ type: string; requestId: string }>;

    expect(request.type).toBe('bridge.getSettings');

    webView.dispatch({
      type: 'bridge.getSettings',
      requestId: request.requestId,
      ok: true,
      payload: {
        apiKey: '',
        baseUrl: 'https://api.example.com',
        model: 'gpt-5-mini',
      },
    });

    await expect(pending).resolves.toEqual({
      apiKey: '',
      baseUrl: 'https://api.example.com',
      model: 'gpt-5-mini',
    });
  });
});
