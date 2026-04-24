import { describe, expect, it, vi } from 'vitest';
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
  const uploadToProjectA = '\u628A\u9009\u4E2D\u6570\u636E\u4E0A\u4F20\u5230\u9879\u76EEA';
  const uploadWithoutTargetSeparator = '\u628A\u9009\u4E2D\u6570\u636E\u4E0A\u4F20\u9879\u76EEA';
  const projectA = '\u9879\u76EEA';

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
      businessBaseUrl: '',
      model: 'gpt-5-mini',
      ssoUrl: '',
      ssoLoginSuccessPath: '',
      uiLanguageOverride: 'system',
    });

    await expect(bridge.getHostContext()).resolves.toEqual({
      resolvedUiLocale: 'en',
      uiLanguageOverride: 'system',
    });

    await expect(bridge.saveSettings({
      apiKey: 'key-123',
      baseUrl: 'https://api.example.com',
      businessBaseUrl: 'https://business.example.com',
      model: 'gpt-5-mini',
      ssoUrl: 'https://sso.example.com',
      ssoLoginSuccessPath: '/auth/success',
      uiLanguageOverride: 'zh',
    })).resolves.toEqual({
      apiKey: 'key-123',
      baseUrl: 'https://api.example.com',
      businessBaseUrl: 'https://business.example.com',
      model: 'gpt-5-mini',
      ssoUrl: 'https://sso.example.com',
      ssoLoginSuccessPath: '/auth/success',
      uiLanguageOverride: 'zh',
    });

    await expect(bridge.getSettings()).resolves.toEqual({
      apiKey: 'key-123',
      baseUrl: 'https://api.example.com',
      businessBaseUrl: 'https://business.example.com',
      model: 'gpt-5-mini',
      ssoUrl: 'https://sso.example.com',
      ssoLoginSuccessPath: '/auth/success',
      uiLanguageOverride: 'zh',
    });

    await expect(bridge.getHostContext()).resolves.toEqual({
      resolvedUiLocale: 'en',
      uiLanguageOverride: 'zh',
    });

    await expect(bridge.getSelectionContext()).resolves.toEqual({
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

  it('returns write-command previews in browser preview mode', async () => {
    const bridge = new NativeBridge(undefined);

    await expect(bridge.executeExcelCommand({
      commandType: 'excel.renameWorksheet',
      sheetName: 'Sheet1',
      newSheetName: 'Summary',
      confirmed: false,
    })).resolves.toEqual({
      commandType: 'excel.renameWorksheet',
      requiresConfirmation: true,
      status: 'preview',
      message: 'Confirm this Excel action before the workbook is modified.',
      preview: {
        title: 'Rename worksheet',
        summary: 'Rename worksheet "Sheet1" to "Summary"',
        details: ['Workbook: Browser Preview.xlsx'],
      },
      selectionContext: {
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
      },
    });
  });

  it('rejects invalid write-range payloads in browser preview mode', async () => {
    const bridge = new NativeBridge(undefined);

    await expect(bridge.executeExcelCommand({
      commandType: 'excel.writeRange',
      targetAddress: 'Sheet2!A1:B2',
      sheetName: 'Sheet1',
      values: [['Name', 'Region']],
      confirmed: false,
    })).rejects.toMatchObject({
      code: 'invalid_command',
    });
  });

  it('rejects sheet-qualified write targets without a cell reference in browser preview mode', async () => {
    const bridge = new NativeBridge(undefined);

    await expect(bridge.executeExcelCommand({
      commandType: 'excel.writeRange',
      targetAddress: 'Sheet1!',
      values: [['Name', 'Region']],
      confirmed: false,
    })).rejects.toMatchObject({
      code: 'invalid_command',
    });
  });

  it('routes real Chinese upload input to the upload_data preview in browser preview mode', async () => {
    const bridge = new NativeBridge(undefined);

    await expect(bridge.runSkill({
      userInput: uploadToProjectA,
      confirmed: false,
    })).resolves.toEqual({
      route: 'skill',
      skillName: 'upload_data',
      requiresConfirmation: true,
      status: 'preview',
      message: `Review the upload payload before sending it to ${projectA}.`,
      preview: {
        title: 'Upload selected data',
        summary: `Upload 2 row(s) to ${projectA}`,
        details: [
          'Source: Sheet1!A1:C3',
          'Fields: Name, Region',
        ],
      },
      uploadPreview: {
        projectName: projectA,
        sheetName: 'Sheet1',
        address: 'A1:C3',
        headers: ['Name', 'Region'],
        rows: [
          ['Project A', 'CN'],
          ['Project B', 'US'],
        ],
        records: [
          { Name: 'Project A', Region: 'CN' },
          { Name: 'Project B', Region: 'US' },
        ],
      },
    });
  });

  it('returns the chat fallback for English upload text without a target project in browser preview mode', async () => {
    const bridge = new NativeBridge(undefined);

    await expect(bridge.runSkill({
      userInput: 'upload project data',
      confirmed: false,
    })).resolves.toEqual({
      route: 'chat',
      requiresConfirmation: false,
      status: 'completed',
      message: 'General chat routing is not implemented yet. Use /upload_data ... or a direct Excel command.',
    });
  });

  it('returns the chat fallback for Chinese upload text without an explicit target separator in browser preview mode', async () => {
    const bridge = new NativeBridge(undefined);

    await expect(bridge.runSkill({
      userInput: uploadWithoutTargetSeparator,
      confirmed: false,
    })).resolves.toEqual({
      route: 'chat',
      requiresConfirmation: false,
      status: 'completed',
      message: 'General chat routing is not implemented yet. Use /upload_data ... or a direct Excel command.',
    });
  });

  it('rejects malformed confirmed upload previews in browser preview mode', async () => {
    const bridge = new NativeBridge(undefined);

    await expect(bridge.runSkill({
      userInput: uploadToProjectA,
      confirmed: true,
      uploadPreview: {
        projectName: projectA,
        sheetName: 'Sheet1',
        address: 'A1:C3',
        headers: null as unknown as string[],
        rows: null as unknown as string[][],
        records: null as unknown as Array<Record<string, string>>,
      },
    })).rejects.toMatchObject({
      code: 'invalid_command',
      message: 'upload_data confirmation requires a complete preview payload.',
    });
  });

  it('returns the chat fallback for unknown skill input in browser preview mode', async () => {
    const bridge = new NativeBridge(undefined);

    await expect(bridge.runSkill({
      userInput: '帮我总结一下这个工作簿',
      confirmed: false,
    })).resolves.toEqual({
      route: 'chat',
      requiresConfirmation: false,
      status: 'completed',
      message: 'General chat routing is not implemented yet. Use /upload_data ... or a direct Excel command.',
    });
  });

  it('returns a plan preview in browser preview mode for natural-language planner input', async () => {
    const bridge = new NativeBridge(undefined);

    await expect(bridge.runAgent({
      userInput: 'Create a summary sheet from the current selection',
      confirmed: false,
    })).resolves.toEqual({
      route: 'plan',
      requiresConfirmation: true,
      status: 'preview',
      message: 'I prepared a plan. Review it before Excel is changed.',
      planner: {
        mode: 'plan',
        assistantMessage: 'I prepared a plan. Review it before Excel is changed.',
        plan: {
          summary: 'Create a Summary sheet and write the selected rows.',
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
        },
      },
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
        businessBaseUrl: 'https://business.example.com',
        model: 'gpt-5-mini',
        ssoUrl: '',
        ssoLoginSuccessPath: '',
        uiLanguageOverride: 'system',
      },
    });

    await expect(pending).resolves.toEqual({
      apiKey: '',
      baseUrl: 'https://api.example.com',
      businessBaseUrl: 'https://business.example.com',
      model: 'gpt-5-mini',
      ssoUrl: '',
      ssoLoginSuccessPath: '',
      uiLanguageOverride: 'system',
    });
  });

  it('sends getHostContext requests through the structured bridge contract', async () => {
    const webView = createMockWebView();
    const bridge = new NativeBridge(webView);

    const pending = bridge.getHostContext();
    const [request] = webView.postedMessages as Array<{ type: string; requestId: string }>;

    expect(request.type).toBe('bridge.getHostContext');

    webView.dispatch({
      type: 'bridge.getHostContext',
      requestId: request.requestId,
      ok: true,
      payload: {
        resolvedUiLocale: 'zh',
        uiLanguageOverride: 'en',
      },
    });

    await expect(pending).resolves.toEqual({
      resolvedUiLocale: 'zh',
      uiLanguageOverride: 'en',
    });
  });

  it('sends executeExcelCommand requests through the structured bridge contract', async () => {
    const webView = createMockWebView();
    const bridge = new NativeBridge(webView);

    const pending = bridge.executeExcelCommand({
      commandType: 'excel.readSelectionTable',
      confirmed: false,
    });
    const [request] = webView.postedMessages as Array<{ type: string; requestId: string }>;

    expect(request.type).toBe('bridge.executeExcelCommand');

    webView.dispatch({
      type: 'bridge.executeExcelCommand',
      requestId: request.requestId,
      ok: true,
      payload: {
        commandType: 'excel.readSelectionTable',
        requiresConfirmation: false,
        status: 'completed',
        message: 'Read selection from Sheet1 A1:C4.',
      },
    });

    await expect(pending).resolves.toEqual({
      commandType: 'excel.readSelectionTable',
      requiresConfirmation: false,
      status: 'completed',
      message: 'Read selection from Sheet1 A1:C4.',
    });
  });

  it('sends runAgent requests through the structured bridge contract', async () => {
    const webView = createMockWebView();
    const bridge = new NativeBridge(webView);

    const pending = bridge.runAgent({
      userInput: 'Create a summary sheet from the current selection',
      confirmed: false,
    });
    const [request] = webView.postedMessages as Array<{ type: string; requestId: string }>;

    expect(request.type).toBe('bridge.runAgent');

    webView.dispatch({
      type: 'bridge.runAgent',
      requestId: request.requestId,
      ok: true,
      payload: {
        route: 'plan',
        requiresConfirmation: true,
        status: 'preview',
        message: 'I prepared a plan. Review it before Excel is changed.',
        planner: {
          mode: 'plan',
          assistantMessage: 'I prepared a plan. Review it before Excel is changed.',
          plan: {
            summary: 'Create a Summary sheet and write the selected rows.',
            steps: [],
          },
        },
      },
    });

    await expect(pending).resolves.toEqual({
      route: 'plan',
      requiresConfirmation: true,
      status: 'preview',
      message: 'I prepared a plan. Review it before Excel is changed.',
      planner: {
        mode: 'plan',
        assistantMessage: 'I prepared a plan. Review it before Excel is changed.',
        plan: {
          summary: 'Create a Summary sheet and write the selected rows.',
          steps: [],
        },
      },
    });
  });

  it('notifies selection context subscribers when native events arrive', async () => {
    const webView = createMockWebView();
    const bridge = new NativeBridge(webView);
    const listener = vi.fn();

    const unsubscribe = bridge.onSelectionContextChanged(listener);

    webView.dispatch({
      type: 'bridge.selectionContextChanged',
      payload: {
        hasSelection: true,
        workbookName: 'Quarterly Report.xlsx',
        sheetName: 'Sheet2',
        address: 'B2:D5',
        rowCount: 4,
        columnCount: 3,
        isContiguous: false,
        headerPreview: [],
        sampleRows: [],
        warningMessage: 'Multiple selection areas are not supported yet.',
      },
    });

    expect(listener).toHaveBeenCalledWith({
      hasSelection: true,
      workbookName: 'Quarterly Report.xlsx',
      sheetName: 'Sheet2',
      address: 'B2:D5',
      rowCount: 4,
      columnCount: 3,
      isContiguous: false,
      headerPreview: [],
      sampleRows: [],
      warningMessage: 'Multiple selection areas are not supported yet.',
    });

    unsubscribe();
    listener.mockClear();

    webView.dispatch({
      type: 'bridge.selectionContextChanged',
      payload: {
        hasSelection: true,
        workbookName: 'Quarterly Report.xlsx',
        sheetName: 'Sheet3',
        address: 'C3:E6',
        rowCount: 4,
        columnCount: 3,
        isContiguous: true,
        headerPreview: ['Name'],
        sampleRows: [],
        warningMessage: null,
      },
    });

    expect(listener).not.toHaveBeenCalled();
  });
});
