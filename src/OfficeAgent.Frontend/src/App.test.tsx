import { act, cleanup, render, screen, within } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import App from './App';
import { nativeBridge } from './bridge/nativeBridge';
import type { SelectionContext } from './types/bridge';

vi.mock('./bridge/nativeBridge', () => ({
  nativeBridge: {
    ping: vi.fn(),
    getHostContext: vi.fn(),
    getSelectionContext: vi.fn(),
    getSessions: vi.fn(),
    saveSessions: vi.fn(),
    onSelectionContextChanged: vi.fn(),
    getSettings: vi.fn(),
    saveSettings: vi.fn(),
    executeExcelCommand: vi.fn(),
    runSkill: vi.fn(),
    runAgent: vi.fn(),
    login: vi.fn(),
    logout: vi.fn(),
    getLoginStatus: vi.fn(),
  },
}));

const mockedBridge = vi.mocked(nativeBridge);
let selectionContextListener: ((context: SelectionContext) => void) | null = null;
const originalScrollTo = HTMLElement.prototype.scrollTo;
let scrollToSpy: ReturnType<typeof vi.fn>;

function createDeferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (reason?: unknown) => void;

  const promise = new Promise<T>((resolvePromise, rejectPromise) => {
    resolve = resolvePromise;
    reject = rejectPromise;
  });

  return { promise, resolve, reject };
}

beforeEach(() => {
  scrollToSpy = vi.fn();
  Object.defineProperty(HTMLElement.prototype, 'scrollTo', {
    configurable: true,
    writable: true,
    value: scrollToSpy,
  });
  mockedBridge.ping.mockResolvedValue({
    host: 'browser-preview',
    version: 'dev',
  });
  mockedBridge.getHostContext.mockResolvedValue({
    resolvedUiLocale: 'zh',
    uiLanguageOverride: 'system',
  });
  mockedBridge.getSelectionContext.mockResolvedValue({
    hasSelection: true,
    workbookName: 'Quarterly Report.xlsx',
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
  mockedBridge.getSessions.mockResolvedValue({
    activeSessionId: 'browser-preview-session',
    sessions: [
      {
        id: 'browser-preview-session',
        title: 'Browser preview',
        createdAtUtc: '2026-03-29T00:00:00.0000000Z',
        updatedAtUtc: '2026-03-29T00:00:00.0000000Z',
        messages: [],
      },
      {
        id: 'review-session',
        title: 'Review notes',
        createdAtUtc: '2026-03-29T01:00:00.0000000Z',
        updatedAtUtc: '2026-03-29T01:00:00.0000000Z',
        messages: [],
      },
    ],
  });
  mockedBridge.getSettings.mockResolvedValue({
    apiKey: '',
    baseUrl: 'https://api.example.com',
    businessBaseUrl: 'https://business.example.com',
    model: 'gpt-5-mini',
    ssoUrl: '',
    ssoLoginSuccessPath: '',
    uiLanguageOverride: 'system',
  });
  mockedBridge.getLoginStatus.mockResolvedValue({
    isLoggedIn: false,
    ssoUrl: '',
  });
  mockedBridge.onSelectionContextChanged.mockImplementation((listener) => {
    selectionContextListener = listener;
    return () => {
      selectionContextListener = null;
    };
  });
  mockedBridge.saveSessions.mockImplementation(async (state) => state);
  mockedBridge.saveSettings.mockImplementation(async (settings) => settings);
  mockedBridge.executeExcelCommand.mockResolvedValue({
    commandType: 'excel.readSelectionTable',
    requiresConfirmation: false,
    status: 'completed',
    message: 'Read selection from Sheet1 A1:C4.',
    table: {
      sheetName: 'Sheet1',
      address: 'A1:C4',
      headers: ['Name', 'Region', 'Amount'],
      rows: [
        ['Project A', 'CN', '42'],
        ['Project B', 'US', '36'],
      ],
    },
  });
  mockedBridge.runSkill.mockResolvedValue({
    route: 'chat',
    requiresConfirmation: false,
    status: 'completed',
    message: 'General chat routing is not implemented yet. Use /upload_data ... or a direct Excel command.',
  });
  mockedBridge.runAgent.mockResolvedValue({
    route: 'chat',
    requiresConfirmation: false,
    status: 'completed',
    message: 'General chat routing is not implemented yet. Use /upload_data ... or a direct Excel command.',
  });
});

afterEach(() => {
  cleanup();
  vi.clearAllMocks();
  vi.useRealTimers();
  selectionContextListener = null;
  if (originalScrollTo) {
    Object.defineProperty(HTMLElement.prototype, 'scrollTo', {
      configurable: true,
      writable: true,
      value: originalScrollTo,
    });
    return;
  }

  Reflect.deleteProperty(HTMLElement.prototype, 'scrollTo');
});

describe('App shell', () => {
  it('loads host context and keeps the Chinese fixed UI path stable', async () => {
    render(<App />);

    expect(mockedBridge.getHostContext).toHaveBeenCalledTimes(1);
    expect(await screen.findByText(/欢迎使用\s*ISDP/)).toBeInTheDocument();
    expect(screen.getByRole('button', { name: /打开设置/i })).toBeInTheDocument();
    expect(screen.getByRole('button', { name: /发送/i })).toBeInTheDocument();
    expect(document.documentElement.lang).toBe('zh');
  });

  it('falls back to English fixed UI when getHostContext fails', async () => {
    mockedBridge.getHostContext.mockRejectedValueOnce(new Error('host context unavailable'));

    render(<App />);

    expect(await screen.findByText(/welcome to isdp/i)).toBeInTheDocument();
    expect(screen.getByRole('button', { name: /open settings/i })).toBeInTheDocument();
    expect(screen.getByRole('button', { name: /send/i })).toBeInTheDocument();
    expect(document.documentElement.lang).toBe('en');
  });

  it('renders the expected task pane regions', async () => {
    const user = userEvent.setup();

    render(<App />);

    expect(await screen.findByRole('banner', { name: /聊天页眉/i })).toBeInTheDocument();
    expect(
      screen.getByRole('region', { name: /消息线程/i }),
    ).toBeInTheDocument();
    expect(
      screen.getByText(/欢迎使用\s*ISDP/),
    ).toBeInTheDocument();
    expect(
      screen.getByRole('status', { name: /选区胶囊/i }),
    ).toBeInTheDocument();
    expect(
      screen.getByRole('textbox', { name: /消息输入框/i }),
    ).toBeInTheDocument();
    expect(
      screen.getByRole('button', { name: /打开会话列表/i }),
    ).toBeInTheDocument();
    expect(
      screen.getByRole('button', { name: /打开设置/i }),
    ).toBeInTheDocument();
    expect(
      await screen.findByText(/未命名会话/i, { selector: 'h1' }),
    ).toBeInTheDocument();
    expect(
      await screen.findByText(/sheet1 · a1:c4/i),
    ).toBeInTheDocument();
    expect(
      screen.queryByText(/quarterly report\.xlsx/i),
    ).not.toBeInTheDocument();
    expect(
      screen.queryByText(/4 rows x 3 columns/i),
    ).not.toBeInTheDocument();
    expect(
      screen.queryByText(/headers: name, region, amount/i),
    ).not.toBeInTheDocument();
    expect(
      screen.queryByText(/project a · cn · 42/i),
    ).not.toBeInTheDocument();
    expect(
      screen.queryByRole('complementary', { name: /会话抽屉/i }),
    ).not.toBeInTheDocument();

    await user.click(screen.getByRole('button', { name: /打开会话列表/i }));

    expect(
      await screen.findByRole('complementary', { name: /会话抽屉/i }),
    ).toBeInTheDocument();
    expect(
      await screen.findByRole('button', { name: /browser preview/i }),
    ).toBeInTheDocument();

    await user.click(screen.getByRole('button', { name: /打开设置/i }));

    const settingsDialog = screen.getByRole('dialog', { name: /设置对话框/i });
    expect(settingsDialog).toBeInTheDocument();
    expect(
      within(settingsDialog).getAllByLabelText(/api 密钥/i)[0],
    ).toHaveValue('');
    expect(
      within(settingsDialog).getByRole('textbox', { name: /^基础 url$/i }),
    ).toHaveValue('https://api.example.com');
    expect(
      within(settingsDialog).getByRole('textbox', { name: /^业务基础 url$/i }),
    ).toHaveValue('https://business.example.com');
    expect(
      within(settingsDialog).getByText(/已连接 browser-preview \(dev\)/i),
    ).toBeInTheDocument();
    expect(
      within(settingsDialog).getByRole('textbox', { name: /^模型$/i }),
    ).toHaveValue('gpt-5-mini');
    expect(
      within(settingsDialog).getByRole('textbox', { name: /^sso 地址$/i }),
    ).toHaveValue('');
  });

  it('renders English fixed UI when the host locale resolves to English', async () => {
    const user = userEvent.setup();
    mockedBridge.getHostContext.mockResolvedValueOnce({
      resolvedUiLocale: 'en',
      uiLanguageOverride: 'system',
    });

    render(<App />);

    expect(await screen.findByText(/welcome to isdp/i)).toBeInTheDocument();
    expect(screen.getByRole('button', { name: /open settings/i })).toBeInTheDocument();
    expect(screen.getByRole('button', { name: /send/i })).toBeInTheDocument();
    expect(await screen.findByText(/untitled/i, { selector: 'h1' })).toBeInTheDocument();
    expect(await screen.findByText(/^sheet1 · a1:c4$/i)).toBeInTheDocument();
    expect(document.documentElement.lang).toBe('en');

    await user.click(screen.getByRole('button', { name: /open settings/i }));

    const settingsDialog = screen.getByRole('dialog', { name: /settings dialog/i });
    expect(settingsDialog).toBeInTheDocument();
    expect(screen.getByRole('button', { name: /^save$/i })).toBeInTheDocument();
    expect(screen.getByRole('button', { name: /^cancel$/i })).toBeInTheDocument();
    expect(within(settingsDialog).getAllByLabelText(/^(api key|api 密钥)$/i)[0]).toHaveValue('');
    expect(within(settingsDialog).getByRole('textbox', { name: /^(base url|基础 url)$/i })).toHaveValue('https://api.example.com');
    expect(within(settingsDialog).getByRole('textbox', { name: /^(business base url|业务基础 url)$/i })).toHaveValue('https://business.example.com');
    expect(within(settingsDialog).getByRole('textbox', { name: /^(model|模型)$/i })).toHaveValue('gpt-5-mini');
    expect(within(settingsDialog).getByRole('textbox', { name: /^(sso url|sso 地址)$/i })).toHaveValue('');
  });

  it('keeps preview messages and confirmation cards in English when the host locale resolves to English', async () => {
    const user = userEvent.setup();
    mockedBridge.getHostContext.mockResolvedValueOnce({
      resolvedUiLocale: 'en',
      uiLanguageOverride: 'system',
    });
    mockedBridge.runSkill.mockResolvedValueOnce({
      route: 'skill',
      skillName: 'upload_data',
      requiresConfirmation: true,
      status: 'preview',
      message: 'Review the upload payload before sending it to Project A.',
      preview: {
        title: 'Upload selected data',
        summary: 'Upload 2 row(s) to Project A',
        details: ['Source: Sheet1!A1:C3', 'Fields: Name, Region'],
      },
      uploadPreview: {
        projectName: 'Project A',
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

    render(<App />);

    await user.type(screen.getByRole('textbox', { name: /message composer/i }), 'upload selected data to Project A');
    await user.click(screen.getByRole('button', { name: /send/i }));

    const confirmationCard = await screen.findByRole('article', { name: /confirm excel action/i });
    expect(await screen.findByText(/review the upload payload before sending it to project a/i)).toBeInTheDocument();
    expect(confirmationCard).toBeInTheDocument();
    expect(within(confirmationCard).getByText(/^upload selected data$/i)).toBeInTheDocument();
    expect(within(confirmationCard).getByText(/upload 2 row\(s\) to project a/i)).toBeInTheDocument();
    expect(within(confirmationCard).getByText(/source: sheet1!a1:c3/i)).toBeInTheDocument();
    expect(within(confirmationCard).getByText(/fields: name, region/i)).toBeInTheDocument();
  });

  it('keeps an explicitly chosen legacy placeholder title after reload when persisted ownership is false', async () => {
    const user = userEvent.setup();
    mockedBridge.getHostContext.mockResolvedValueOnce({
      resolvedUiLocale: 'en',
      uiLanguageOverride: 'system',
    });
    mockedBridge.getSessions.mockResolvedValueOnce({
      activeSessionId: 'browser-preview-session',
      sessions: [
        {
          id: 'browser-preview-session',
          title: 'Browser preview',
          isSystemUntitled: false,
          createdAtUtc: '2026-03-29T00:00:00.0000000Z',
          updatedAtUtc: '2026-03-29T00:00:00.0000000Z',
          messages: [],
        } as never,
        {
          id: 'user-new-chat-session',
          title: 'New chat',
          isSystemUntitled: false,
          createdAtUtc: '2026-03-29T00:00:00.0000000Z',
          updatedAtUtc: '2026-03-29T00:00:00.0000000Z',
          messages: [],
        } as never,
      ],
    });

    render(<App />);

    await user.click(await screen.findByRole('button', { name: /open sessions drawer/i }));
    const sidebar = await screen.findByRole('complementary', { name: /sessions drawer/i });
    await user.click(within(sidebar).getByRole('button', { name: /^new chat$/i }));

    expect(await screen.findByRole('heading', { name: /^new chat$/i })).toBeInTheDocument();
    expect(screen.queryByRole('heading', { name: /^untitled$/i })).not.toBeInTheDocument();
  });

  it('does not persist a localized placeholder when renaming an untitled session without editing it', async () => {
    const user = userEvent.setup();

    render(<App />);

    await user.click(await screen.findByRole('button', { name: /打开会话列表/i }));
    const sidebar = await screen.findByRole('complementary', { name: /会话抽屉/i });
    await user.click(within(sidebar).getAllByRole('button', { name: /重命名会话/i })[0]);
    await user.click(screen.getByRole('button', { name: /确认重命名/i }));
    await user.click(within(sidebar).getByRole('button', { name: /review notes/i }));

    expect(mockedBridge.saveSessions).toHaveBeenCalledWith(expect.objectContaining({
      sessions: expect.arrayContaining([
        expect.objectContaining({
          id: expect.any(String),
          title: 'New chat',
        }),
      ]),
    }));
  });

  it('keeps an explicitly chosen placeholder title as user data', async () => {
    const user = userEvent.setup();

    render(<App />);

    await user.click(await screen.findByRole('button', { name: /打开会话列表/i }));
    const sidebar = await screen.findByRole('complementary', { name: /会话抽屉/i });
    await user.click(within(sidebar).getAllByRole('button', { name: /重命名会话/i })[0]);
    const renameInput = within(sidebar).getByDisplayValue('未命名会话');
    await user.clear(renameInput);
    await user.type(renameInput, '未命名会话');
    await user.click(screen.getByRole('button', { name: /确认重命名/i }));
    await user.click(within(sidebar).getByRole('button', { name: /关闭会话列表/i }));

    expect(await screen.findByText(/^未命名会话$/)).toBeInTheDocument();

    await user.click(screen.getByRole('textbox', { name: /消息输入框/i }));
    await user.type(screen.getByRole('textbox', { name: /消息输入框/i }), 'hello{enter}');

    expect(await screen.findByRole('heading', { name: /未命名会话/i })).toBeInTheDocument();
  });

  it('uses xISDP AI as the fallback panel title before sessions load', async () => {
    const sessionsDeferred = createDeferred<{
      activeSessionId: string;
      sessions: Array<{
        id: string;
        title: string;
        createdAtUtc: string;
        updatedAtUtc: string;
        messages: never[];
      }>;
    }>();
    mockedBridge.getSessions.mockReturnValueOnce(sessionsDeferred.promise);

    render(<App />);

    expect(await screen.findByText(/xISDP AI/i, { selector: 'h1' })).toBeInTheDocument();

    sessionsDeferred.resolve({
      activeSessionId: 'loaded-session',
      sessions: [
        {
          id: 'loaded-session',
          title: 'Loaded session',
          createdAtUtc: '2026-03-29T00:00:00.0000000Z',
          updatedAtUtc: '2026-03-29T00:00:00.0000000Z',
          messages: [],
        },
      ],
    });
  });

  it('saves business base url independently from the llm base url', async () => {
    const user = userEvent.setup();

    render(<App />);

    await user.click(screen.getByRole('button', { name: /打开设置|open settings/i }));
    await user.clear(screen.getByRole('textbox', { name: /^(business base url|业务基础 url)$/i }));
    await user.type(screen.getByRole('textbox', { name: /^(business base url|业务基础 url)$/i }), 'http://localhost:3200');
    await user.click(screen.getByRole('button', { name: /保存|save/i }));

    expect(mockedBridge.saveSettings).toHaveBeenCalledWith(expect.objectContaining({
      baseUrl: 'https://api.example.com',
      businessBaseUrl: 'http://localhost:3200',
    }));
  });

  it('switches the active session when a session chip is clicked', async () => {
    const user = userEvent.setup();

    render(<App />);
    await user.click(await screen.findByRole('button', { name: /打开会话列表|open sessions drawer/i }));
    const sidebar = await screen.findByRole('complementary', { name: /会话抽屉|sessions drawer/i });

    expect(
      await screen.findByRole('heading', { name: /未命名会话|untitled/i }),
    ).toBeInTheDocument();

    await user.click(within(sidebar).getByRole('button', { name: /review notes/i }));

    expect(
      screen.getByRole('heading', { name: /review notes/i }),
    ).toBeInTheDocument();
  });

  it('resets unsaved settings changes when the dialog is cancelled or closed', async () => {
    const user = userEvent.setup();

    render(<App />);

    await user.click(screen.getByRole('button', { name: /打开设置|open settings/i }));
    await user.clear(screen.getByRole('textbox', { name: /^(base url|基础 url)$/i }));
    await user.type(screen.getByRole('textbox', { name: /^(base url|基础 url)$/i }), 'https://changed.example.com');
    await user.click(screen.getByRole('button', { name: /取消|cancel/i }));

    await user.click(screen.getByRole('button', { name: /打开设置|open settings/i }));
    expect(
      screen.getByRole('textbox', { name: /^(base url|基础 url)$/i }),
    ).toHaveValue('https://api.example.com');

    await user.clear(screen.getByRole('textbox', { name: /^(base url|基础 url)$/i }));
    await user.type(screen.getByRole('textbox', { name: /^(base url|基础 url)$/i }), 'https://closed.example.com');
    await user.click(screen.getByRole('button', { name: /关闭|close/i }));

    await user.click(screen.getByRole('button', { name: /打开设置|open settings/i }));
    expect(
      screen.getByRole('textbox', { name: /^(base url|基础 url)$/i }),
    ).toHaveValue('https://api.example.com');
  });

  it('does not block settings save when the follow-up host-context refresh never resolves', async () => {
    const user = userEvent.setup();
    const pendingHostContext = createDeferred<{ resolvedUiLocale: 'zh' | 'en'; uiLanguageOverride: 'system' | 'zh' | 'en' }>();

    mockedBridge.getHostContext
      .mockResolvedValueOnce({
        resolvedUiLocale: 'zh',
        uiLanguageOverride: 'system',
      })
      .mockReturnValueOnce(pendingHostContext.promise);

    render(<App />);

    await user.click(await screen.findByRole('button', { name: /打开设置/i }));
    await user.click(screen.getByRole('button', { name: /保存/i }));

    expect(screen.queryByRole('dialog', { name: /设置对话框/i })).not.toBeInTheDocument();
  });

  it('applies a late host locale after the startup timeout fallback already rendered English UI', async () => {
    vi.useFakeTimers();
    const pendingHostContext = createDeferred<{ resolvedUiLocale: 'zh' | 'en'; uiLanguageOverride: 'system' | 'zh' | 'en' }>();
    mockedBridge.getHostContext.mockReturnValueOnce(pendingHostContext.promise);

    render(<App />);

    await vi.advanceTimersByTimeAsync(1500);

    expect(screen.getByText(/welcome to isdp/i)).toBeInTheDocument();
    expect(document.documentElement.lang).toBe('en');

    await act(async () => {
      pendingHostContext.resolve({
        resolvedUiLocale: 'zh',
        uiLanguageOverride: 'system',
      });
      await Promise.resolve();
    });

    expect(screen.getByText(/欢迎使用\s*ISDP/)).toBeInTheDocument();
    expect(document.documentElement.lang).toBe('zh');
  });

  it('shows an inline error when saving settings fails', async () => {
    const user = userEvent.setup();
    mockedBridge.saveSettings.mockRejectedValueOnce(new Error('Bridge write failed'));

    render(<App />);

    await user.click(screen.getByRole('button', { name: /打开设置|open settings/i }));
    await user.click(screen.getByRole('button', { name: /保存|save/i }));

    expect(
      await screen.findByRole('alert'),
    ).toBeInTheDocument();
    expect(screen.getByRole('alert')).toHaveTextContent(/bridge write failed/i);
    expect(
      screen.getByRole('dialog', { name: /设置对话框|settings dialog/i }),
    ).toBeInTheDocument();
  });

  it('does not overwrite unsaved settings edits when settings load late', async () => {
    const user = userEvent.setup();
    const delayedSettings = createDeferred<{
      apiKey: string;
      baseUrl: string;
      businessBaseUrl: string;
      model: string;
      ssoUrl: string;
      ssoLoginSuccessPath: string;
    }>();
    mockedBridge.getSettings.mockReturnValueOnce(delayedSettings.promise);

    render(<App />);

    await user.click(screen.getByRole('button', { name: /打开设置|open settings/i }));
    const baseUrlInput = screen.getByRole('textbox', { name: /^(base url|基础 url)$/i });
    await user.clear(baseUrlInput);
    await user.type(baseUrlInput, 'https://draft.example.com');

    delayedSettings.resolve({
      apiKey: '',
      baseUrl: 'https://loaded.example.com',
      businessBaseUrl: 'https://business-loaded.example.com',
      model: 'gpt-5-mini',
      ssoUrl: '',
      ssoLoginSuccessPath: '',
    });

    expect(
      await screen.findByDisplayValue('https://draft.example.com'),
    ).toBeInTheDocument();
  });

  it('treats the settings dialog as a modal and restores focus on close', async () => {
    const user = userEvent.setup();

    render(<App />);

    const settingsButton = screen.getByRole('button', { name: /打开设置|open settings/i });
    await user.click(settingsButton);

    const dialog = screen.getByRole('dialog', { name: /设置对话框|settings dialog/i });
    const apiKeyInput = within(dialog).getAllByLabelText(/^(api key|api 密钥)$/i)[0];

    expect(dialog).toHaveAttribute('aria-modal', 'true');
    expect(apiKeyInput).toHaveFocus();

    await user.tab({ shift: true });
    expect(screen.getByRole('button', { name: /关闭|close/i })).toHaveFocus();

    await user.tab({ shift: true });
    expect(screen.getByRole('button', { name: /保存|save/i })).toHaveFocus();

    await user.click(screen.getByRole('button', { name: /关闭|close/i }));
    expect(settingsButton).toHaveFocus();
  });

  it('keeps the settings dialog scrollable within short task pane viewports', async () => {
    const user = userEvent.setup();

    render(<App />);

    await user.click(screen.getByRole('button', { name: /打开设置|open settings/i }));

    const dialog = screen.getByRole('dialog', { name: /设置对话框|settings dialog/i });

    expect(dialog).toHaveClass('settings-dialog');
    expect(dialog).toHaveStyle({
      maxHeight: 'calc(100vh - 32px)',
      overflowY: 'auto',
    });
  });

  it('disables save until settings finish loading', async () => {
    const user = userEvent.setup();
    const delayedSettings = createDeferred<{
      apiKey: string;
      baseUrl: string;
      businessBaseUrl: string;
      model: string;
      ssoUrl: string;
      ssoLoginSuccessPath: string;
    }>();
    mockedBridge.getSettings.mockReturnValueOnce(delayedSettings.promise);

    render(<App />);

    await user.click(screen.getByRole('button', { name: /打开设置|open settings/i }));
    expect(screen.getByRole('button', { name: /保存|save/i })).toBeDisabled();

    delayedSettings.resolve({
      apiKey: 'loaded-key',
      baseUrl: 'https://loaded.example.com',
      businessBaseUrl: 'https://business-loaded.example.com',
      model: 'gpt-5-mini',
      ssoUrl: '',
      ssoLoginSuccessPath: '',
    });

    expect(
      await screen.findByDisplayValue('loaded-key'),
    ).toBeInTheDocument();
    expect(screen.getByRole('button', { name: /保存|save/i })).toBeEnabled();
  });

  it('keeps save disabled when settings fail to load', async () => {
    const user = userEvent.setup();
    mockedBridge.getSettings.mockRejectedValueOnce(new Error('Load failed'));

    render(<App />);

    await user.click(screen.getByRole('button', { name: /打开设置|open settings/i }));

    expect(screen.getByRole('button', { name: /保存|save/i })).toBeDisabled();
    expect(await screen.findByRole('alert')).toHaveTextContent(/无法从宿主加载设置/i);
  });

  it('disables the settings form while save is in flight', async () => {
    const user = userEvent.setup();
    const pendingSave = createDeferred<{
      apiKey: string;
      baseUrl: string;
      businessBaseUrl: string;
      model: string;
      ssoUrl: string;
      ssoLoginSuccessPath: string;
    }>();
    mockedBridge.saveSettings.mockReturnValueOnce(pendingSave.promise);

    render(<App />);

    await user.click(screen.getByRole('button', { name: /打开设置|open settings/i }));
    await user.click(screen.getByRole('button', { name: /保存|save/i }));

    expect(screen.getAllByLabelText(/^(api key|api 密钥)$/i)[0]).toBeDisabled();
    expect(screen.getByRole('textbox', { name: /^(base url|基础 url)$/i })).toBeDisabled();
    expect(screen.getByRole('textbox', { name: /^(business base url|业务基础 url)$/i })).toBeDisabled();
    expect(screen.getByRole('textbox', { name: /^(model|模型)$/i })).toBeDisabled();
    expect(screen.getByRole('button', { name: /关闭|close/i })).toBeDisabled();
    expect(screen.getByRole('button', { name: /取消|cancel/i })).toBeDisabled();
    expect(screen.getByRole('button', { name: /保存|save/i })).toBeDisabled();

    pendingSave.resolve({
      apiKey: '',
      baseUrl: 'https://api.example.com',
      businessBaseUrl: 'https://business.example.com',
      model: 'gpt-5-mini',
      ssoUrl: '',
      ssoLoginSuccessPath: '',
    });

    expect(
      await screen.findByRole('button', { name: /打开设置|open settings/i }),
    ).toHaveFocus();
  });

  it('updates the selection badge when native selection events arrive', async () => {
    render(<App />);

    expect(
      await screen.findByText(/sheet1 · a1:c4/i),
    ).toBeInTheDocument();

    selectionContextListener?.({
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

    expect(
      await screen.findByText(/sheet2 · b2:d5/i),
    ).toBeInTheDocument();
    expect(screen.queryByText(/non-contiguous selection/i)).not.toBeInTheDocument();
    expect(screen.queryByText(/multiple selection areas are not supported yet/i)).not.toBeInTheDocument();
  });

  it('shows the empty selection state when the native host has no selection context', async () => {
    mockedBridge.getSelectionContext.mockResolvedValueOnce({
      hasSelection: false,
      workbookName: '',
      sheetName: '',
      address: '',
      rowCount: 0,
      columnCount: 0,
      isContiguous: true,
      headerPreview: [],
      sampleRows: [],
      warningMessage: 'No selection available.',
    });

    render(<App />);

    expect(
      await screen.findByText(/^未选中$|^no selection$/i),
    ).toBeInTheDocument();
  });

  it('submits read-selection commands with Enter without requiring confirmation', async () => {
    const user = userEvent.setup();

    render(<App />);

    await user.type(screen.getByRole('textbox', { name: /消息输入框|message composer/i }), 'read selection{enter}');

    expect(mockedBridge.executeExcelCommand).toHaveBeenCalledWith({
      commandType: 'excel.readSelectionTable',
      confirmed: false,
    });
    expect(screen.queryByText(/确认 Excel 操作|confirm excel action/i)).not.toBeInTheDocument();
    expect(
      await screen.findByText(/read selection from sheet1 a1:c4/i),
    ).toBeInTheDocument();
    expect(screen.getByText(/name \| region \| amount/i)).toBeInTheDocument();
    expect(screen.getByText(/project a \| cn \| 42/i)).toBeInTheDocument();
  });

  it('keeps Shift+Enter for multiline composer input', async () => {
    const user = userEvent.setup();

    render(<App />);

    const composer = screen.getByRole('textbox', { name: /消息输入框|message composer/i });
    await user.type(composer, 'line 1');
    await user.keyboard('{Shift>}{Enter}{/Shift}');
    await user.type(composer, 'line 2');

    expect(mockedBridge.executeExcelCommand).not.toHaveBeenCalled();
    expect(mockedBridge.runSkill).not.toHaveBeenCalled();
    expect(mockedBridge.runAgent).not.toHaveBeenCalled();
    expect(composer).toHaveValue('line 1\nline 2');
  });

  it('routes plain natural language through the agent bridge', async () => {
    const user = userEvent.setup();
    mockedBridge.runAgent.mockResolvedValueOnce({
      route: 'chat',
      requiresConfirmation: false,
      status: 'completed',
      message: 'I can help with the current selection.',
    });

    render(<App />);

    await user.type(screen.getByRole('textbox', { name: /消息输入框|message composer/i }), 'Create a summary sheet from the current selection');
    await user.click(screen.getByRole('button', { name: /发送|send/i }));

    expect(mockedBridge.runAgent).toHaveBeenCalledWith({
      userInput: 'Create a summary sheet from the current selection',
      confirmed: false,
      sessionId: expect.any(String),
      conversationHistory: expect.any(Array),
    });
    expect(mockedBridge.runSkill).not.toHaveBeenCalled();
    expect(
      await screen.findByText(/i can help with the current selection/i),
    ).toBeInTheDocument();
  });

  it('renders a plan preview and confirms the frozen plan through the agent bridge', async () => {
    const user = userEvent.setup();
    const frozenPlan = {
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
    };

    mockedBridge.runAgent
      .mockResolvedValueOnce({
        route: 'plan',
        requiresConfirmation: true,
        status: 'preview',
        message: 'I prepared a plan. Review it before Excel is changed.',
        planner: {
          mode: 'plan',
          assistantMessage: 'I prepared a plan. Review it before Excel is changed.',
          plan: frozenPlan,
        },
      })
      .mockResolvedValueOnce({
        route: 'plan',
        requiresConfirmation: false,
        status: 'completed',
        message: 'Plan executed successfully.',
        journal: {
          hasFailures: false,
          errorMessage: '',
          steps: [
            {
              type: 'excel.addWorksheet',
              title: 'Add worksheet Summary',
              status: 'completed',
              message: 'Worksheet "Summary" created.',
              errorMessage: '',
            },
            {
              type: 'excel.writeRange',
              title: 'Write range Summary!A1:B3',
              status: 'completed',
              message: 'Wrote 3 row(s) to Summary!A1:B3.',
              errorMessage: '',
            },
          ],
        },
      });

    render(<App />);

    await user.type(screen.getByRole('textbox', { name: /消息输入框|message composer/i }), 'Create a summary sheet from the current selection');
    await user.click(screen.getByRole('button', { name: /发送|send/i }));

    expect(
      await screen.findByText(/create a summary sheet and write the selected rows/i),
    ).toBeInTheDocument();
    expect(screen.getByText(/i prepared a plan\. review it before excel is changed/i)).toBeInTheDocument();

    await user.click(screen.getByRole('button', { name: /确认|confirm/i }));

    expect(mockedBridge.runAgent).toHaveBeenNthCalledWith(1, {
      userInput: 'Create a summary sheet from the current selection',
      confirmed: false,
      sessionId: expect.any(String),
      conversationHistory: expect.any(Array),
    });
    expect(mockedBridge.runAgent).toHaveBeenNthCalledWith(2, {
      userInput: 'Create a summary sheet from the current selection',
      confirmed: true,
      sessionId: expect.any(String),
      plan: frozenPlan,
      conversationHistory: expect.any(Array),
    });
    expect(
      await screen.findByText(/plan executed successfully/i),
    ).toBeInTheDocument();
    expect(screen.getByText(/已完成 · add worksheet summary/i)).toBeInTheDocument();
    expect(screen.getByText(/已完成 · write range summary!a1:b3/i)).toBeInTheDocument();
  });

  it('queues write commands for confirmation before executing them', async () => {
    const user = userEvent.setup();
    mockedBridge.executeExcelCommand
      .mockResolvedValueOnce({
        commandType: 'excel.addWorksheet',
        requiresConfirmation: true,
        status: 'preview',
        message: 'Confirm worksheet creation before Excel is modified.',
        preview: {
          title: 'Add worksheet',
          summary: 'Add worksheet "Summary"',
          details: ['Workbook: Quarterly Report.xlsx'],
        },
      })
      .mockResolvedValueOnce({
        commandType: 'excel.addWorksheet',
        requiresConfirmation: false,
        status: 'completed',
        message: 'Worksheet "Summary" created.',
      });

    render(<App />);

    await user.type(screen.getByRole('textbox', { name: /消息输入框|message composer/i }), 'add sheet Summary');
    await user.click(screen.getByRole('button', { name: /发送|send/i }));

    expect(
      await screen.findByText(/确认 Excel 操作|confirm excel action/i),
    ).toBeInTheDocument();
    expect(screen.getByText(/确认此 excel 操作后再修改工作簿/i)).toBeInTheDocument();
    expect(screen.getByText(/新增工作表.*summary/i)).toBeInTheDocument();

    await user.click(screen.getByRole('button', { name: /确认|confirm/i }));

    expect(mockedBridge.executeExcelCommand).toHaveBeenNthCalledWith(1, {
      commandType: 'excel.addWorksheet',
      newSheetName: 'Summary',
      confirmed: false,
    });
    expect(mockedBridge.executeExcelCommand).toHaveBeenNthCalledWith(2, {
      commandType: 'excel.addWorksheet',
      newSheetName: 'Summary',
      confirmed: true,
    });
    expect(
      await screen.findByText(/worksheet "summary" created/i),
    ).toBeInTheDocument();
  });

  it('preserves extra host Excel preview details while localizing the confirmation card', async () => {
    const user = userEvent.setup();
    mockedBridge.executeExcelCommand.mockResolvedValueOnce({
      commandType: 'excel.addWorksheet',
      requiresConfirmation: true,
      status: 'preview',
      message: 'Confirm worksheet creation before Excel is modified.',
      preview: {
        title: 'Add worksheet',
        summary: 'Add worksheet "Summary"',
        details: [
          'Workbook: Quarterly Report.xlsx',
          'This action also updates formulas in the workbook.',
        ],
      },
    });

    render(<App />);

    await user.type(screen.getByRole('textbox', { name: /消息输入框|message composer/i }), 'add sheet Summary');
    await user.click(screen.getByRole('button', { name: /发送|send/i }));

    const confirmationCard = await screen.findByRole('article', { name: /确认 Excel 操作|confirm excel action/i });
    expect(within(confirmationCard).getByText(/工作簿：quarterly report\.xlsx/i)).toBeInTheDocument();
    expect(within(confirmationCard).getByText(/this action also updates formulas in the workbook/i)).toBeInTheDocument();
  });

  it('disables the composer while a confirmation card is pending', async () => {
    const user = userEvent.setup();
    mockedBridge.runSkill.mockResolvedValueOnce({
      route: 'skill',
      skillName: 'upload_data',
      requiresConfirmation: true,
      status: 'preview',
      message: 'Review the upload payload before sending it to Project A.',
      preview: {
        title: 'Upload selected data',
        summary: 'Upload 2 row(s) to Project A',
        details: ['Source: Sheet1!A1:C3', 'Fields: Name, Region'],
      },
      uploadPreview: {
        projectName: 'Project A',
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

    render(<App />);

    const composer = screen.getByRole('textbox', { name: /消息输入框|message composer/i });
    const sendButton = screen.getByRole('button', { name: /发送|send/i });
    await user.type(composer, '/upload_data upload selected data to Project A');
    await user.click(sendButton);

    expect(
      await screen.findByText(/确认 Excel 操作|confirm excel action/i),
    ).toBeInTheDocument();
    expect(composer).toBeDisabled();
    expect(sendButton).toBeDisabled();
    expect(mockedBridge.runSkill).toHaveBeenCalledTimes(1);
  });

  it('keeps command results in the session that launched them when the user switches threads mid-flight', async () => {
    const user = userEvent.setup();
    const pendingCommand = createDeferred<{
      commandType: string;
      requiresConfirmation: boolean;
      status: string;
      message: string;
    }>();
    mockedBridge.executeExcelCommand.mockReturnValueOnce(pendingCommand.promise);

    render(<App />);

    await user.click(await screen.findByRole('button', { name: /打开会话列表|open sessions drawer/i }));
    const sidebar = await screen.findByRole('complementary', { name: /会话抽屉|sessions drawer/i });
    await screen.findByRole('heading', { name: /未命名会话|untitled/i });

    await user.type(screen.getByRole('textbox', { name: /消息输入框|message composer/i }), 'read selection');
    await user.click(screen.getByRole('button', { name: /发送|send/i }));
    await user.click(within(sidebar).getByRole('button', { name: /review notes/i }));

    expect(
      screen.getByRole('heading', { name: /review notes/i }),
    ).toBeInTheDocument();

    pendingCommand.resolve({
      commandType: 'excel.readSelectionTable',
      requiresConfirmation: false,
      status: 'completed',
      message: 'Read selection from Sheet1 A1:C4.',
    });

    expect(
      screen.queryByText(/read selection from sheet1 a1:c4/i),
    ).not.toBeInTheDocument();

    await user.click(screen.getByRole('button', { name: /打开会话列表|open sessions drawer/i }));
    const reopenedSidebar = await screen.findByRole('complementary', { name: /会话抽屉|sessions drawer/i });
    await user.click(within(reopenedSidebar).getByText(/read selection/i));

    expect(
      await screen.findByText(/read selection from sheet1 a1:c4/i),
    ).toBeInTheDocument();
  });

  it('routes explicit slash upload_data through the skill bridge and confirms with the returned preview payload', async () => {
    const user = userEvent.setup();
    const uploadPreview = {
      projectName: '项目A',
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
    };

    mockedBridge.runSkill
      .mockResolvedValueOnce({
        route: 'skill',
        skillName: 'upload_data',
        requiresConfirmation: true,
        status: 'preview',
        message: 'Review the upload payload before sending it to 项目A.',
        preview: {
          title: 'Upload selected data',
          summary: 'Upload 2 row(s) to 项目A',
          details: ['Source: Sheet1!A1:C3', 'Fields: Name, Region'],
        },
        uploadPreview,
      })
      .mockResolvedValueOnce({
        route: 'skill',
        skillName: 'upload_data',
        requiresConfirmation: false,
        status: 'completed',
        message: 'Uploaded 2 row(s) to 项目A.',
        uploadPreview,
      });

    render(<App />);

    await user.type(screen.getByRole('textbox', { name: /消息输入框|message composer/i }), '把选中数据上传到项目A');
    await user.click(screen.getByRole('button', { name: /发送|send/i }));

    expect(mockedBridge.runSkill).toHaveBeenNthCalledWith(1, {
      userInput: '把选中数据上传到项目A',
      confirmed: false,
    });
    expect(screen.getByText(/请先确认发往项目a的上传内容/i)).toBeInTheDocument();
    expect(
      await screen.findByText(/上传 2 行数据到 项目a/i),
    ).toBeInTheDocument();
    expect(screen.getByText(/上传所选数据/i)).toBeInTheDocument();
    expect(screen.getByText(/来源：sheet1!a1:c3/i)).toBeInTheDocument();
    expect(screen.getByText(/字段：name, region/i)).toBeInTheDocument();

    await user.click(screen.getByRole('button', { name: /确认|confirm/i }));

    expect(mockedBridge.runSkill).toHaveBeenNthCalledWith(2, {
      userInput: '把选中数据上传到项目A',
      skillName: 'upload_data',
      confirmed: true,
      uploadPreview,
    });
    expect(
      await screen.findByText(/uploaded 2 row\(s\) to 项目a/i),
    ).toBeInTheDocument();
  });

  it('preserves extra host upload preview details while localizing the confirmation card', async () => {
    const user = userEvent.setup();
    mockedBridge.runSkill.mockResolvedValueOnce({
      route: 'skill',
      skillName: 'upload_data',
      requiresConfirmation: true,
      status: 'preview',
      message: 'Review the upload payload before sending it to 项目A.',
      preview: {
        title: 'Upload selected data',
        summary: 'Upload 2 row(s) to 项目A',
        details: [
          'Source: Sheet1!A1:C3',
          'Fields: Name, Region',
          'Rows with blank IDs will be skipped.',
        ],
      },
      uploadPreview: {
        projectName: '项目A',
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

    render(<App />);

    await user.type(screen.getByRole('textbox', { name: /消息输入框|message composer/i }), '把选中数据上传到项目A');
    await user.click(screen.getByRole('button', { name: /发送|send/i }));

    const confirmationCard = await screen.findByRole('article', { name: /确认 Excel 操作|confirm excel action/i });
    expect(within(confirmationCard).getByText(/来源：sheet1!a1:c3/i)).toBeInTheDocument();
    expect(within(confirmationCard).getByText(/字段：name, region/i)).toBeInTheDocument();
    expect(within(confirmationCard).getByText(/rows with blank ids will be skipped/i)).toBeInTheDocument();
  });

  it('auto-scrolls the message thread when messages change', async () => {
    const user = userEvent.setup();
    mockedBridge.runAgent.mockResolvedValueOnce({
      route: 'chat',
      requiresConfirmation: false,
      status: 'completed',
      message: 'I can help with the current selection.',
    });

    render(<App />);

    expect(await screen.findByRole('region', { name: /消息线程|message thread/i })).toBeInTheDocument();
    expect(scrollToSpy).toHaveBeenCalled();

    scrollToSpy.mockClear();
    await user.type(screen.getByRole('textbox', { name: /消息输入框|message composer/i }), 'Create a summary sheet{enter}');

    expect(await screen.findByText(/i can help with the current selection/i)).toBeInTheDocument();
    expect(scrollToSpy).toHaveBeenCalled();
  });
});
