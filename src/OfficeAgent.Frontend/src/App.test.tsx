import { cleanup, render, screen, within } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import App from './App';
import { nativeBridge } from './bridge/nativeBridge';

vi.mock('./bridge/nativeBridge', () => ({
  nativeBridge: {
    ping: vi.fn(),
    getSessions: vi.fn(),
    getSettings: vi.fn(),
    saveSettings: vi.fn(),
  },
}));

const mockedBridge = vi.mocked(nativeBridge);

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
  mockedBridge.ping.mockResolvedValue({
    host: 'browser-preview',
    version: 'dev',
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
    model: 'gpt-5-mini',
  });
  mockedBridge.saveSettings.mockImplementation(async (settings) => settings);
});

afterEach(() => {
  cleanup();
  vi.clearAllMocks();
});

describe('App shell', () => {
  it('renders the expected task pane regions', async () => {
    const user = userEvent.setup();

    render(<App />);

    expect(
      screen.getByRole('complementary', { name: /session sidebar placeholder/i }),
    ).toBeInTheDocument();
    expect(screen.getByRole('banner', { name: /chat header/i })).toBeInTheDocument();
    expect(
      screen.getByRole('region', { name: /message thread/i }),
    ).toBeInTheDocument();
    expect(
      screen.getByText(/welcome to office agent/i),
    ).toBeInTheDocument();
    expect(
      screen.getByRole('status', { name: /selection badge placeholder/i }),
    ).toBeInTheDocument();
    expect(
      screen.getByRole('textbox', { name: /message composer/i }),
    ).toBeInTheDocument();
    expect(
      screen.getByRole('button', { name: /settings/i }),
    ).toBeInTheDocument();
    expect(
      await screen.findByText(/connected to browser-preview \(dev\)/i),
    ).toBeInTheDocument();
    expect(
      await screen.findByRole('button', { name: /browser preview/i }),
    ).toBeInTheDocument();

    await user.click(screen.getByRole('button', { name: /settings/i }));

    expect(
      screen.getByRole('dialog', { name: /settings dialog/i }),
    ).toBeInTheDocument();
    expect(
      screen.getByRole('textbox', { name: /api key/i }),
    ).toHaveValue('');
    expect(
      screen.getByRole('textbox', { name: /base url/i }),
    ).toHaveValue('https://api.example.com');
    expect(
      screen.getByRole('textbox', { name: /model/i }),
    ).toHaveValue('gpt-5-mini');
  });

  it('switches the active session when a session chip is clicked', async () => {
    const user = userEvent.setup();

    render(<App />);
    const sidebar = screen.getByRole('complementary', { name: /session sidebar placeholder/i });

    expect(
      await screen.findByRole('heading', { name: /browser preview/i }),
    ).toBeInTheDocument();

    await user.click(within(sidebar).getByRole('button', { name: /review notes/i }));

    expect(
      screen.getByRole('heading', { name: /review notes/i }),
    ).toBeInTheDocument();
  });

  it('resets unsaved settings changes when the dialog is cancelled or closed', async () => {
    const user = userEvent.setup();

    render(<App />);

    await user.click(screen.getByRole('button', { name: /settings/i }));
    await user.clear(screen.getByRole('textbox', { name: /base url/i }));
    await user.type(screen.getByRole('textbox', { name: /base url/i }), 'https://changed.example.com');
    await user.click(screen.getByRole('button', { name: /cancel/i }));

    await user.click(screen.getByRole('button', { name: /settings/i }));
    expect(
      screen.getByRole('textbox', { name: /base url/i }),
    ).toHaveValue('https://api.example.com');

    await user.clear(screen.getByRole('textbox', { name: /base url/i }));
    await user.type(screen.getByRole('textbox', { name: /base url/i }), 'https://closed.example.com');
    await user.click(screen.getByRole('button', { name: /close/i }));

    await user.click(screen.getByRole('button', { name: /settings/i }));
    expect(
      screen.getByRole('textbox', { name: /base url/i }),
    ).toHaveValue('https://api.example.com');
  });

  it('shows an inline error when saving settings fails', async () => {
    const user = userEvent.setup();
    mockedBridge.saveSettings.mockRejectedValueOnce(new Error('Bridge write failed'));

    render(<App />);

    await user.click(screen.getByRole('button', { name: /settings/i }));
    await user.click(screen.getByRole('button', { name: /save/i }));

    expect(
      await screen.findByRole('alert'),
    ).toBeInTheDocument();
    expect(screen.getByRole('alert')).toHaveTextContent(/bridge write failed/i);
    expect(
      screen.getByRole('dialog', { name: /settings dialog/i }),
    ).toBeInTheDocument();
  });

  it('does not overwrite unsaved settings edits when settings load late', async () => {
    const user = userEvent.setup();
    const delayedSettings = createDeferred<{
      apiKey: string;
      baseUrl: string;
      model: string;
    }>();
    mockedBridge.getSettings.mockReturnValueOnce(delayedSettings.promise);

    render(<App />);

    await user.click(screen.getByRole('button', { name: /settings/i }));
    const baseUrlInput = screen.getByRole('textbox', { name: /base url/i });
    await user.clear(baseUrlInput);
    await user.type(baseUrlInput, 'https://draft.example.com');

    delayedSettings.resolve({
      apiKey: '',
      baseUrl: 'https://loaded.example.com',
      model: 'gpt-5-mini',
    });

    expect(
      await screen.findByDisplayValue('https://draft.example.com'),
    ).toBeInTheDocument();
  });

  it('treats the settings dialog as a modal and restores focus on close', async () => {
    const user = userEvent.setup();

    render(<App />);

    const settingsButton = screen.getByRole('button', { name: /settings/i });
    await user.click(settingsButton);

    const dialog = screen.getByRole('dialog', { name: /settings dialog/i });
    const apiKeyInput = screen.getByRole('textbox', { name: /api key/i });

    expect(dialog).toHaveAttribute('aria-modal', 'true');
    expect(apiKeyInput).toHaveFocus();

    await user.tab({ shift: true });
    expect(screen.getByRole('button', { name: /close/i })).toHaveFocus();

    await user.tab({ shift: true });
    expect(screen.getByRole('button', { name: /save/i })).toHaveFocus();

    await user.click(screen.getByRole('button', { name: /close/i }));
    expect(settingsButton).toHaveFocus();
  });

  it('disables save until settings finish loading', async () => {
    const user = userEvent.setup();
    const delayedSettings = createDeferred<{
      apiKey: string;
      baseUrl: string;
      model: string;
    }>();
    mockedBridge.getSettings.mockReturnValueOnce(delayedSettings.promise);

    render(<App />);

    await user.click(screen.getByRole('button', { name: /settings/i }));
    expect(screen.getByRole('button', { name: /save/i })).toBeDisabled();

    delayedSettings.resolve({
      apiKey: 'loaded-key',
      baseUrl: 'https://loaded.example.com',
      model: 'gpt-5-mini',
    });

    expect(
      await screen.findByDisplayValue('loaded-key'),
    ).toBeInTheDocument();
    expect(screen.getByRole('button', { name: /save/i })).toBeEnabled();
  });

  it('keeps save disabled when settings fail to load', async () => {
    const user = userEvent.setup();
    mockedBridge.getSettings.mockRejectedValueOnce(new Error('Load failed'));

    render(<App />);

    await user.click(screen.getByRole('button', { name: /settings/i }));

    expect(screen.getByRole('button', { name: /save/i })).toBeDisabled();
    expect(await screen.findByRole('alert')).toHaveTextContent(/unable to load settings/i);
  });

  it('disables the settings form while save is in flight', async () => {
    const user = userEvent.setup();
    const pendingSave = createDeferred<{
      apiKey: string;
      baseUrl: string;
      model: string;
    }>();
    mockedBridge.saveSettings.mockReturnValueOnce(pendingSave.promise);

    render(<App />);

    await user.click(screen.getByRole('button', { name: /settings/i }));
    await user.click(screen.getByRole('button', { name: /save/i }));

    expect(screen.getByRole('textbox', { name: /api key/i })).toBeDisabled();
    expect(screen.getByRole('textbox', { name: /base url/i })).toBeDisabled();
    expect(screen.getByRole('textbox', { name: /model/i })).toBeDisabled();
    expect(screen.getByRole('button', { name: /close/i })).toBeDisabled();
    expect(screen.getByRole('button', { name: /cancel/i })).toBeDisabled();
    expect(screen.getByRole('button', { name: /save/i })).toBeDisabled();

    pendingSave.resolve({
      apiKey: '',
      baseUrl: 'https://api.example.com',
      model: 'gpt-5-mini',
    });

    expect(
      await screen.findByRole('button', { name: /settings/i }),
    ).toHaveFocus();
  });
});
