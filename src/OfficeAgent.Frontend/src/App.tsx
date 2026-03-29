import { useEffect, useRef, useState, type KeyboardEvent as ReactKeyboardEvent } from 'react';
import { nativeBridge } from './bridge/nativeBridge';
import type { AppSettings, ChatSession } from './types/bridge';

const DEFAULT_SETTINGS: AppSettings = {
  apiKey: '',
  baseUrl: 'https://api.example.com',
  model: 'gpt-5-mini',
};

export function App() {
  const [bridgeStatus, setBridgeStatus] = useState('Connecting to native host...');
  const [sessions, setSessions] = useState<ChatSession[]>([]);
  const [activeSessionId, setActiveSessionId] = useState('');
  const [settings, setSettings] = useState<AppSettings | null>(null);
  const [draftSettings, setDraftSettings] = useState<AppSettings>(DEFAULT_SETTINGS);
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [settingsLoadError, setSettingsLoadError] = useState('');
  const [settingsSaveError, setSettingsSaveError] = useState('');
  const [isSettingsLoading, setIsSettingsLoading] = useState(true);
  const [isSettingsSaving, setIsSettingsSaving] = useState(false);
  const settingsButtonRef = useRef<HTMLButtonElement | null>(null);
  const settingsDialogRef = useRef<HTMLElement | null>(null);
  const apiKeyInputRef = useRef<HTMLInputElement | null>(null);
  const isSettingsOpenRef = useRef(false);
  const isSettingsDirtyRef = useRef(false);
  const shouldRestoreSettingsButtonFocusRef = useRef(false);

  useEffect(() => {
    let isActive = true;

    nativeBridge
      .ping()
      .then((result) => {
        if (!isActive) {
          return;
        }

        setBridgeStatus(`Connected to ${result.host} (${result.version})`);
      })
      .catch((error: Error) => {
        if (!isActive) {
          return;
        }

        setBridgeStatus(`Native bridge unavailable: ${error.message}`);
      });

    nativeBridge
      .getSessions()
      .then((result) => {
        if (!isActive) {
          return;
        }

        setSessions(result.sessions);
        setActiveSessionId(result.activeSessionId);
      })
      .catch(() => {
        if (!isActive) {
          return;
        }

        setSessions([]);
        setActiveSessionId('');
      });

    nativeBridge
      .getSettings()
      .then((result) => {
        if (!isActive) {
          return;
        }

        setSettings(result);
        if (!(isSettingsOpenRef.current && isSettingsDirtyRef.current)) {
          setDraftSettings(result);
        }
        setIsSettingsLoading(false);
        setSettingsLoadError('');
        setSettingsSaveError('');
      })
      .catch(() => {
        if (!isActive) {
          return;
        }

        setSettings(null);
        if (!(isSettingsOpenRef.current && isSettingsDirtyRef.current)) {
          setDraftSettings(DEFAULT_SETTINGS);
        }
        setIsSettingsLoading(false);
        setSettingsLoadError('Unable to load settings from native host.');
        setSettingsSaveError('');
      });

    return () => {
      isActive = false;
    };
  }, []);

  const activeSession = sessions.find((session) => session.id === activeSessionId) ?? sessions[0];

  useEffect(() => {
    if (isSettingsOpen) {
      apiKeyInputRef.current?.focus();
      return;
    }

    if (shouldRestoreSettingsButtonFocusRef.current) {
      settingsButtonRef.current?.focus();
      shouldRestoreSettingsButtonFocusRef.current = false;
    }
  }, [isSettingsOpen]);

  function resetDraftSettings() {
    setDraftSettings(settings ?? DEFAULT_SETTINGS);
    isSettingsDirtyRef.current = false;
    setSettingsSaveError('');
  }

  function openSettings() {
    resetDraftSettings();
    isSettingsOpenRef.current = true;
    setIsSettingsOpen(true);
  }

  function closeSettings() {
    resetDraftSettings();
    isSettingsOpenRef.current = false;
    shouldRestoreSettingsButtonFocusRef.current = true;
    setIsSettingsOpen(false);
  }

  function updateDraftSettings(update: Partial<AppSettings>) {
    isSettingsDirtyRef.current = true;
    setDraftSettings((current) => ({ ...current, ...update }));
  }

  function handleSettingsDialogKeyDown(event: ReactKeyboardEvent<HTMLElement>) {
    if (event.key !== 'Tab') {
      return;
    }

    const focusableElements = settingsDialogRef.current?.querySelectorAll<HTMLElement>(
      'button:not([disabled]), input:not([disabled]), textarea:not([disabled]), select:not([disabled]), [tabindex]:not([tabindex="-1"])',
    );

    if (!focusableElements || focusableElements.length === 0) {
      return;
    }

    const firstFocusableElement = focusableElements[0];
    const lastFocusableElement = focusableElements[focusableElements.length - 1];

    if (event.shiftKey && document.activeElement === firstFocusableElement) {
      event.preventDefault();
      lastFocusableElement.focus();
      return;
    }

    if (!event.shiftKey && document.activeElement === lastFocusableElement) {
      event.preventDefault();
      firstFocusableElement.focus();
    }
  }

  async function handleSettingsSave() {
    if (isSettingsLoading || isSettingsSaving || settingsLoadError) {
      return;
    }

    setIsSettingsSaving(true);
    setSettingsSaveError('');

    try {
      const savedSettings = await nativeBridge.saveSettings(draftSettings);
      setSettings(savedSettings);
      setDraftSettings(savedSettings);
      isSettingsDirtyRef.current = false;
      isSettingsOpenRef.current = false;
      shouldRestoreSettingsButtonFocusRef.current = true;
      setIsSettingsOpen(false);
    } catch (error) {
      setSettingsSaveError(error instanceof Error ? error.message : 'Unable to save settings.');
    } finally {
      setIsSettingsSaving(false);
    }
  }

  return (
    <div className="app-shell">
      <aside className="sidebar" aria-label="Session sidebar placeholder">
        <div className="sidebar__title">Sessions</div>
        {sessions.length === 0 ? (
          <div className="sidebar__empty">No sessions yet</div>
        ) : (
          <div className="sidebar__list">
            {sessions.map((session) => (
              <button
                key={session.id}
                type="button"
                className={`session-chip${session.id === activeSession?.id ? ' session-chip--active' : ''}`}
                onClick={() => setActiveSessionId(session.id)}
              >
                {session.title}
              </button>
            ))}
          </div>
        )}
      </aside>

      <main className="workspace">
        <header className="chat-header" aria-label="Chat header">
          <div>
            <div className="eyebrow">Office Agent</div>
            <h1 className="title">{activeSession?.title ?? 'Task pane shell'}</h1>
            <div className="subtitle">{settings?.baseUrl ?? 'Settings not loaded yet'}</div>
          </div>

          <button
            type="button"
            className="icon-button"
            aria-label="Settings"
            ref={settingsButtonRef}
            onClick={openSettings}
          >
            Settings
          </button>
        </header>

        <section className="selection-badge" aria-label="Selection badge placeholder" role="status">
          <div className="selection-badge__label">Bridge status</div>
          <div>{bridgeStatus}</div>
        </section>

        <section className="thread" aria-label="Message thread">
          <article className="message message--assistant">
            <p>Welcome to Office Agent. This shell is ready for the chat experience.</p>
          </article>
        </section>

        <footer className="composer" aria-label="Message composer">
          <textarea
            aria-label="Message composer"
            placeholder="Type a message..."
            rows={3}
          />
          <button type="button" className="send-button">
            Send
          </button>
        </footer>
      </main>

      {isSettingsOpen ? (
        <div className="settings-backdrop">
          <section
            ref={settingsDialogRef}
            className="settings-dialog"
            role="dialog"
            aria-modal="true"
            aria-label="Settings dialog"
            onKeyDown={handleSettingsDialogKeyDown}
          >
            <div className="settings-dialog__header">
              <div>
                <div className="eyebrow">Configuration</div>
                <h2 className="settings-dialog__title">Settings</h2>
              </div>
              <button type="button" className="icon-button" onClick={closeSettings} disabled={isSettingsSaving}>
                Close
              </button>
            </div>

            {settingsLoadError ? <p className="settings-error" role="alert">{settingsLoadError}</p> : null}
            {settingsSaveError ? <p className="settings-error" role="alert">{settingsSaveError}</p> : null}

            <label className="settings-field">
              <span>API Key</span>
              <input
                ref={apiKeyInputRef}
                aria-label="API Key"
                type="text"
                value={draftSettings.apiKey}
                disabled={isSettingsSaving}
                onChange={(event) => updateDraftSettings({ apiKey: event.target.value })}
              />
            </label>

            <label className="settings-field">
              <span>Base URL</span>
              <input
                aria-label="Base URL"
                type="text"
                value={draftSettings.baseUrl}
                disabled={isSettingsSaving}
                onChange={(event) => updateDraftSettings({ baseUrl: event.target.value })}
              />
            </label>

            <label className="settings-field">
              <span>Model</span>
              <input
                aria-label="Model"
                type="text"
                value={draftSettings.model}
                disabled={isSettingsSaving}
                onChange={(event) => updateDraftSettings({ model: event.target.value })}
              />
            </label>

            <div className="settings-actions">
              <button type="button" className="ghost-button" onClick={closeSettings} disabled={isSettingsSaving}>
                Cancel
              </button>
              <button
                type="button"
                className="send-button"
                onClick={handleSettingsSave}
                disabled={isSettingsLoading || isSettingsSaving || Boolean(settingsLoadError)}
              >
                Save
              </button>
            </div>
          </section>
        </div>
      ) : null}
    </div>
  );
}

export default App;
