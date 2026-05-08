import { useEffect, useRef, useState, type CSSProperties, type KeyboardEvent as ReactKeyboardEvent } from 'react';
import { nativeBridge } from './bridge/nativeBridge';
import { ConfirmationCard } from './components/ConfirmationCard';
import { getUiStrings, UNTITLED_SESSION_STORAGE_TITLE } from './i18n/uiStrings';
import type {
  AgentPlan,
  AgentRequestEnvelope,
  AgentResult,
  AppSettings,
  ChatSession,
  ConversationTurn,
  ExcelCommand,
  ExcelCommandPreview,
  ExcelCommandResult,
  ExcelTableData,
  HostContext,
  LoginStatus,
  SelectionContext,
  SkillRequestEnvelope,
  SkillResult,
  UiLocale,
  UploadPreview,
} from './types/bridge';

const DEFAULT_SETTINGS: AppSettings = {
  apiKey: '',
  baseUrl: 'https://api.example.com',
  businessBaseUrl: '',
  model: 'gpt-5-mini',
  ssoUrl: '',
  ssoLoginSuccessPath: '',
  uiLanguageOverride: 'system',
};

const SETTINGS_DIALOG_VIEWPORT_STYLE: CSSProperties = {
  maxHeight: 'calc(100vh - 32px)',
  overflowY: 'auto',
};

type ThreadMessage = {
  id: string;
  role: 'assistant' | 'user' | 'system';
  content: string;
  table?: ExcelTableData;
};

type PendingConfirmation = {
  kind: 'excel' | 'skill' | 'agent';
  command?: ExcelCommand;
  skillRequest?: SkillRequestEnvelope;
  agentRequest?: AgentRequestEnvelope;
  plan?: AgentPlan;
  preview: ExcelCommandPreview;
};

type BridgeStatusState =
  | { kind: 'connecting' }
  | { kind: 'connected'; host: string; version: string }
  | { kind: 'unavailable'; errorMessage: string };

export function App() {
  const [uiLocale, setUiLocale] = useState<UiLocale>('en');
  const strings = getUiStrings(uiLocale);
  const [bridgeStatus, setBridgeStatus] = useState<BridgeStatusState>({ kind: 'connecting' });
  const [sessions, setSessions] = useState<ChatSession[]>([]);
  const [activeSessionId, setActiveSessionId] = useState('');
  const [isSessionsDrawerOpen, setIsSessionsDrawerOpen] = useState(false);
  const [selectionContext, setSelectionContext] = useState<SelectionContext | null>(null);
  const [settings, setSettings] = useState<AppSettings | null>(null);
  const [draftSettings, setDraftSettings] = useState<AppSettings>(DEFAULT_SETTINGS);
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [settingsLoadError, setSettingsLoadError] = useState('');
  const [settingsSaveError, setSettingsSaveError] = useState('');
  const [isSettingsLoading, setIsSettingsLoading] = useState(true);
  const [isSettingsSaving, setIsSettingsSaving] = useState(false);
  const [showApiKey, setShowApiKey] = useState(false);
  const [composerValue, setComposerValue] = useState('');
  const [sessionThreads, setSessionThreads] = useState<Record<string, ThreadMessage[]>>({});
  const [pendingConfirmations, setPendingConfirmations] = useState<Record<string, PendingConfirmation>>({});
  const [pendingCommandSessions, setPendingCommandSessions] = useState<Record<string, boolean>>({});
  const [systemUntitledSessionIds, setSystemUntitledSessionIds] = useState<Record<string, true>>({});
  const sessionsButtonRef = useRef<HTMLButtonElement | null>(null);
  const settingsButtonRef = useRef<HTMLButtonElement | null>(null);
  const settingsDialogRef = useRef<HTMLElement | null>(null);
  const apiKeyInputRef = useRef<HTMLInputElement | null>(null);
  const threadRef = useRef<HTMLElement | null>(null);
  const isSettingsOpenRef = useRef(false);
  const isSettingsDirtyRef = useRef(false);
  const shouldRestoreSettingsButtonFocusRef = useRef(false);
  const saveTimeoutRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const [renamingSessionId, setRenamingSessionId] = useState<string | null>(null);
  const [renameValue, setRenameValue] = useState('');
  const [renameValueDirty, setRenameValueDirty] = useState(false);
  const [deleteConfirmSessionId, setDeleteConfirmSessionId] = useState<string | null>(null);
  const [loginStatus, setLoginStatus] = useState<LoginStatus | null>(null);
  const [isLoggingIn, setIsLoggingIn] = useState(false);
  const [loginError, setLoginError] = useState('');
  const bridgeStatusText = formatBridgeStatus(bridgeStatus, strings);

  useEffect(() => {
    document.documentElement.lang = uiLocale;
  }, [uiLocale]);

  useEffect(() => {
    setSessionThreads((current) => relocalizeWelcomeMessages(current, uiLocale));
  }, [uiLocale]);

  useEffect(() => {
    let isActive = true;
    let startupLocale: UiLocale = 'en';

    const hostContextRequest = nativeBridge.getHostContext().catch(() => null);
    void hostContextRequest.then((result) => {
      if (!isActive || !result) {
        return;
      }

      setUiLocale(result.resolvedUiLocale);
    });

    const hostContextPromise = readHostContextWithTimeout(hostContextRequest)
      .then((result) => {
        if (!isActive) {
          return;
        }

        startupLocale = result?.resolvedUiLocale ?? 'en';
        setUiLocale(startupLocale);
      });

    void hostContextPromise.finally(() => {
      nativeBridge
        .ping()
        .then((result) => {
          if (!isActive) {
            return;
          }

          setBridgeStatus({
            kind: 'connected',
            host: result.host,
            version: result.version,
          });
        })
        .catch((error: Error) => {
          if (!isActive) {
            return;
          }

          setBridgeStatus({
            kind: 'unavailable',
            errorMessage: error.message,
          });
        });

      nativeBridge
        .getSessions()
        .then((result) => {
          if (!isActive) {
            return;
          }

          const allSessions = result.sessions;
          const latestSession = allSessions[0];

          // Check if the most recent session is a usable new chat (empty " untitled and no messages)
          let reusableSession: ChatSession | undefined;
          if (latestSession && latestSession.isSystemUntitled === true && latestSession.messages.length === 0) {
            reusableSession = latestSession;
          }

          let newSessionId: string;
          let displaySessions: ChatSession[];

          if (reusableSession) {
            // Reuse the latest "New chat" session, keep all sessions in sidebar for reusing
            newSessionId = reusableSession.id;
            displaySessions = allSessions;
          } else {
            // Create a brand new session
            const id = createMessageId();
            const now = new Date().toISOString();
            const newSession: ChatSession = {
              id,
              title: UNTITLED_SESSION_STORAGE_TITLE,
              isSystemUntitled: true,
              createdAtUtc: now,
              updatedAtUtc: now,
              messages: [],
            };
            newSessionId = id;
            displaySessions = [newSession, ...allSessions];
          }

          setSessions(displaySessions);
          setSystemUntitledSessionIds(deriveSystemUntitledSessionIds(displaySessions));
          setSessionThreads((current) => hydrateSessionThreads(current, displaySessions, startupLocale));
          setActiveSessionId(newSessionId);
        })
        .catch(() => {
          if (!isActive) {
            return;
          }

          setSessions([]);
          setSystemUntitledSessionIds({});
          setActiveSessionId('');
        });

      nativeBridge
      .getSettings()
      .then((result) => {
        if (!isActive) {
          return;
        }

        const normalizedSettings = normalizeSettings(result);
        setSettings(normalizedSettings);
        if (!(isSettingsOpenRef.current && isSettingsDirtyRef.current)) {
          setDraftSettings(normalizedSettings);
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
          setSettingsLoadError(getUiStrings(startupLocale).settingsLoadFailed);
          setSettingsSaveError('');
        });
    });

    nativeBridge
      .getSelectionContext()
      .then((result) => {
        if (!isActive) {
          return;
        }

        setSelectionContext(result);
      })
      .catch(() => {
        if (!isActive) {
          return;
        }

        setSelectionContext(null);
      });

    const unsubscribeSelectionContext = nativeBridge.onSelectionContextChanged((result) => {
      if (!isActive) {
        return;
      }

      setSelectionContext(result);
    });

    nativeBridge
      .getLoginStatus()
      .then((result) => {
        if (!isActive) {
          return;
        }

        setLoginStatus(result);
      })
      .catch(() => {
        if (!isActive) {
          return;
        }

        setLoginStatus({ isLoggedIn: false, ssoUrl: '' });
      });

    return () => {
      isActive = false;
      unsubscribeSelectionContext();
    };
  }, []);

  const activeSession = sessions.find((session) => session.id === activeSessionId) ?? sessions[0];
  const activeThreadMessages = activeSession
    ? sessionThreads[activeSession.id] ?? createInitialThreadMessages(activeSession, uiLocale)
    : createInitialThreadMessages(undefined, uiLocale);
  const activePendingConfirmation = activeSession ? pendingConfirmations[activeSession.id] ?? null : null;
  const isCommandPending = activeSession ? pendingCommandSessions[activeSession.id] === true : false;
  const isComposerDisabled = isCommandPending || activePendingConfirmation !== null;

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

  useEffect(() => {
    const threadElement = threadRef.current;
    if (!threadElement) {
      return;
    }

    threadElement.scrollTo({
      top: threadElement.scrollHeight,
      behavior: 'auto',
    });
  }, [activeSession?.id, activeThreadMessages.length, isCommandPending]);

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

  async function handleLogin() {
    if (!draftSettings.ssoUrl.trim() || isLoggingIn) {
      return;
    }

    setIsLoggingIn(true);
    setLoginError('');

    try {
      const result = await nativeBridge.login({ ssoUrl: draftSettings.ssoUrl.trim(), ssoLoginSuccessPath: draftSettings.ssoLoginSuccessPath?.trim() });
      if (result.success) {
        setLoginStatus({ isLoggedIn: true, ssoUrl: draftSettings.ssoUrl.trim() });
      } else {
        setLoginError(result.error ?? strings.loginFailed);
      }
    } catch (error) {
      setLoginError(error instanceof Error ? error.message : strings.loginFailed);
    } finally {
      setIsLoggingIn(false);
    }
  }

  async function handleLogout() {
    try {
      await nativeBridge.logout();
      setLoginStatus({ isLoggedIn: false, ssoUrl: loginStatus?.ssoUrl ?? '' });
    } catch {
      // best-effort
    }
  }

  function toggleSessionsDrawer() {
    setIsSessionsDrawerOpen((current) => !current);
  }

  function closeSessionsDrawer() {
    setIsSessionsDrawerOpen(false);
    sessionsButtonRef.current?.focus();
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

  function handleComposerKeyDown(event: ReactKeyboardEvent<HTMLTextAreaElement>) {
    if (event.key !== 'Enter' || event.shiftKey || event.nativeEvent.isComposing) {
      return;
    }

    event.preventDefault();
    void handleComposerSend();
  }

  function handleSessionSelect(sessionId: string) {
    if (sessionId !== activeSessionId) {
      void saveCurrentSessions();
    }

    setActiveSessionId(sessionId);
    setIsSessionsDrawerOpen(false);
  }

  function handleCreateNewSession() {
    void saveCurrentSessions();
    const id = createMessageId();
    const now = new Date().toISOString();
    const newSession: ChatSession = {
      id,
      title: UNTITLED_SESSION_STORAGE_TITLE,
      isSystemUntitled: true,
      createdAtUtc: now,
      updatedAtUtc: now,
      messages: [],
    };
    setSessions((current) => [newSession, ...current]);
    setSessionThreads((current) => ({
      ...current,
      [id]: createInitialThreadMessages(undefined, uiLocale),
    }));
    setSystemUntitledSessionIds((current) => ({
      ...current,
      [id]: true,
    }));
    setActiveSessionId(id);
    setIsSessionsDrawerOpen(false);
  }

  function handleRenameStart(session: ChatSession) {
    setRenamingSessionId(session.id);
    setRenameValue(getSessionDisplayTitle(session, systemUntitledSessionIds, strings));
    setRenameValueDirty(false);
  }

  function handleRenameConfirm() {
    if (!renamingSessionId) return;
    const session = findSessionById(sessions, renamingSessionId);
    if (!session) {
      setRenamingSessionId(null);
      setRenameValue('');
      setRenameValueDirty(false);
      return;
    }

    const trimmed = renameValue.trim();
    if (!trimmed) return;
    const isSystemUntitled = systemUntitledSessionIds[renamingSessionId] === true;
    if (isSystemUntitled && !renameValueDirty && trimmed === strings.untitledSessionTitle) {
      setRenamingSessionId(null);
      setRenameValue('');
      setRenameValueDirty(false);
      return;
    }

    setSessions((current) =>
      current.map((s) => s.id === renamingSessionId ? { ...s, title: trimmed, isSystemUntitled: false } : s),
    );
    setSystemUntitledSessionIds((current) => {
      if (!isSystemUntitled) {
        return current;
      }

      const { [renamingSessionId]: _ignored, ...rest } = current;
      return rest;
    });
    setRenamingSessionId(null);
    setRenameValue('');
    setRenameValueDirty(false);
  }

  function handleRenameCancel() {
    setRenamingSessionId(null);
    setRenameValue('');
    setRenameValueDirty(false);
  }

  function handleRenameKeyDown(event: ReactKeyboardEvent<HTMLInputElement>) {
    if (event.key === 'Enter') {
      event.preventDefault();
      handleRenameConfirm();
    } else if (event.key === 'Escape') {
      handleRenameCancel();
    }
  }

  function handleDeleteConfirm() {
    const targetId = deleteConfirmSessionId;
    if (!targetId) return;
    setDeleteConfirmSessionId(null);

    let nextActiveId: string | undefined;
    setSessions((current) => {
      const filtered = current.filter((s) => s.id !== targetId);
      if (activeSessionId === targetId) {
        nextActiveId = filtered.length > 0 ? filtered[0].id : '';
      }
      return filtered;
    });
    setSessionThreads((current) => {
      const { [targetId]: _, ...rest } = current;
      return rest;
    });
    setPendingConfirmations((current) => {
      const { [targetId]: _, ...rest } = current;
      return rest;
    });
    setPendingCommandSessions((current) => {
      const { [targetId]: _, ...rest } = current;
      return rest;
    });
    setSystemUntitledSessionIds((current) => {
      const { [targetId]: _ignored, ...rest } = current;
      return rest;
    });

    if (nextActiveId !== undefined) {
      setActiveSessionId(nextActiveId);
    }
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
      const savedSettings = normalizeSettings(await nativeBridge.saveSettings(draftSettings));
      setSettings(savedSettings);
      setDraftSettings(savedSettings);
      isSettingsDirtyRef.current = false;
      isSettingsOpenRef.current = false;
      shouldRestoreSettingsButtonFocusRef.current = true;
      setIsSettingsOpen(false);
      void nativeBridge.getHostContext()
        .then((hostContext) => {
          setUiLocale(hostContext.resolvedUiLocale);
        })
        .catch(() => {
          // best-effort refresh
        });
    } catch (error) {
      setSettingsSaveError(error instanceof Error ? error.message : strings.settingsSaveFailed);
    } finally {
      setIsSettingsSaving(false);
    }
  }

  async function handleComposerSend() {
    const trimmedValue = composerValue.trim();
    const sessionId = activeSession?.id ?? activeSessionId;
    if (!trimmedValue || !sessionId || isComposerDisabled) {
      return;
    }

    appendThreadMessage(sessionId, {
      id: createMessageId(),
      role: 'user',
      content: trimmedValue,
    });
    setComposerValue('');

    const command = parseExcelCommand(trimmedValue);
    if (command) {
      await dispatchExcelCommand(command, sessionId);
      return;
    }

    if (matchesDirectSkillInput(trimmedValue)) {
      await dispatchSkill({
        userInput: trimmedValue,
        confirmed: false,
      }, sessionId);
      return;
    }

    await dispatchAgent({
      userInput: trimmedValue,
      confirmed: false,
      sessionId,
    }, sessionId);
  }

  async function saveCurrentSessions() {
    const chatSessions = sessions.map((session) => {
      const thread = sessionThreads[session.id] ?? [];
      return {
        ...session,
        messages: threadToChatMessages(thread),
        updatedAtUtc: new Date().toISOString(),
      };
    });

    try {
      await nativeBridge.saveSessions({
        activeSessionId,
        sessions: chatSessions,
      });
    } catch {
      // best-effort save
    }
  }

  // Debounced auto-save when session threads change
  useEffect(() => {
    if (Object.keys(sessionThreads).length === 0) {
      return;
    }

    if (saveTimeoutRef.current) {
      clearTimeout(saveTimeoutRef.current);
    }

    saveTimeoutRef.current = setTimeout(() => {
      void saveCurrentSessions();
    }, 1000);

    return () => {
      if (saveTimeoutRef.current) {
        clearTimeout(saveTimeoutRef.current);
      }
    };
  }, [sessionThreads]);

  // Auto-rename session when first user message arrives
  useEffect(() => {
    if (!activeSession || systemUntitledSessionIds[activeSession.id] !== true) {
      return;
    }

    const thread = sessionThreads[activeSession.id] ?? [];
    const firstUserMessage = thread.find((m) => m.role === 'user');
    if (!firstUserMessage) {
      return;
    }

    const newTitle = firstUserMessage.content.length > 20
      ? firstUserMessage.content.slice(0, 20) + '...'
      : firstUserMessage.content;

    setSessions((current) =>
      current.map((s) =>
        s.id === activeSession.id ? { ...s, title: newTitle, isSystemUntitled: false } : s,
      ),
    );
    setSystemUntitledSessionIds((current) => {
      const { [activeSession.id]: _ignored, ...rest } = current;
      return rest;
    });
  }, [activeSession?.id, systemUntitledSessionIds, sessionThreads]);

  async function handlePendingConfirmationConfirm() {
    if (!activePendingConfirmation || !activeSession?.id) {
      return;
    }

    if (activePendingConfirmation.kind === 'excel' && activePendingConfirmation.command) {
      await dispatchExcelCommand({
        ...activePendingConfirmation.command,
        confirmed: true,
      }, activeSession.id);
      return;
    }

    if (activePendingConfirmation.kind === 'skill' && activePendingConfirmation.skillRequest) {
      await dispatchSkill({
        ...activePendingConfirmation.skillRequest,
        confirmed: true,
      }, activeSession.id);
      return;
    }

    if (activePendingConfirmation.kind === 'agent' && activePendingConfirmation.agentRequest && activePendingConfirmation.plan) {
      await dispatchAgent({
        ...activePendingConfirmation.agentRequest,
        confirmed: true,
        plan: activePendingConfirmation.plan,
      }, activeSession.id);
    }
  }

  function handlePendingConfirmationCancel() {
    if (!activeSession?.id || !activePendingConfirmation) {
      return;
    }

    setSessionPendingConfirmation(activeSession.id, null);
    appendThreadMessage(activeSession.id, {
      id: createMessageId(),
      role: 'system',
      content: activePendingConfirmation.kind === 'skill'
        ? strings.cancellationSkill
        : activePendingConfirmation.kind === 'agent'
          ? strings.cancellationPlan
          : strings.cancellationExcel,
    });
  }

  async function dispatchExcelCommand(command: ExcelCommand, sessionId: string) {
    setCommandPending(sessionId, true);

    try {
      const result = await nativeBridge.executeExcelCommand(command);
      if (result.selectionContext) {
        setSelectionContext(result.selectionContext);
      }

      if (result.requiresConfirmation && result.preview) {
        setSessionPendingConfirmation(sessionId, {
          kind: 'excel',
          command,
          preview: createLocalizedExcelPreview(command, result.preview, result.selectionContext ?? selectionContext, uiLocale),
        });
        appendThreadMessage(sessionId, {
          id: createMessageId(),
          role: 'assistant',
          content: getUiStrings(uiLocale).browserPreviewExcelConfirmMessage,
        });
        return;
      }

      setSessionPendingConfirmation(sessionId, null);
      appendThreadMessage(sessionId, createResultMessage(result));
    } catch (error) {
      appendThreadMessage(sessionId, {
        id: createMessageId(),
        role: 'assistant',
        content: strings.requestFailed(error instanceof Error ? error.message : strings.excelRequestFallback),
      });
    } finally {
      setCommandPending(sessionId, false);
    }
  }

  async function dispatchSkill(request: SkillRequestEnvelope, sessionId: string) {
    setCommandPending(sessionId, true);

    try {
      const result = await nativeBridge.runSkill(request);
      if (result.requiresConfirmation && result.preview) {
        setSessionPendingConfirmation(sessionId, {
          kind: 'skill',
          skillRequest: {
            userInput: request.userInput,
            skillName: result.skillName,
            confirmed: false,
            uploadPreview: result.uploadPreview,
          },
          preview: createLocalizedSkillPreview(result.preview, result.uploadPreview, uiLocale),
        });
        appendThreadMessage(sessionId, createLocalizedSkillPreviewMessage(result, uiLocale));
        return;
      }

      setSessionPendingConfirmation(sessionId, null);
      appendThreadMessage(sessionId, createSkillResultMessage(result));
    } catch (error) {
      appendThreadMessage(sessionId, {
        id: createMessageId(),
        role: 'assistant',
        content: strings.requestFailed(error instanceof Error ? error.message : strings.skillRequestFallback),
      });
    } finally {
      setCommandPending(sessionId, false);
    }
  }

  async function dispatchAgent(request: AgentRequestEnvelope, sessionId: string) {
    setCommandPending(sessionId, true);

    try {
      const threadMessages = sessionThreads[sessionId] ?? [];
      const result = await nativeBridge.runAgent({
        ...request,
        sessionId,
        conversationHistory: extractConversationHistory(threadMessages),
      });

      if (result.requiresConfirmation && result.planner?.mode === 'plan' && result.planner.plan) {
        setSessionPendingConfirmation(sessionId, {
          kind: 'agent',
          agentRequest: {
            userInput: request.userInput,
            confirmed: false,
            sessionId,
          },
          plan: result.planner.plan,
          preview: createPlanPreview(result, uiLocale),
        });
        appendThreadMessage(sessionId, {
          id: createMessageId(),
          role: 'assistant',
          content: result.message,
        });
        return;
      }

      setSessionPendingConfirmation(sessionId, null);
      appendThreadMessages(sessionId, createAgentResultMessages(result, uiLocale));
    } catch (error) {
      appendThreadMessage(sessionId, {
        id: createMessageId(),
        role: 'assistant',
        content: strings.requestFailed(error instanceof Error ? error.message : strings.agentRequestFallback),
      });
    } finally {
      setCommandPending(sessionId, false);
    }
  }

  function appendThreadMessage(sessionId: string, message: ThreadMessage) {
    setSessionThreads((current) => ({
      ...current,
      [sessionId]: [...(current[sessionId] ?? createInitialThreadMessages(findSessionById(sessions, sessionId), uiLocale)), message],
    }));
  }

  function appendThreadMessages(sessionId: string, messages: ThreadMessage[]) {
    if (messages.length === 0) {
      return;
    }

    setSessionThreads((current) => ({
      ...current,
      [sessionId]: [
        ...(current[sessionId] ?? createInitialThreadMessages(findSessionById(sessions, sessionId), uiLocale)),
        ...messages,
      ],
    }));
  }

  function setSessionPendingConfirmation(sessionId: string, value: PendingConfirmation | null) {
    setPendingConfirmations((current) => {
      if (value == null) {
        if (!(sessionId in current)) {
          return current;
        }

        const { [sessionId]: _ignored, ...rest } = current;
        return rest;
      }

      return {
        ...current,
        [sessionId]: value,
      };
    });
  }

  function setCommandPending(sessionId: string, isPending: boolean) {
    setPendingCommandSessions((current) => {
      if (!isPending) {
        if (!(sessionId in current)) {
          return current;
        }

        const { [sessionId]: _ignored, ...rest } = current;
        return rest;
      }

      return {
        ...current,
        [sessionId]: true,
      };
    });
  }

  return (
    <div className="app-shell">
      <main className="workspace">
        <header className="chat-header" aria-label={strings.chatHeaderLabel}>
          <div className="chat-header__leading">
            <button
              type="button"
              className="icon-button icon-button--ghost"
              aria-label={isSessionsDrawerOpen ? strings.closeSessionsDrawer : strings.openSessionsDrawer}
              ref={sessionsButtonRef}
              onClick={toggleSessionsDrawer}
            >
              <MenuIcon />
            </button>

            <h1 className="title">{activeSession ? getSessionDisplayTitle(activeSession, systemUntitledSessionIds, strings) : strings.appHeadingFallback}</h1>
          </div>

          <button
            type="button"
            className="icon-button icon-button--ghost"
            aria-label={strings.openSettings}
            ref={settingsButtonRef}
            onClick={openSettings}
          >
            <SettingsIcon />
          </button>
        </header>

        <section ref={threadRef} className="thread" aria-label={strings.messageThreadLabel}>
          {activeThreadMessages.map((message) => (
            <article key={message.id} className={`message message--${message.role}`}>
              <p>{message.content}</p>
              {message.table ? (
                <div className="message-table">
                  <div>{message.table.headers.join(' | ')}</div>
                  {message.table.rows.map((row, index) => (
                    <div key={`${message.id}-row-${index}`}>{row.join(' | ')}</div>
                  ))}
                </div>
              ) : null}
            </article>
          ))}
          {isCommandPending ? (
            <article className="message message--loading">
              <div className="loading-spinner" />
              <p>{strings.loadingThinking}</p>
            </article>
          ) : null}
        </section>

        <div className="composer-stack">
          {activePendingConfirmation ? (
            <ConfirmationCard
              preview={activePendingConfirmation.preview}
              isBusy={isCommandPending}
              ariaLabel={strings.confirmCardLabel}
              eyebrow={strings.confirmCardEyebrow}
              title={strings.confirmCardTitle}
              cancelLabel={strings.cancel}
              confirmLabel={strings.confirm}
              onConfirm={handlePendingConfirmationConfirm}
              onCancel={handlePendingConfirmationCancel}
            />
          ) : null}

          <footer className="composer" aria-label={strings.messageComposerLabel}>
            <textarea
              aria-label={strings.messageComposerLabel}
              placeholder={strings.messagePlaceholder}
              rows={3}
              value={composerValue}
              disabled={isComposerDisabled}
              onChange={(event) => setComposerValue(event.target.value)}
              onKeyDown={handleComposerKeyDown}
            />
            <div className="composer__divider" />
            <div className="composer__actions">
              <section className="selection-pill" aria-label={strings.selectionCapsuleLabel} role="status">
                {formatSelectionCapsule(selectionContext, uiLocale)}
              </section>

              <button type="button" className="send-button" disabled={isComposerDisabled} onClick={handleComposerSend}>
                {strings.send}
              </button>
            </div>
          </footer>
        </div>
      </main>

      {isSessionsDrawerOpen ? (
        <div className="drawer-backdrop" onClick={closeSessionsDrawer}>
          <aside
            className="session-drawer"
            aria-label={strings.sessionsDrawerLabel}
            onClick={(event) => event.stopPropagation()}
          >
            <div className="session-drawer__header">
              <button
                type="button"
                className="icon-button icon-button--ghost"
                aria-label={strings.closeSessionsDrawer}
                onClick={closeSessionsDrawer}
              >
                <MenuIcon />
              </button>
              <div className="sidebar__title">{strings.sessionsTitle}</div>
              <button
                type="button"
                className="session-drawer__new-chat"
                aria-label={strings.newSession}
                onClick={handleCreateNewSession}
              >
                <NewChatIcon />
              </button>
            </div>
            {sessions.length === 0 ? (
              <div className="sidebar__empty">{strings.noSessions}</div>
            ) : (
              <div className="sidebar__list">
                {sessions.map((session) => (
                  renamingSessionId === session.id ? (
                    <div key={session.id} className={`session-chip${session.id === activeSession?.id ? ' session-chip--active' : ''} session-chip--editing`}>
                      <input
                        className="session-chip__edit-input"
                        type="text"
                        value={renameValue}
                        onChange={(e) => {
                          setRenameValue(e.target.value);
                          setRenameValueDirty(true);
                        }}
                        onKeyDown={handleRenameKeyDown}
                        autoFocus
                      />
                      <div className="session-chip__edit-actions">
                        <button type="button" className="session-chip__action-btn" aria-label={strings.confirmRename} onClick={handleRenameConfirm}>
                          <CheckIcon />
                        </button>
                        <button type="button" className="session-chip__action-btn" aria-label={strings.cancelRename} onClick={handleRenameCancel}>
                          <CloseIcon />
                        </button>
                      </div>
                    </div>
                  ) : (
                    <div
                      key={session.id}
                      className={`session-chip${session.id === activeSession?.id ? ' session-chip--active' : ''}`}
                      onClick={() => handleSessionSelect(session.id)}
                      role="button"
                      tabIndex={0}
                      onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); handleSessionSelect(session.id); } }}
                    >
                      <span className="session-chip__title">{getSessionDisplayTitle(session, systemUntitledSessionIds, strings)}</span>
                      <div className="session-chip__actions">
                        <button type="button" className="session-chip__action-btn" aria-label={strings.renameSession} onClick={(e) => { e.stopPropagation(); handleRenameStart(session); }}>
                          <PencilIcon />
                        </button>
                        <button type="button" className="session-chip__action-btn" aria-label={strings.deleteSession} onClick={(e) => { e.stopPropagation(); setDeleteConfirmSessionId(session.id); }}>
                          <TrashIcon />
                        </button>
                      </div>
                    </div>
                  )
                ))}
              </div>
            )}
          </aside>
        </div>
      ) : null}

      {isSettingsOpen ? (
        <div className="settings-backdrop">
          <section
            ref={settingsDialogRef}
            className="settings-dialog"
            role="dialog"
            style={SETTINGS_DIALOG_VIEWPORT_STYLE}
            aria-modal="true"
            aria-label={strings.settingsDialogLabel}
            onKeyDown={handleSettingsDialogKeyDown}
          >
            <div className="settings-dialog__header">
              <div>
                <div className="eyebrow">{strings.settingsEyebrow}</div>
                <h2 className="settings-dialog__title">{strings.settingsTitle}</h2>
              </div>
              <button
                type="button"
                className="icon-button icon-button--ghost"
                aria-label={strings.close}
                onClick={closeSettings}
                disabled={isSettingsSaving}
              >
                <CloseIcon />
              </button>
            </div>

            {settingsLoadError ? <p className="settings-error" role="alert">{settingsLoadError}</p> : null}
            {settingsSaveError ? <p className="settings-error" role="alert">{settingsSaveError}</p> : null}

            <label className="settings-field">
              <span>{strings.apiKeyFieldLabel}</span>
              <div className="settings-field__input-wrapper">
                <input
                  ref={apiKeyInputRef}
                  aria-label={strings.apiKeyFieldLabel}
                  type={showApiKey ? 'text' : 'password'}
                  value={draftSettings.apiKey}
                  disabled={isSettingsSaving}
                  onChange={(event) => updateDraftSettings({ apiKey: event.target.value })}
                />
                <button
                  type="button"
                  className="icon-button--ghost settings-field__toggle"
                  aria-label={showApiKey ? strings.hideApiKey : strings.showApiKey}
                  onClick={() => setShowApiKey((v) => !v)}
                  disabled={isSettingsSaving}
                >
                  {showApiKey ? <EyeOffIcon /> : <EyeIcon />}
                </button>
              </div>
            </label>

            <label className="settings-field">
              <span>{strings.baseUrlFieldLabel}</span>
              <input
                aria-label={strings.baseUrlFieldLabel}
                type="text"
                value={draftSettings.baseUrl}
                disabled={isSettingsSaving}
                onChange={(event) => updateDraftSettings({ baseUrl: event.target.value })}
              />
            </label>

            <label className="settings-field">
              <span>{strings.businessBaseUrlFieldLabel}</span>
              <input
                aria-label={strings.businessBaseUrlFieldLabel}
                type="text"
                value={draftSettings.businessBaseUrl}
                disabled={isSettingsSaving}
                onChange={(event) => updateDraftSettings({ businessBaseUrl: event.target.value })}
              />
            </label>

            <label className="settings-field">
              <span>{strings.modelFieldLabel}</span>
              <input
                aria-label={strings.modelFieldLabel}
                type="text"
                value={draftSettings.model}
                disabled={isSettingsSaving}
                onChange={(event) => updateDraftSettings({ model: event.target.value })}
              />
            </label>

            <label className="settings-field">
              <span>{strings.ssoUrlFieldLabel}</span>
              <input
                aria-label={strings.ssoUrlFieldLabel}
                type="text"
                value={draftSettings.ssoUrl}
                disabled={isSettingsSaving || isLoggingIn}
                onChange={(event) => updateDraftSettings({ ssoUrl: event.target.value })}
              />
            </label>

            <label className="settings-field">
              <span>{strings.loginSuccessPath}</span>
              <input
                aria-label={strings.loginSuccessPath}
                type="text"
                value={draftSettings.ssoLoginSuccessPath}
                disabled={isSettingsSaving || isLoggingIn}
                onChange={(event) => updateDraftSettings({ ssoLoginSuccessPath: event.target.value })}
              />
            </label>

            <div className="login-status">
              {loginStatus?.isLoggedIn ? (
                <>
                  <span className="login-badge login-badge--active">{strings.loggedIn}</span>
                  <button type="button" className="ghost-button" onClick={handleLogout} disabled={isSettingsSaving}>
                    {strings.logout}
                  </button>
                </>
              ) : (
                <>
                  <span className="login-badge">{strings.loggedOut}</span>
                  <button
                    type="button"
                    className="ghost-button"
                    onClick={handleLogin}
                    disabled={isSettingsSaving || isLoggingIn || !draftSettings.ssoUrl.trim()}
                  >
                    {isLoggingIn ? strings.loginInProgress : strings.login}
                  </button>
                </>
              )}
            </div>

            {loginError ? <p className="settings-error" role="alert">{loginError}</p> : null}

            <div className="settings-actions">
              <button type="button" className="ghost-button" onClick={closeSettings} disabled={isSettingsSaving}>
                {strings.cancel}
              </button>
              <button
                type="button"
                className="send-button"
                onClick={handleSettingsSave}
                disabled={isSettingsLoading || isSettingsSaving || Boolean(settingsLoadError)}
              >
                {strings.save}
              </button>
            </div>
            <div className="settings-version">{bridgeStatusText}</div>
          </section>
        </div>
      ) : null}

      {deleteConfirmSessionId ? (
        <div className="delete-dialog-backdrop">
          <div className="delete-dialog">
            <h2 className="delete-dialog__title">{strings.deleteSessionDialogTitle}</h2>
            <p className="delete-dialog__message">
              {strings.deleteSessionPrompt(getSessionDisplayTitle(sessions.find((s) => s.id === deleteConfirmSessionId), systemUntitledSessionIds, strings))}
            </p>
            <div className="delete-dialog__actions">
              <button type="button" className="ghost-button" onClick={() => setDeleteConfirmSessionId(null)}>{strings.cancel}</button>
              <button type="button" className="send-button" onClick={handleDeleteConfirm}>{strings.deleteSession}</button>
            </div>
          </div>
        </div>
      ) : null}
    </div>
  );
}

const MAX_CONVERSATION_HISTORY_TURNS = 10;

function extractConversationHistory(messages: ThreadMessage[]): ConversationTurn[] {
  const eligible = messages.filter((m) => m.role === 'user' || m.role === 'assistant');
  const clipped = eligible.slice(-MAX_CONVERSATION_HISTORY_TURNS * 2);
  return clipped.map((m) => ({
    role: m.role,
    content: m.content,
  }));
}

function threadToChatMessages(messages: ThreadMessage[]): Array<{ id: string; role: string; content: string; createdAtUtc: string }> {
  return messages
    .filter((m) => (m.role === 'user' || m.role === 'assistant') && m.id !== 'welcome-message')
    .map((m) => ({
      id: m.id,
      role: m.role,
      content: m.content,
      createdAtUtc: new Date().toISOString(),
    }));
}

function relocalizeWelcomeMessages(currentThreads: Record<string, ThreadMessage[]>, locale: UiLocale) {
  const welcomeMessage = getUiStrings(locale).welcomeMessage;
  let hasChanges = false;
  const nextThreads: Record<string, ThreadMessage[]> = {};

  Object.entries(currentThreads).forEach(([sessionId, thread]) => {
    let threadChanged = false;
    const nextThread = thread.map((message) => {
      if (message.id !== 'welcome-message' || message.content === welcomeMessage) {
        return message;
      }

      threadChanged = true;
      return {
        ...message,
        content: welcomeMessage,
      };
    });

    nextThreads[sessionId] = threadChanged ? nextThread : thread;
    hasChanges = hasChanges || threadChanged;
  });

  return hasChanges ? nextThreads : currentThreads;
}

function createInitialThreadMessages(session: ChatSession | undefined, locale: UiLocale): ThreadMessage[] {
  const persistedMessages = session?.messages ?? [];
  if (persistedMessages.length > 0) {
    return persistedMessages.map((message) => ({
      id: message.id,
      role: message.role === 'user' ? 'user' : 'assistant',
      content: message.content,
    }));
  }

  return [
    {
      id: 'welcome-message',
      role: 'assistant',
      content: getUiStrings(locale).welcomeMessage,
    },
  ];
}

function hydrateSessionThreads(
  currentThreads: Record<string, ThreadMessage[]>,
  sessions: ChatSession[],
  locale: UiLocale,
): Record<string, ThreadMessage[]> {
  const nextThreads: Record<string, ThreadMessage[]> = {};

  sessions.forEach((session) => {
    nextThreads[session.id] = currentThreads[session.id] ?? createInitialThreadMessages(session, locale);
  });

  return nextThreads;
}

function findSessionById(sessions: ChatSession[], sessionId: string): ChatSession | undefined {
  return sessions.find((session) => session.id === sessionId);
}

function deriveSystemUntitledSessionIds(sessions: ChatSession[]) {
  return sessions.reduce<Record<string, true>>((accumulator, session) => {
    if (session.isSystemUntitled === true) {
      accumulator[session.id] = true;
    }

    return accumulator;
  }, {});
}

function getSessionDisplayTitle(
  session: ChatSession | undefined,
  systemUntitledSessionIds: Record<string, true>,
  strings: ReturnType<typeof getUiStrings>,
) {
  if (!session) {
    return '';
  }

  return systemUntitledSessionIds[session.id] === true ? strings.untitledSessionTitle : session.title;
}

function createResultMessage(result: ExcelCommandResult): ThreadMessage {
  return {
    id: createMessageId(),
    role: 'assistant',
    content: result.message,
    table: result.table,
  };
}

function createSkillResultMessage(result: SkillResult): ThreadMessage {
  return {
    id: createMessageId(),
    role: 'assistant',
    content: result.message,
    table: createTableFromUploadPreview(result.uploadPreview),
  };
}

function createAgentResultMessages(result: AgentResult, locale: UiLocale): ThreadMessage[] {
  const strings = getUiStrings(locale);
  const messages: ThreadMessage[] = [
    {
      id: createMessageId(),
      role: 'assistant',
      content: result.message,
    },
  ];

  if (result.journal) {
    result.journal.steps.forEach((step) => {
      messages.push({
        id: createMessageId(),
        role: 'system',
        content: `${strings.formatJournalStatus(step.status)} · ${step.title}${step.errorMessage ? ` · ${step.errorMessage}` : ''}`.trim(),
      });
    });
  }

  return messages;
}

function createPlanPreview(result: AgentResult, locale: UiLocale): ExcelCommandPreview {
  const strings = getUiStrings(locale);
  const plan = result.planner?.plan;
  return {
    title: strings.planPreviewTitle,
    summary: plan?.summary ?? result.message,
    details: plan?.steps.map((step) => formatPlanStep(step, locale)) ?? [],
  };
}

function createLocalizedExcelPreview(
  command: ExcelCommand,
  preview: ExcelCommandPreview,
  selectionContext: SelectionContext | null,
  locale: UiLocale,
): ExcelCommandPreview {
  const strings = getUiStrings(locale);
  const workbookName = selectionContext?.workbookName?.trim() ?? '';
  const localizedWorkbookDetails = localizeWorkbookDetails(preview.details, workbookName, locale);

  switch (command.commandType) {
    case 'excel.addWorksheet':
      return {
        title: strings.excelAddWorksheetPreviewTitle,
        summary: strings.formatExcelAddWorksheetPreviewSummary(String(command.newSheetName ?? '').trim()),
        details: localizedWorkbookDetails,
      };
    case 'excel.renameWorksheet':
      return {
        title: strings.excelRenameWorksheetPreviewTitle,
        summary: strings.formatExcelRenameWorksheetPreviewSummary(
          String(command.sheetName ?? '').trim(),
          String(command.newSheetName ?? '').trim(),
        ),
        details: localizedWorkbookDetails,
      };
    case 'excel.deleteWorksheet':
      return {
        title: strings.excelDeleteWorksheetPreviewTitle,
        summary: strings.formatExcelDeleteWorksheetPreviewSummary(String(command.sheetName ?? '').trim()),
        details: localizedWorkbookDetails,
      };
    case 'excel.writeRange':
      return {
        title: strings.excelWriteRangePreviewTitle,
        summary: strings.formatExcelWriteRangePreviewSummary(
          command.values?.length ?? 0,
          command.values?.[0]?.length ?? 0,
          String(command.targetAddress ?? '').trim(),
        ),
        details: preview.details,
      };
    default:
      return preview;
  }
}

function createLocalizedSkillPreview(preview: ExcelCommandPreview, uploadPreview: UploadPreview | undefined, locale: UiLocale) {
  if (!uploadPreview) {
    return preview;
  }

  const strings = getUiStrings(locale);
  return {
    title: strings.uploadPreviewTitle,
    summary: strings.formatUploadPreviewSummary(uploadPreview.records.length, uploadPreview.projectName),
    details: localizeUploadPreviewDetails(preview.details, uploadPreview, locale),
  };
}

function createLocalizedSkillPreviewMessage(result: SkillResult, locale: UiLocale): ThreadMessage {
  const projectName = result.uploadPreview?.projectName?.trim();
  return {
    id: createMessageId(),
    role: 'assistant',
    content: projectName
      ? getUiStrings(locale).browserPreviewUploadReviewMessage(projectName)
      : result.message,
    table: createTableFromUploadPreview(result.uploadPreview),
  };
}

function readHostContextWithTimeout(hostContextPromise: Promise<HostContext | null>, timeoutMs: number = 1500): Promise<HostContext | null> {
  let timeoutId: ReturnType<typeof setTimeout> | null = null;

  const timeoutPromise = new Promise<null>((resolve) => {
    timeoutId = setTimeout(() => resolve(null), timeoutMs);
  });

  return Promise.race<HostContext | null>([
    hostContextPromise,
    timeoutPromise,
  ])
    .catch(() => null)
    .finally(() => {
      if (timeoutId !== null) {
        clearTimeout(timeoutId);
      }
    });
}

function createTableFromUploadPreview(preview?: UploadPreview): ExcelTableData | undefined {
  if (!preview) {
    return undefined;
  }

  return {
    sheetName: preview.sheetName,
    address: preview.address,
    headers: preview.headers,
    rows: preview.rows,
  };
}

function localizeWorkbookDetails(hostDetails: string[] | undefined, workbookName: string, locale: UiLocale) {
  if (!workbookName) {
    return hostDetails ?? [];
  }

  const localizedDetail = getUiStrings(locale).formatWorkbookDetail(workbookName);
  const knownDetails = [
    localizedDetail,
    getUiStrings('zh').formatWorkbookDetail(workbookName),
    getUiStrings('en').formatWorkbookDetail(workbookName),
  ];

  return dedupeDetails([
    localizedDetail,
    ...(hostDetails ?? []).filter((detail) => !matchesKnownDetail(detail, knownDetails)),
  ]);
}

function localizeUploadPreviewDetails(hostDetails: string[] | undefined, uploadPreview: UploadPreview, locale: UiLocale) {
  const localizedStrings = getUiStrings(locale);
  const localizedSourceDetail = localizedStrings.formatUploadPreviewSourceDetail(uploadPreview.sheetName, uploadPreview.address);
  const localizedFieldsDetail = localizedStrings.formatUploadPreviewFieldsDetail(uploadPreview.headers);
  const knownSourceDetails = [
    localizedSourceDetail,
    getUiStrings('zh').formatUploadPreviewSourceDetail(uploadPreview.sheetName, uploadPreview.address),
    getUiStrings('en').formatUploadPreviewSourceDetail(uploadPreview.sheetName, uploadPreview.address),
  ];
  const knownFieldsDetails = [
    localizedFieldsDetail,
    getUiStrings('zh').formatUploadPreviewFieldsDetail(uploadPreview.headers),
    getUiStrings('en').formatUploadPreviewFieldsDetail(uploadPreview.headers),
  ];

  return dedupeDetails([
    localizedSourceDetail,
    localizedFieldsDetail,
    ...(hostDetails ?? []).filter((detail) => (
      !matchesKnownDetail(detail, knownSourceDetails) &&
      !matchesKnownDetail(detail, knownFieldsDetails)
    )),
  ]);
}

function matchesKnownDetail(detail: string, knownDetails: string[]) {
  const normalizedDetail = normalizeDetail(detail);
  return normalizedDetail.length > 0 && knownDetails.some((candidate) => normalizeDetail(candidate) === normalizedDetail);
}

function normalizeDetail(detail: string) {
  return detail.trim().replace(/\s+/g, ' ').toLowerCase();
}

function dedupeDetails(details: string[]) {
  const seen = new Set<string>();
  return details.filter((detail) => {
    const normalized = normalizeDetail(detail);
    if (!normalized || seen.has(normalized)) {
      return false;
    }

    seen.add(normalized);
    return true;
  });
}

function formatBridgeStatus(status: BridgeStatusState, strings: ReturnType<typeof getUiStrings>) {
  switch (status.kind) {
    case 'connected':
      return strings.bridgeConnected(status.host, status.version);
    case 'unavailable':
      return strings.bridgeUnavailable(status.errorMessage);
    default:
      return strings.bridgeConnecting;
  }
}

function formatPlanStep(step: AgentPlan['steps'][number], locale: UiLocale) {
  const strings = getUiStrings(locale);
  switch (step.type) {
    case 'excel.addWorksheet':
      return strings.formatPlanStepAddWorksheet(String(step.args?.newSheetName ?? '').trim());
    case 'excel.writeRange':
      return strings.formatPlanStepWriteRange(String(step.args?.targetAddress ?? '').trim());
    case 'excel.renameWorksheet':
      return strings.formatPlanStepRenameWorksheet(String(step.args?.sheetName ?? '').trim(), String(step.args?.newSheetName ?? '').trim());
    case 'excel.deleteWorksheet':
      return strings.formatPlanStepDeleteWorksheet(String(step.args?.sheetName ?? '').trim());
    case 'skill.upload_data':
      return strings.formatPlanStepUploadData;
    default:
      return step.type;
  }
}

function matchesDirectSkillInput(input: string) {
  const trimmedInput = input.trim();
  return (
    trimmedInput.startsWith('/upload_data') ||
    trimmedInput.includes('\u4E0A\u4F20\u5230') ||
    /\bupload\b.+\bto\s+.+$/i.test(trimmedInput)
  );
}

function parseExcelCommand(input: string): ExcelCommand | null {
  const trimmed = input.trim();
  if (/^\/?(read[_ ]selection)$/i.test(trimmed)) {
    return {
      commandType: 'excel.readSelectionTable',
      confirmed: false,
    };
  }

  const addSheetMatch = trimmed.match(/^\/?(?:add[_ ]sheet)\s+(.+)$/i);
  if (addSheetMatch) {
    return {
      commandType: 'excel.addWorksheet',
      newSheetName: addSheetMatch[1].trim(),
      confirmed: false,
    };
  }

  const renameSheetMatch = trimmed.match(/^\/?(?:rename[_ ]sheet)\s+(.+?)\s+(?:to|=>)\s+(.+)$/i);
  if (renameSheetMatch) {
    return {
      commandType: 'excel.renameWorksheet',
      sheetName: renameSheetMatch[1].trim(),
      newSheetName: renameSheetMatch[2].trim(),
      confirmed: false,
    };
  }

  const deleteSheetMatch = trimmed.match(/^\/?(?:delete[_ ]sheet)\s+(.+)$/i);
  if (deleteSheetMatch) {
    return {
      commandType: 'excel.deleteWorksheet',
      sheetName: deleteSheetMatch[1].trim(),
      confirmed: false,
    };
  }

  const writeRangeMatch = trimmed.match(/^\/?(?:write[_ ]range)\s+([^\s=]+)\s*=\s*(.+)$/i);
  if (writeRangeMatch) {
    return {
      commandType: 'excel.writeRange',
      targetAddress: writeRangeMatch[1].trim(),
      values: parseWriteRows(writeRangeMatch[2]),
      confirmed: false,
    };
  }

  return null;
}

function parseWriteRows(value: string): string[][] {
  return value
    .split('|')
    .map((row) => row.split(',').map((cell) => cell.trim()))
    .filter((row) => row.some((cell) => cell.length > 0));
}

function createMessageId() {
  if (globalThis.crypto?.randomUUID) {
    return globalThis.crypto.randomUUID();
  }

  return `msg-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function normalizeSettings(settings: Partial<AppSettings> | null | undefined): AppSettings {
  return {
    ...DEFAULT_SETTINGS,
    ...settings,
    uiLanguageOverride: settings?.uiLanguageOverride ?? DEFAULT_SETTINGS.uiLanguageOverride,
  };
}

function formatSelectionCapsule(selectionContext: SelectionContext | null, locale: UiLocale) {
  if (!selectionContext?.hasSelection || !selectionContext.sheetName || !selectionContext.address) {
    return getUiStrings(locale).noSelection;
  }

  return `${selectionContext.sheetName} · ${selectionContext.address}`;
}

function MenuIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true" className="icon-svg">
      <path d="M4 7h16M4 12h16M4 17h16" />
    </svg>
  );
}

function SettingsIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true" className="icon-svg">
      <path d="M6 7h12M6 12h12M6 17h12M9 7v0M15 12v0M11 17v0" />
      <circle cx="9" cy="7" r="2" />
      <circle cx="15" cy="12" r="2" />
      <circle cx="11" cy="17" r="2" />
    </svg>
  );
}

function NewChatIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true" className="icon-svg" width="18" height="18">
      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
      <polyline points="14 2 14 8 20 8" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
      <line x1="12" y1="18" x2="12" y2="12" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
      <line x1="9" y1="15" x2="15" y2="15" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
    </svg>
  );
}

function PencilIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true" className="icon-svg" width="14" height="14">
      <path d="M17 3a2.828 2.828 0 1 1 4 4L7.5 17.5 2 22l4.5-1.5L21 7.5z" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
      <path d="M15 5l4 4" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  );
}

function TrashIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true" className="icon-svg" width="14" height="14">
      <polyline points="3 6 5 6 21 6" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
      <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  );
}

function CheckIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true" className="icon-svg" width="14" height="14">
      <polyline points="20 6 9 17 4 12" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  );
}

function CloseIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true" className="icon-svg">
      <path d="M6 6l12 12M18 6L6 18" />
    </svg>
  );
}

function EyeIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true" className="icon-svg">
      <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z" />
      <circle cx="12" cy="12" r="3" />
    </svg>
  );
}

function EyeOffIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true" className="icon-svg">
      <path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19m-6.72-1.07a3 3 0 1 1-4.24-4.24" />
      <line x1="1" y1="1" x2="23" y2="23" />
    </svg>
  );
}

export default App;
