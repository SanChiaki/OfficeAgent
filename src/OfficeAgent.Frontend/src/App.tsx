import { useEffect, useRef, useState, type KeyboardEvent as ReactKeyboardEvent } from 'react';
import { nativeBridge } from './bridge/nativeBridge';
import { ConfirmationCard } from './components/ConfirmationCard';
import type {
  AgentPlan,
  AgentRequestEnvelope,
  AgentResult,
  AppSettings,
  ChatSession,
  ExcelCommand,
  ExcelCommandPreview,
  ExcelCommandResult,
  ExcelTableData,
  SelectionContext,
  SkillRequestEnvelope,
  SkillResult,
  UploadPreview,
} from './types/bridge';

const DEFAULT_SETTINGS: AppSettings = {
  apiKey: '',
  baseUrl: 'https://api.example.com',
  model: 'gpt-5-mini',
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

export function App() {
  const [bridgeStatus, setBridgeStatus] = useState('正在连接宿主...');
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
  const [composerValue, setComposerValue] = useState('');
  const [sessionThreads, setSessionThreads] = useState<Record<string, ThreadMessage[]>>({});
  const [pendingConfirmations, setPendingConfirmations] = useState<Record<string, PendingConfirmation>>({});
  const [pendingCommandSessions, setPendingCommandSessions] = useState<Record<string, boolean>>({});
  const sessionsButtonRef = useRef<HTMLButtonElement | null>(null);
  const settingsButtonRef = useRef<HTMLButtonElement | null>(null);
  const settingsDialogRef = useRef<HTMLElement | null>(null);
  const apiKeyInputRef = useRef<HTMLInputElement | null>(null);
  const threadRef = useRef<HTMLElement | null>(null);
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

        setBridgeStatus(`已连接 ${result.host} (${result.version})`);
      })
      .catch((error: Error) => {
        if (!isActive) {
          return;
        }

        setBridgeStatus(`宿主不可用: ${error.message}`);
      });

    nativeBridge
      .getSessions()
      .then((result) => {
        if (!isActive) {
          return;
        }

        setSessions(result.sessions);
        setSessionThreads((current) => hydrateSessionThreads(current, result.sessions));
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
        setSettingsLoadError('无法从宿主加载设置。');
        setSettingsSaveError('');
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

    return () => {
      isActive = false;
      unsubscribeSelectionContext();
    };
  }, []);

  const activeSession = sessions.find((session) => session.id === activeSessionId) ?? sessions[0];
  const activeThreadMessages = activeSession
    ? sessionThreads[activeSession.id] ?? createInitialThreadMessages(activeSession)
    : createInitialThreadMessages();
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
    setActiveSessionId(sessionId);
    setIsSessionsDrawerOpen(false);
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
      setSettingsSaveError(error instanceof Error ? error.message : '保存设置失败。');
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
        ? '已取消待处理的上传操作。'
        : activePendingConfirmation.kind === 'agent'
          ? '已取消待执行的计划。'
          : '已取消待处理的 Excel 操作。',
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
          preview: result.preview,
        });
        appendThreadMessage(sessionId, {
          id: createMessageId(),
          role: 'assistant',
          content: result.message,
        });
        return;
      }

      setSessionPendingConfirmation(sessionId, null);
      appendThreadMessage(sessionId, createResultMessage(result));
    } catch (error) {
      appendThreadMessage(sessionId, {
        id: createMessageId(),
        role: 'assistant',
        content: error instanceof Error ? error.message : 'Excel 命令执行失败。',
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
          preview: result.preview,
        });
        appendThreadMessage(sessionId, createSkillResultMessage(result));
        return;
      }

      setSessionPendingConfirmation(sessionId, null);
      appendThreadMessage(sessionId, createSkillResultMessage(result));
    } catch (error) {
      appendThreadMessage(sessionId, {
        id: createMessageId(),
        role: 'assistant',
        content: error instanceof Error ? error.message : 'Skill 执行失败。',
      });
    } finally {
      setCommandPending(sessionId, false);
    }
  }

  async function dispatchAgent(request: AgentRequestEnvelope, sessionId: string) {
    setCommandPending(sessionId, true);

    try {
      const result = await nativeBridge.runAgent({
        ...request,
        sessionId,
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
          preview: createPlanPreview(result),
        });
        appendThreadMessage(sessionId, {
          id: createMessageId(),
          role: 'assistant',
          content: result.message,
        });
        return;
      }

      setSessionPendingConfirmation(sessionId, null);
      appendThreadMessages(sessionId, createAgentResultMessages(result));
    } catch (error) {
      appendThreadMessage(sessionId, {
        id: createMessageId(),
        role: 'assistant',
        content: error instanceof Error ? error.message : 'Agent 执行失败。',
      });
    } finally {
      setCommandPending(sessionId, false);
    }
  }

  function appendThreadMessage(sessionId: string, message: ThreadMessage) {
    setSessionThreads((current) => ({
      ...current,
      [sessionId]: [...(current[sessionId] ?? createInitialThreadMessages(findSessionById(sessions, sessionId))), message],
    }));
  }

  function appendThreadMessages(sessionId: string, messages: ThreadMessage[]) {
    if (messages.length === 0) {
      return;
    }

    setSessionThreads((current) => ({
      ...current,
      [sessionId]: [
        ...(current[sessionId] ?? createInitialThreadMessages(findSessionById(sessions, sessionId))),
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
        <header className="chat-header" aria-label="Chat header">
          <div className="chat-header__leading">
            <button
              type="button"
              className="icon-button icon-button--ghost"
              aria-label={isSessionsDrawerOpen ? '关闭会话列表' : '打开会话列表'}
              ref={sessionsButtonRef}
              onClick={toggleSessionsDrawer}
            >
              <MenuIcon />
            </button>

            <div>
              <div className="eyebrow">Office Agent</div>
              <h1 className="title">{activeSession?.title ?? 'Office Agent 任务窗格'}</h1>
              <div className="subtitle">{settings?.baseUrl ?? '设置尚未加载'}</div>
              <div className="status-line">{bridgeStatus}</div>
            </div>
          </div>

          <button
            type="button"
            className="icon-button icon-button--ghost"
            aria-label="打开设置"
            ref={settingsButtonRef}
            onClick={openSettings}
          >
            <SettingsIcon />
          </button>
        </header>

        <section ref={threadRef} className="thread" aria-label="Message thread">
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
              <p>{'\u6B63\u5728\u601D\u8003\u2026'}</p>
            </article>
          ) : null}
        </section>

        <div className="composer-stack">
          {activePendingConfirmation ? (
            <ConfirmationCard
              preview={activePendingConfirmation.preview}
              isBusy={isCommandPending}
              onConfirm={handlePendingConfirmationConfirm}
              onCancel={handlePendingConfirmationCancel}
            />
          ) : null}

          <footer className="composer" aria-label="Message composer">
            <textarea
              aria-label="Message composer"
              placeholder="输入消息..."
              rows={3}
              value={composerValue}
              disabled={isComposerDisabled}
              onChange={(event) => setComposerValue(event.target.value)}
              onKeyDown={handleComposerKeyDown}
            />
            <div className="composer__divider" />
            <div className="composer__actions">
              <section className="selection-pill" aria-label="Selection capsule" role="status">
                {formatSelectionCapsule(selectionContext)}
              </section>

              <button type="button" className="send-button" disabled={isComposerDisabled} onClick={handleComposerSend}>
                发送
              </button>
            </div>
          </footer>
        </div>
      </main>

      {isSessionsDrawerOpen ? (
        <div className="drawer-backdrop" onClick={closeSessionsDrawer}>
          <aside
            className="session-drawer"
            aria-label="Sessions drawer"
            onClick={(event) => event.stopPropagation()}
          >
            <div className="session-drawer__header">
              <button
                type="button"
                className="icon-button icon-button--ghost"
                aria-label="关闭会话列表"
                onClick={closeSessionsDrawer}
              >
                <MenuIcon />
              </button>
              <div className="sidebar__title">会话</div>
            </div>
            {sessions.length === 0 ? (
              <div className="sidebar__empty">暂无会话</div>
            ) : (
              <div className="sidebar__list">
                {sessions.map((session) => (
                  <button
                    key={session.id}
                    type="button"
                    className={`session-chip${session.id === activeSession?.id ? ' session-chip--active' : ''}`}
                    onClick={() => handleSessionSelect(session.id)}
                  >
                    {session.title}
                  </button>
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
            aria-modal="true"
            aria-label="设置对话框"
            onKeyDown={handleSettingsDialogKeyDown}
          >
            <div className="settings-dialog__header">
              <div>
                <div className="eyebrow">配置</div>
                <h2 className="settings-dialog__title">设置</h2>
              </div>
              <button
                type="button"
                className="icon-button icon-button--ghost"
                aria-label="关闭"
                onClick={closeSettings}
                disabled={isSettingsSaving}
              >
                <CloseIcon />
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
                取消
              </button>
              <button
                type="button"
                className="send-button"
                onClick={handleSettingsSave}
                disabled={isSettingsLoading || isSettingsSaving || Boolean(settingsLoadError)}
              >
                保存
              </button>
            </div>
          </section>
        </div>
      ) : null}
    </div>
  );
}

function createInitialThreadMessages(session?: ChatSession): ThreadMessage[] {
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
      content: '欢迎使用 Office Agent。你可以直接使用 Excel 命令，完整的 Agent 路由功能正在接入中。',
    },
  ];
}

function hydrateSessionThreads(
  currentThreads: Record<string, ThreadMessage[]>,
  sessions: ChatSession[],
): Record<string, ThreadMessage[]> {
  const nextThreads: Record<string, ThreadMessage[]> = {};

  sessions.forEach((session) => {
    nextThreads[session.id] = currentThreads[session.id] ?? createInitialThreadMessages(session);
  });

  return nextThreads;
}

function findSessionById(sessions: ChatSession[], sessionId: string): ChatSession | undefined {
  return sessions.find((session) => session.id === sessionId);
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

function createAgentResultMessages(result: AgentResult): ThreadMessage[] {
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
        content: `${step.status} · ${step.title}${step.errorMessage ? ` · ${step.errorMessage}` : ''}`.trim(),
      });
    });
  }

  return messages;
}

function createPlanPreview(result: AgentResult): ExcelCommandPreview {
  const plan = result.planner?.plan;
  return {
    title: '执行计划',
    summary: plan?.summary ?? result.message,
    details: plan?.steps.map(formatPlanStep) ?? [],
  };
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

function formatPlanStep(step: AgentPlan['steps'][number]) {
  switch (step.type) {
    case 'excel.addWorksheet':
      return `新增工作表 ${String(step.args?.newSheetName ?? '').trim()}`.trim();
    case 'excel.writeRange':
      return `写入范围 ${String(step.args?.targetAddress ?? '').trim()}`.trim();
    case 'excel.renameWorksheet':
      return `重命名工作表 ${String(step.args?.sheetName ?? '').trim()} 为 ${String(step.args?.newSheetName ?? '').trim()}`.trim();
    case 'excel.deleteWorksheet':
      return `删除工作表 ${String(step.args?.sheetName ?? '').trim()}`.trim();
    case 'skill.upload_data':
      return '上传所选数据';
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

function formatSelectionCapsule(selectionContext: SelectionContext | null) {
  if (!selectionContext?.hasSelection || !selectionContext.sheetName || !selectionContext.address) {
    return '未选中';
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

function CloseIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true" className="icon-svg">
      <path d="M6 6l12 12M18 6L6 18" />
    </svg>
  );
}

export default App;
