import { useEffect, useMemo, useRef, useState } from "react";
import { Composer } from "./components/Composer";
import { ConfirmationCard } from "./components/ConfirmationCard";
import { MessageThread } from "./components/MessageThread";
import { SettingsDialog } from "./components/SettingsDialog";
import { SelectionBadge } from "./components/SelectionBadge";
import { SessionSidebar } from "./components/SessionSidebar";
import { decideRoute } from "./agent/agentOrchestrator";
import { classifyAction, createExcelAdapter } from "./excel/excelAdapter";
import { subscribeToSelectionChanges } from "./excel/selectionContextService";
import { createSessionStore } from "./state/sessionStore";
import { createSettingsStore } from "./state/settingsStore";
import type { ExcelAction } from "./excel/excelAdapter";
import type { SettingsState } from "./state/settingsStore";
import type { ChatMessage, SelectionContext } from "./types";

type SessionStore = ReturnType<typeof createSessionStore>;
type ExcelAdapter = ReturnType<typeof createExcelAdapter>;

export interface PendingExcelConfirmation {
  requestId: number;
  sessionId: string;
  action: ExcelAction;
  isExecuting: boolean;
  error: string | null;
}

export interface AppProps {
  sessionStoreFactory?: () => SessionStore;
  excelAdapterFactory?: () => ExcelAdapter;
  initialPendingConfirmation?: PendingExcelConfirmation | null;
}

const initialMessages: ChatMessage[] = [
  {
    id: "assistant-welcome",
    role: "assistant",
    content: "\u4f60\u597d\uff0c\u6211\u662f OfficeAgent\u3002",
  },
];

function formatExecutionError(error: unknown) {
  if (error instanceof Error && error.message) {
    return `\u6267\u884c\u5931\u8d25\uff1a${error.message}`;
  }

  return "\u6267\u884c\u5931\u8d25\uff0c\u8bf7\u91cd\u8bd5";
}

export default function App({
  sessionStoreFactory = createSessionStore,
  excelAdapterFactory = createExcelAdapter,
  initialPendingConfirmation = null,
}: AppProps = {}) {
  const sessionStore = useMemo(() => sessionStoreFactory(), [sessionStoreFactory]);
  const settingsStore = useMemo(() => createSettingsStore(), []);
  const excelAdapter = useMemo(() => excelAdapterFactory(), [excelAdapterFactory]);
  const [{ sessions, activeSessionId }, setSessionState] = useState(sessionStore.getState());
  const [settings, setSettings] = useState(settingsStore.load());
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [draft, setDraft] = useState("");
  const [messages, setMessages] = useState<ChatMessage[]>(initialMessages);
  const [selection, setSelection] = useState<SelectionContext | null>(null);
  const [pendingConfirmation, setPendingConfirmation] = useState<PendingExcelConfirmation | null>(
    initialPendingConfirmation,
  );
  const pendingConfirmationRef = useRef<PendingExcelConfirmation | null>(initialPendingConfirmation);
  const nextRequestIdRef = useRef((initialPendingConfirmation?.requestId ?? 0) + 1);

  const activeSession = sessions.find((session) => session.id === activeSessionId) ?? null;
  const visibleConfirmation =
    pendingConfirmation && pendingConfirmation.sessionId === activeSessionId
      ? pendingConfirmation
      : null;

  function refreshSessions() {
    setSessionState(sessionStore.getState());
  }

  function syncMessages(nextMessages: ChatMessage[]) {
    setMessages(nextMessages);

    if (activeSessionId) {
      sessionStore.replaceMessages(activeSessionId, nextMessages);
      refreshSessions();
    }
  }

  function updatePendingConfirmation(nextConfirmation: PendingExcelConfirmation | null) {
    pendingConfirmationRef.current = nextConfirmation;
    setPendingConfirmation(nextConfirmation);
  }

  useEffect(() => {
    if (!sessionStore.getState().sessions.length) {
      sessionStore.createSession();
      refreshSessions();
    }
  }, [sessionStore]);

  useEffect(() => {
    setMessages(activeSession?.messages.length ? activeSession.messages : initialMessages);
  }, [activeSessionId, activeSession]);

  useEffect(() => {
    return subscribeToSelectionChanges((nextSelection) => {
      setSelection(nextSelection);
    });
  }, []);

  function handleCreateSession() {
    sessionStore.createSession();
    setDraft("");
    refreshSessions();
  }

  function handleSelectSession(id: string) {
    sessionStore.setActiveSession(id);
    refreshSessions();
  }

  function handleDeleteSession(id: string) {
    sessionStore.deleteSession(id);
    refreshSessions();
  }

  function handleOpenSettings() {
    setSettings(settingsStore.load());
    setIsSettingsOpen(true);
  }

  function handleSaveSettings(nextSettings: SettingsState) {
    settingsStore.save(nextSettings);
    setSettings(nextSettings);
    setIsSettingsOpen(false);
  }

  async function executeExcelAction(action: ExcelAction, confirmation: PendingExcelConfirmation | null = null) {
    await excelAdapter.run(action);

    if (confirmation && pendingConfirmationRef.current?.requestId === confirmation.requestId) {
      updatePendingConfirmation(null);
    }
  }

  async function confirmPendingConfirmation() {
    const current = pendingConfirmationRef.current;
    if (!current || current.isExecuting) {
      return;
    }

    const executing = {
      ...current,
      isExecuting: true,
      error: null,
    };
    updatePendingConfirmation(executing);

    try {
      await executeExcelAction(current.action, current);
    } catch (error) {
      if (pendingConfirmationRef.current?.requestId === current.requestId) {
        updatePendingConfirmation({
          ...current,
          isExecuting: false,
          error: formatExecutionError(error),
        });
      }
    }
  }

  function cancelPendingConfirmation() {
    const current = pendingConfirmationRef.current;
    if (!current || current.isExecuting) {
      return;
    }

    updatePendingConfirmation(null);
  }

  function queueAction(action: ExcelAction) {
    if (classifyAction(action).requiresConfirmation) {
      const sessionId = activeSessionId;
      if (!sessionId) {
        return;
      }

      updatePendingConfirmation({
        requestId: nextRequestIdRef.current,
        sessionId,
        action,
        isExecuting: false,
        error: null,
      });
      nextRequestIdRef.current += 1;
      return;
    }

    void executeExcelAction(action);
  }

  function handleSubmit() {
    const content = draft.trim();
    if (!content) {
      return;
    }

    const route = decideRoute(content);
    if (route.mode !== "chat") {
      return;
    }

    const userMessage: ChatMessage = {
      id: crypto.randomUUID(),
      role: "user",
      content,
    };

    const nextMessages = activeSession?.messages.length ? [...messages, userMessage] : [userMessage];
    syncMessages(nextMessages);
    setDraft("");
  }

  return (
    <main className="layout">
      <SessionSidebar
        sessions={sessions}
        activeSessionId={activeSessionId ?? undefined}
        onCreateSession={handleCreateSession}
        onSelectSession={handleSelectSession}
        onDeleteSession={handleDeleteSession}
      />
      <section className="chat-panel">
        <header className="chat-header">
          <h1>OfficeAgent</h1>
          <button type="button" className="chat-header-action" onClick={handleOpenSettings}>
            {"\u8bbe\u7f6e"}
          </button>
        </header>
        {isSettingsOpen ? (
          <SettingsDialog
            initialValue={settings}
            onSave={handleSaveSettings}
            onClose={() => setIsSettingsOpen(false)}
          />
        ) : null}
        <MessageThread
          messages={messages}
          confirmation={
            visibleConfirmation ? (
              <ConfirmationCard
                summary={`Prepare to run ${visibleConfirmation.action.type}`}
                error={visibleConfirmation.error}
                isExecuting={visibleConfirmation.isExecuting}
                onConfirm={() => {
                  void confirmPendingConfirmation();
                }}
                onCancel={cancelPendingConfirmation}
              />
            ) : null
          }
        />
        <SelectionBadge selection={selection} />
        <Composer value={draft} onChange={setDraft} onSubmit={handleSubmit} />
      </section>
    </main>
  );
}
