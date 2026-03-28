import { useEffect, useMemo, useState } from "react";
import { Composer } from "./components/Composer";
import { MessageThread } from "./components/MessageThread";
import { SelectionBadge } from "./components/SelectionBadge";
import { SessionSidebar } from "./components/SessionSidebar";
import { createSessionStore } from "./state/sessionStore";
import type { ChatMessage } from "./types";

const initialMessages: ChatMessage[] = [
  {
    id: "assistant-welcome",
    role: "assistant",
    content: "你好，我是 OfficeAgent。",
  },
];

export default function App() {
  const sessionStore = useMemo(() => createSessionStore(), []);
  const [{ sessions, activeSessionId }, setSessionState] = useState(sessionStore.getState());
  const [draft, setDraft] = useState("");
  const [messages, setMessages] = useState<ChatMessage[]>(initialMessages);

  const activeSession = sessions.find((session) => session.id === activeSessionId) ?? null;

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

  useEffect(() => {
    if (!sessionStore.getState().sessions.length) {
      sessionStore.createSession();
      refreshSessions();
    }
  }, [sessionStore]);

  useEffect(() => {
    setMessages(activeSession?.messages.length ? activeSession.messages : initialMessages);
  }, [activeSessionId, activeSession]);

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

  function handleSubmit() {
    const content = draft.trim();
    if (!content) {
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
        </header>
        <MessageThread messages={messages} />
        <SelectionBadge selection={null} />
        <Composer value={draft} onChange={setDraft} onSubmit={handleSubmit} />
      </section>
    </main>
  );
}
