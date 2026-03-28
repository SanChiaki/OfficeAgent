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

  function refreshSessions() {
    setSessionState(sessionStore.getState());
  }

  useEffect(() => {
    if (!sessionStore.getState().sessions.length) {
      sessionStore.createSession();
      refreshSessions();
    }
  }, [sessionStore]);

  useEffect(() => {
    const activeSession = sessions.find((session) => session.id === activeSessionId) ?? null;
    setMessages(activeSession?.messages.length ? activeSession.messages : initialMessages);
  }, [activeSessionId, sessions]);

  function handleCreateSession() {
    sessionStore.createSession();
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
        <Composer value={draft} onChange={setDraft} onSubmit={() => {}} />
      </section>
    </main>
  );
}
