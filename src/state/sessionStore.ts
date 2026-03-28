import type { ChatMessage, ChatSession } from "../types";
import { getJson, setJson } from "./localStorageAdapter";

const INDEX_KEY = "oa:sessions:index";
const ACTIVE_KEY = "oa:runtime:activeSessionId";

interface SessionState {
  sessions: ChatSession[];
  activeSessionId: string | null;
}

export function createSessionStore() {
  let sessions = getJson<ChatSession[]>(INDEX_KEY, []);
  let activeSessionId = getJson<string | null>(ACTIVE_KEY, null);

  return {
    createSession() {
      const session: ChatSession = {
        id: crypto.randomUUID(),
        title: "新对话",
        messages: [],
      };

      sessions = [session, ...sessions];
      activeSessionId = session.id;
      setJson(INDEX_KEY, sessions);
      setJson(ACTIVE_KEY, activeSessionId);
      return session;
    },
    deleteSession(id: string) {
      sessions = sessions.filter((session) => session.id !== id);
      if (activeSessionId === id) {
        activeSessionId = sessions[0]?.id ?? null;
      }
      setJson(INDEX_KEY, sessions);
      setJson(ACTIVE_KEY, activeSessionId);
    },
    setActiveSession(id: string) {
      activeSessionId = id;
      setJson(ACTIVE_KEY, activeSessionId);
    },
    replaceMessages(id: string, messages: ChatMessage[]) {
      sessions = sessions.map((session) =>
        session.id === id
          ? {
              ...session,
              messages,
              title: messages[0]?.content.slice(0, 12) || session.title,
            }
          : session
      );
      setJson(INDEX_KEY, sessions);
    },
    getState(): SessionState {
      return { sessions, activeSessionId };
    },
  };
}
