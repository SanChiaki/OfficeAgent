import type { ChatSession } from "../types";

interface SessionSidebarProps {
  sessions: Array<Pick<ChatSession, "id" | "title">>;
  activeSessionId?: string;
  onCreateSession(): void;
  onSelectSession(id: string): void;
  onDeleteSession(id: string): void;
}

export function SessionSidebar(props: SessionSidebarProps) {
  return (
    <aside className="session-sidebar" aria-label="会话列表">
      <button type="button" className="sidebar-action" onClick={props.onCreateSession}>
        新建对话
      </button>
      <ul className="session-list">
        {props.sessions.map((session) => (
          <li key={session.id}>
            <div className={session.id === props.activeSessionId ? "session-item is-active" : "session-item"}>
              <button type="button" className="session-title" onClick={() => props.onSelectSession(session.id)}>
                {session.title}
              </button>
              <button type="button" className="session-remove" onClick={() => props.onDeleteSession(session.id)}>
                删除
              </button>
            </div>
          </li>
        ))}
      </ul>
    </aside>
  );
}
