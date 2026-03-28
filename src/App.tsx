import { useState } from "react";
import { Composer } from "./components/Composer";
import { MessageThread } from "./components/MessageThread";
import { SelectionBadge } from "./components/SelectionBadge";
import { SessionSidebar } from "./components/SessionSidebar";
import type { ChatMessage } from "./types";

const initialMessages: ChatMessage[] = [
  {
    id: "assistant-welcome",
    role: "assistant",
    content: "你好，我是 OfficeAgent。"
  }
];

export default function App() {
  const [draft, setDraft] = useState("");

  return (
    <main className="layout">
      <SessionSidebar
        sessions={[]}
        activeSessionId="default"
        onCreateSession={() => {}}
        onSelectSession={() => {}}
        onDeleteSession={() => {}}
      />
      <section className="chat-panel">
        <header className="chat-header">
          <h1>OfficeAgent</h1>
        </header>
        <MessageThread messages={initialMessages} />
        <SelectionBadge selection={null} />
        <Composer value={draft} onChange={setDraft} onSubmit={() => {}} />
      </section>
    </main>
  );
}
