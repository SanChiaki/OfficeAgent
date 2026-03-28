import type { ChatMessage } from "../types";

export function MessageThread({ messages }: { messages: ChatMessage[] }) {
  return (
    <section className="message-thread" aria-label="消息线程">
      {messages.map((message) => (
        <article key={message.id} className={`message message-${message.role}`}>
          {message.content}
        </article>
      ))}
    </section>
  );
}
