import { expect, test } from "vitest";
import { createSessionStore } from "../../src/state/sessionStore";

test("creates, switches, and deletes local sessions", () => {
  const store = createSessionStore();
  const first = store.createSession();
  const second = store.createSession();
  store.replaceMessages(second.id, [
    { id: "m1", role: "assistant", content: "你好，我是 OfficeAgent。" },
  ]);

  store.setActiveSession(second.id);
  store.deleteSession(first.id);

  expect(store.getState().activeSessionId).toBe(second.id);
  expect(store.getState().sessions).toHaveLength(1);
  expect(store.getState().sessions[0].messages).toHaveLength(1);
});
