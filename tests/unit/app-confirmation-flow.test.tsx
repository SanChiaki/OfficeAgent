import { cleanup, fireEvent, render, screen, within } from "@testing-library/react";
import { afterEach, beforeEach, expect, test, vi } from "vitest";
import App from "../../src/App";
import type { ChatMessage } from "../../src/types";
import type { ExcelAction } from "../../src/excel/excelAdapter";

type SessionState = {
  sessions: Array<{
    id: string;
    title: string;
    messages: ChatMessage[];
  }>;
  activeSessionId: string | null;
};

function createMockSessionStore(initialState: SessionState) {
  let state = initialState;

  return {
    createSession() {
      const session = {
        id: `session-${state.sessions.length + 1}`,
        title: "新对话",
        messages: [] as ChatMessage[],
      };
      state = {
        sessions: [session, ...state.sessions],
        activeSessionId: session.id,
      };
      return session;
    },
    deleteSession(id: string) {
      state = {
        sessions: state.sessions.filter((session) => session.id !== id),
        activeSessionId: state.activeSessionId === id ? state.sessions[0]?.id ?? null : state.activeSessionId,
      };
    },
    setActiveSession(id: string) {
      state = { ...state, activeSessionId: id };
    },
    replaceMessages(id: string, messages: ChatMessage[]) {
      state = {
        ...state,
        sessions: state.sessions.map((session) =>
          session.id === id ? { ...session, messages } : session,
        ),
      };
    },
    getState() {
      return state;
    },
  };
}

function createPendingConfirmation(sessionId: string, requestId = 1) {
  const action: ExcelAction = { type: "excel.writeRange", args: {} };

  return {
    requestId,
    sessionId,
    action,
    isExecuting: false,
    error: null,
  };
}

function createMockAdapter(run: () => Promise<string>) {
  return {
    readSelectionTable: async () => ({
      headers: ["Name", "Owner"],
      rows: [["项目A", "张三"]],
    }),
    run,
  };
}

beforeEach(() => {
  window.localStorage.clear();
});

afterEach(() => {
  cleanup();
});

test("confirm is single-flight while execution is in progress", () => {
  const run = vi.fn(() => new Promise<string>(() => {}));
  const sessionStore = createMockSessionStore({
    sessions: [
      {
        id: "session-1",
        title: "Alpha",
        messages: [],
      },
    ],
    activeSessionId: "session-1",
  });

  render(
    <App
      sessionStoreFactory={() => sessionStore}
      excelAdapterFactory={() => createMockAdapter(run)}
      initialPendingConfirmation={createPendingConfirmation("session-1")}
    />,
  );

  const confirmButton = screen.getByRole("button", { name: "确认" });
  const cancelButton = screen.getByRole("button", { name: "取消" });

  fireEvent.click(confirmButton);
  fireEvent.click(confirmButton);

  expect(run).toHaveBeenCalledTimes(1);
  expect(confirmButton).toBeDisabled();
  expect(cancelButton).toBeDisabled();
});

test("failure keeps the confirmation visible and allows retry", async () => {
  const run = vi.fn(() => Promise.reject(new Error("boom")) as Promise<string>);
  const sessionStore = createMockSessionStore({
    sessions: [
      {
        id: "session-1",
        title: "Alpha",
        messages: [],
      },
    ],
    activeSessionId: "session-1",
  });

  render(
    <App
      sessionStoreFactory={() => sessionStore}
      excelAdapterFactory={() => createMockAdapter(run)}
      initialPendingConfirmation={createPendingConfirmation("session-1")}
    />,
  );

  fireEvent.click(screen.getByRole("button", { name: "确认" }));

  expect(await screen.findByText("执行失败：boom")).toBeInTheDocument();
  expect(screen.getByRole("button", { name: "确认" })).not.toBeDisabled();
  fireEvent.click(screen.getByRole("button", { name: "确认" }));
  expect(run).toHaveBeenCalledTimes(2);
});

test("pending confirmation stays scoped to its owning session", () => {
  const sessionStore = createMockSessionStore({
    sessions: [
      {
        id: "session-1",
        title: "Alpha",
        messages: [],
      },
    ],
    activeSessionId: "session-1",
  });

  render(
    <App
      sessionStoreFactory={() => sessionStore}
      excelAdapterFactory={() =>
        createMockAdapter(vi.fn(() => Promise.resolve("excel.writeRange")))
      }
      initialPendingConfirmation={createPendingConfirmation("session-1")}
    />,
  );

  expect(screen.getByRole("button", { name: "确认" })).toBeInTheDocument();

  fireEvent.click(screen.getByRole("button", { name: "新建对话" }));

  expect(screen.queryByRole("button", { name: "确认" })).not.toBeInTheDocument();

  fireEvent.click(screen.getByRole("button", { name: "Alpha" }));

  expect(screen.getByRole("button", { name: "确认" })).toBeInTheDocument();
});
