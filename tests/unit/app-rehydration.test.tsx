import { render, screen } from "@testing-library/react";
import { beforeEach, expect, test } from "vitest";
import App from "../../src/App";
import type { ChatSession } from "../../src/types";

beforeEach(() => {
  window.localStorage.clear();
});

test("rehydrates a fresh app instance from localStorage", () => {
  const session: ChatSession = {
    id: "session-1",
    title: "已保存会话",
    messages: [{ id: "m1", role: "assistant", content: "persisted message" }],
  };

  window.localStorage.setItem("oa:sessions:index", JSON.stringify([session]));
  window.localStorage.setItem("oa:runtime:activeSessionId", JSON.stringify(session.id));

  render(<App />);

  expect(screen.getByRole("button", { name: "已保存会话" })).toBeInTheDocument();
  expect(screen.getByText("persisted message")).toBeInTheDocument();
});
