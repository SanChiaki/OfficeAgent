import { fireEvent, render, screen, within } from "@testing-library/react";
import { beforeEach, expect, test } from "vitest";
import App from "../../src/App";

beforeEach(() => {
  window.localStorage.clear();
});

test("keeps messages isolated within the active session", () => {
  render(<App />);

  const input = screen.getByPlaceholderText("输入你的问题或命令");
  const sendButton = screen.getByRole("button", { name: "发送" });
  const createButton = screen.getByRole("button", { name: "新建对话" });
  const thread = screen.getByRole("region", { name: "消息线程" });

  fireEvent.change(input, { target: { value: "alpha" } });
  fireEvent.click(sendButton);

  expect(within(thread).getByText("alpha")).toBeInTheDocument();

  fireEvent.click(createButton);
  fireEvent.change(input, { target: { value: "beta" } });
  fireEvent.click(sendButton);

  expect(within(thread).getByText("beta")).toBeInTheDocument();

  fireEvent.click(screen.getByRole("button", { name: "alpha" }));

  expect(within(thread).getByText("alpha")).toBeInTheDocument();
  expect(within(thread).queryByText("beta")).not.toBeInTheDocument();
});
