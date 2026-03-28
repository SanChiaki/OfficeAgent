import { render, screen } from "@testing-library/react";
import { expect, test } from "vitest";
import App from "../../src/App";

test("renders task pane layout primitives", () => {
  render(<App />);
  expect(screen.getByRole("button", { name: "新建对话" })).toBeInTheDocument();
  expect(screen.getByPlaceholderText("输入你的问题或命令")).toBeInTheDocument();
  expect(screen.getByText("当前选区：未选择")).toBeInTheDocument();
});
