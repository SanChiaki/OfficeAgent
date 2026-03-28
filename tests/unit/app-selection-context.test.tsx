import { act, render, screen, waitFor } from "@testing-library/react";
import { afterEach, expect, test, vi } from "vitest";
import App from "../../src/App";

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

test("updates the selection badge when Office reports a new selection", async () => {
  let registeredHandler: ((eventArgs: unknown) => Promise<void> | void) | null = null;
  const addHandlerAsync = vi.fn((eventType: string, handler: typeof registeredHandler, callback?: (result: { status: string }) => void) => {
    registeredHandler = handler;
    callback?.({ status: "succeeded" });
  });
  const removeHandlerAsync = vi.fn((eventType: string, payload: { handler: typeof registeredHandler }, callback?: (result: { status: string }) => void) => {
    callback?.({ status: "succeeded" });
  });
  const load = vi.fn();
  const sync = vi.fn(async () => {});
  const getSelectedRange = vi.fn(() => ({
    address: "D4:E6",
    rowCount: 3,
    columnCount: 2,
    load,
    worksheet: {
      load: vi.fn(),
      name: "Sheet7",
    },
  }));

  vi.stubGlobal("Office", {
    context: {
      document: {
        addHandlerAsync,
        removeHandlerAsync,
      },
    },
  });
  vi.stubGlobal("Excel", {
    run: vi.fn(async (callback: (context: { workbook: { getSelectedRange: typeof getSelectedRange }; sync: typeof sync }) => Promise<void>) =>
      callback({
        workbook: { getSelectedRange },
        sync,
      }),
    ),
  });

  render(<App />);

  expect(screen.getByText("当前选区：未选择")).toBeInTheDocument();
  expect(registeredHandler).toEqual(expect.any(Function));

  await act(async () => {
    await registeredHandler?.({ document: {} });
  });

  await waitFor(() => {
    expect(screen.getByText("当前选区：Sheet7!D4:E6 ｜ 3 行 ｜ 2 列")).toBeInTheDocument();
  });

  expect(removeHandlerAsync).not.toHaveBeenCalled();
});
