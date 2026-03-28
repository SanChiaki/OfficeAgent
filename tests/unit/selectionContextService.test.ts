import { afterEach, expect, test, vi } from "vitest";
import { subscribeToSelectionChanges } from "../../src/excel/selectionContextService";

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

test("registers a selection handler, emits normalized metadata, and cleans up", async () => {
  type SelectionHandler = (eventArgs: unknown) => Promise<void> | void;

  let registeredHandler: SelectionHandler | null = null;
  const removeHandlerAsync = vi.fn((eventType: string, payload: { handler: SelectionHandler | null }, callback?: (result: { status: string }) => void) => {
    callback?.({ status: "succeeded" });
  });
  const addHandlerAsync = vi.fn((eventType: string, handler: SelectionHandler, callback?: (result: { status: string }) => void) => {
    registeredHandler = handler;
    callback?.({ status: "succeeded" });
  });
  const load = vi.fn();
  const worksheetLoad = vi.fn();
  const sync = vi.fn(async () => {});
  const getSelectedRange = vi.fn(() => ({
    address: "B2:C3",
    rowCount: 2,
    columnCount: 2,
    load,
    worksheet: {
      load: worksheetLoad,
      name: "Budget",
    },
  }));
  const run = vi.fn(async (callback: (context: { workbook: { getSelectedRange: typeof getSelectedRange }; sync: typeof sync }) => Promise<void>) => {
    return callback({
      workbook: { getSelectedRange },
      sync,
    });
  });

  vi.stubGlobal("Office", {
    context: {
      document: {
        addHandlerAsync,
        removeHandlerAsync,
      },
    },
  });
  vi.stubGlobal("Excel", { run });

  const onChange = vi.fn();
  const cleanup = subscribeToSelectionChanges(onChange);

  expect(addHandlerAsync).toHaveBeenCalledWith(
    "documentSelectionChanged",
    expect.any(Function),
    expect.any(Function),
  );

  const selectionHandler = registeredHandler;
  expect(selectionHandler).toEqual(expect.any(Function));
  if (!selectionHandler) {
    throw new Error("Selection handler was not registered");
  }

  await (selectionHandler as SelectionHandler)({ document: {} });

  expect(run).toHaveBeenCalledTimes(1);
  expect(getSelectedRange).toHaveBeenCalledTimes(1);
  expect(load).toHaveBeenCalledWith(["address", "rowCount", "columnCount"]);
  expect(worksheetLoad).toHaveBeenCalledWith("name");
  expect(sync).toHaveBeenCalledTimes(1);
  expect(onChange).toHaveBeenCalledWith({
    sheetName: "Budget",
    address: "B2:C3",
    rowCount: 2,
    columnCount: 2,
    hasHeaders: false,
  });

  cleanup();

  expect(removeHandlerAsync).toHaveBeenCalledWith(
    "documentSelectionChanged",
    { handler: selectionHandler },
    expect.any(Function),
  );
});
