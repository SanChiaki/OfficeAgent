import type { SelectionContext } from "../types";

export interface RawSelection {
  sheetName: string;
  address: string;
  rowCount: number;
  columnCount: number;
}

type SelectionRange = {
  address?: string;
  rowCount?: number;
  columnCount?: number;
  load?: (properties: string[] | string) => void;
  worksheet?: {
    name?: string;
    load?: (properties: string[] | string) => void;
  };
};

type SelectionContextObject = {
  workbook: {
    getSelectedRange: () => SelectionRange;
  };
  sync: () => Promise<void>;
};

type ExcelRuntime = {
  run?: <T>(callback: (context: SelectionContextObject) => Promise<T> | T) => Promise<T>;
};

type OfficeDocument = {
  addHandlerAsync?: (
    eventType: string,
    handler: () => Promise<void> | void,
    callback?: (result: { status: string }) => void,
  ) => void;
  removeHandlerAsync?: (
    eventType: string,
    handler: { handler?: () => Promise<void> | void } | (() => Promise<void> | void),
    callback?: (result: { status: string }) => void,
  ) => void;
};

type OfficeRuntime = {
  context?: {
    document?: OfficeDocument;
  };
};

export function normalizeSelection(raw: RawSelection): SelectionContext {
  return {
    ...raw,
    hasHeaders: false,
  };
}

async function readCurrentSelection(): Promise<SelectionContext | null> {
  const runtime = window as unknown as { Excel?: ExcelRuntime };
  if (!runtime.Excel?.run) {
    return null;
  }

  return runtime.Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load?.(["address", "rowCount", "columnCount"]);
    range.worksheet?.load?.("name");
    await context.sync();

    const sheetName = range.worksheet?.name;
    const address = range.address;
    const rowCount = range.rowCount;
    const columnCount = range.columnCount;

    if (
      typeof sheetName !== "string" ||
      typeof address !== "string" ||
      typeof rowCount !== "number" ||
      typeof columnCount !== "number"
    ) {
      return null;
    }

    return normalizeSelection({
      sheetName,
      address,
      rowCount,
      columnCount,
    });
  });
}

export function subscribeToSelectionChanges(onChange: (selection: SelectionContext) => void) {
  const office = (window as unknown as { Office?: OfficeRuntime }).Office;
  const document = office?.context?.document;

  if (!document?.addHandlerAsync || !document?.removeHandlerAsync) {
    return () => {};
  }

  const addHandlerAsync = document.addHandlerAsync;
  const removeHandlerAsync = document.removeHandlerAsync;
  let disposed = false;

  const handler = async () => {
    try {
      const selection = await readCurrentSelection();
      if (!disposed && selection) {
        onChange(selection);
      }
    } catch {
      // Selection refresh is best-effort. Ignore runtime failures for now.
    }
  };

  addHandlerAsync("documentSelectionChanged", handler, () => {});

  return () => {
    disposed = true;
    removeHandlerAsync("documentSelectionChanged", { handler }, () => {});
  };
}
