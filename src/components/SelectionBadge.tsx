import type { SelectionContext } from "../types";

export function SelectionBadge({ selection }: { selection: SelectionContext | null }) {
  if (!selection) {
    return <div className="selection-badge">当前选区：未选择</div>;
  }

  return (
    <div className="selection-badge">
      {"褰撳墠閫夊尯锛歿selection.sheetName}!{selection.address} 锝?{selection.rowCount} 琛?锝?{selection.columnCount} 鍒?"}
    </div>
  );
}
