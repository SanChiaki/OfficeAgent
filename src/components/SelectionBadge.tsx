import type { SelectionContext } from "../types";

export function SelectionBadge({ selection }: { selection: SelectionContext | null }) {
  if (!selection) {
    return <div className="selection-badge">当前选区：未选择</div>;
  }

  return (
    <div className="selection-badge">
      当前选区：{selection.sheetName}!{selection.address} 行 {selection.rowCount} 列 {selection.columnCount}
    </div>
  );
}
