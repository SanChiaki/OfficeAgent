using System;
using System.Collections.Generic;

namespace OfficeAgent.ExcelAddIn.Excel
{
    internal sealed class WorksheetPendingEditTracker
    {
        private readonly Dictionary<string, string> beforeValues =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        private readonly Dictionary<string, string> pendingOriginalValues =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public void CaptureBeforeValues(string sheetName, IReadOnlyList<WorksheetCellValue> cells)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                return;
            }

            foreach (var cell in cells ?? Array.Empty<WorksheetCellValue>())
            {
                if (cell == null || cell.Row <= 0 || cell.Column <= 0)
                {
                    continue;
                }

                beforeValues[BuildKey(sheetName, cell.Row, cell.Column)] = cell.Text ?? string.Empty;
            }
        }

        public void MarkChanged(string sheetName, IReadOnlyList<WorksheetCellAddress> cells)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                return;
            }

            foreach (var cell in cells ?? Array.Empty<WorksheetCellAddress>())
            {
                if (cell == null || cell.Row <= 0 || cell.Column <= 0)
                {
                    continue;
                }

                var key = BuildKey(sheetName, cell.Row, cell.Column);
                if (pendingOriginalValues.ContainsKey(key))
                {
                    continue;
                }

                if (beforeValues.TryGetValue(key, out var value))
                {
                    pendingOriginalValues[key] = value ?? string.Empty;
                }
            }
        }

        public bool TryGetOriginalValue(string sheetName, int row, int column, out string value)
        {
            value = string.Empty;
            if (string.IsNullOrWhiteSpace(sheetName) || row <= 0 || column <= 0)
            {
                return false;
            }

            return pendingOriginalValues.TryGetValue(BuildKey(sheetName, row, column), out value);
        }

        public void Clear(string sheetName, int row, int column)
        {
            if (string.IsNullOrWhiteSpace(sheetName) || row <= 0 || column <= 0)
            {
                return;
            }

            var key = BuildKey(sheetName, row, column);
            beforeValues.Remove(key);
            pendingOriginalValues.Remove(key);
        }

        public void Clear(string sheetName, IReadOnlyList<WorksheetCellAddress> cells)
        {
            foreach (var cell in cells ?? Array.Empty<WorksheetCellAddress>())
            {
                if (cell == null)
                {
                    continue;
                }

                Clear(sheetName, cell.Row, cell.Column);
            }
        }

        private static string BuildKey(string sheetName, int row, int column)
        {
            return $"{sheetName}|{row}|{column}";
        }
    }
}
