using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OfficeAgent.ExcelAddIn.Excel
{
    internal sealed class WorksheetChangeLogStore : IWorksheetChangeLogStore
    {
        private const string LogSheetName = "xISDP_Log";
        private const int MaxEntries = 2000;
        private const string TimestampFormat = "yyyy-MM-dd HH:mm:ss";

        private static readonly string[] Headers =
        {
            "key",
            "表头",
            "修改模式",
            "修改值",
            "原始值",
            "修改时间",
        };

        private readonly IWorksheetGridAdapter gridAdapter;
        private readonly Func<DateTime> getNow;

        public WorksheetChangeLogStore(IWorksheetGridAdapter gridAdapter)
            : this(gridAdapter, () => DateTime.Now)
        {
        }

        public WorksheetChangeLogStore(IWorksheetGridAdapter gridAdapter, Func<DateTime> getNow)
        {
            this.gridAdapter = gridAdapter ?? throw new ArgumentNullException(nameof(gridAdapter));
            this.getNow = getNow ?? (() => DateTime.Now);
        }

        public void Append(IReadOnlyList<WorksheetChangeLogEntry> entries)
        {
            var incoming = (entries ?? Array.Empty<WorksheetChangeLogEntry>())
                .Where(entry => entry != null)
                .ToArray();
            if (incoming.Length == 0)
            {
                return;
            }

            using (gridAdapter.BeginBulkOperation())
            {
                gridAdapter.EnsureWorksheetExists(LogSheetName);
                var existing = ReadExistingRows();
                var combined = existing.Concat(incoming).ToArray();
                var rows = combined
                    .Skip(Math.Max(0, combined.Length - MaxEntries))
                    .ToArray();

                RewriteRows(rows);
            }
        }

        private WorksheetChangeLogEntry[] ReadExistingRows()
        {
            var lastRow = gridAdapter.GetLastUsedRow(LogSheetName);
            if (lastRow <= 1)
            {
                return Array.Empty<WorksheetChangeLogEntry>();
            }

            var values = gridAdapter.ReadRangeValues(LogSheetName, 2, lastRow, 1, Headers.Length);
            var rows = new List<WorksheetChangeLogEntry>();
            for (var row = 0; row < values.GetLength(0); row++)
            {
                var key = Convert.ToString(values[row, 0]) ?? string.Empty;
                var headerText = Convert.ToString(values[row, 1]) ?? string.Empty;
                var changeMode = Convert.ToString(values[row, 2]) ?? string.Empty;
                var newValue = Convert.ToString(values[row, 3]) ?? string.Empty;
                var oldValue = Convert.ToString(values[row, 4]) ?? string.Empty;
                var changedAtText = Convert.ToString(values[row, 5]) ?? string.Empty;
                if (string.IsNullOrWhiteSpace(key) &&
                    string.IsNullOrWhiteSpace(headerText) &&
                    string.IsNullOrWhiteSpace(changeMode) &&
                    string.IsNullOrWhiteSpace(newValue) &&
                    string.IsNullOrWhiteSpace(oldValue) &&
                    string.IsNullOrWhiteSpace(changedAtText))
                {
                    continue;
                }

                rows.Add(new WorksheetChangeLogEntry
                {
                    Key = key,
                    HeaderText = headerText,
                    ChangeMode = changeMode,
                    NewValue = newValue,
                    OldValue = oldValue,
                    ChangedAt = ParseTimestamp(changedAtText),
                });
            }

            return rows.ToArray();
        }

        private void RewriteRows(IReadOnlyList<WorksheetChangeLogEntry> rows)
        {
            var existingLastRow = Math.Max(1, gridAdapter.GetLastUsedRow(LogSheetName));
            gridAdapter.ClearRange(LogSheetName, 1, Math.Max(existingLastRow, rows.Count + 1), 1, Headers.Length);
            gridAdapter.WriteRangeValues(LogSheetName, 1, 1, BuildMatrix(rows));
        }

        private object[,] BuildMatrix(IReadOnlyList<WorksheetChangeLogEntry> rows)
        {
            var result = new object[(rows?.Count ?? 0) + 1, Headers.Length];
            for (var column = 0; column < Headers.Length; column++)
            {
                result[0, column] = Headers[column];
            }

            var now = getNow();
            for (var row = 0; row < (rows?.Count ?? 0); row++)
            {
                var entry = rows[row] ?? new WorksheetChangeLogEntry();
                result[row + 1, 0] = entry.Key ?? string.Empty;
                result[row + 1, 1] = entry.HeaderText ?? string.Empty;
                result[row + 1, 2] = entry.ChangeMode ?? string.Empty;
                result[row + 1, 3] = entry.NewValue ?? string.Empty;
                result[row + 1, 4] = entry.OldValue ?? string.Empty;
                result[row + 1, 5] = (entry.ChangedAt == default(DateTime) ? now : entry.ChangedAt)
                    .ToString(TimestampFormat, CultureInfo.InvariantCulture);
            }

            return result;
        }

        private static DateTime ParseTimestamp(string value)
        {
            return DateTime.TryParseExact(
                value,
                TimestampFormat,
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out var parsed)
                ? parsed
                : default(DateTime);
        }
    }
}
