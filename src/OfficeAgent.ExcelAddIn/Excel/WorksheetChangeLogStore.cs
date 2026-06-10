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
        private const string TextNumberFormat = "@";

        private static readonly string[] Headers =
        {
            "Key",
            "Header",
            "Change Mode",
            "New Value",
            "Old Value",
            "Changed At",
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
                var worksheetRow = row + 2;
                var key = ReadStableCellText(values, row, 0, worksheetRow, 1);
                var headerText = ReadStableCellText(values, row, 1, worksheetRow, 2);
                var changeMode = ReadStableCellText(values, row, 2, worksheetRow, 3);
                var newValue = ReadStableCellText(values, row, 3, worksheetRow, 4);
                var oldValue = ReadStableCellText(values, row, 4, worksheetRow, 5);
                var changedAtValue = values[row, 5];
                var changedAtText = Convert.ToString(changedAtValue) ?? string.Empty;
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
                    ChangeMode = NormalizeChangeMode(changeMode),
                    NewValue = newValue,
                    OldValue = oldValue,
                    ChangedAt = ParseTimestamp(changedAtValue),
                });
            }

            return rows.ToArray();
        }

        private void RewriteRows(IReadOnlyList<WorksheetChangeLogEntry> rows)
        {
            var existingLastRow = Math.Max(1, gridAdapter.GetLastUsedRow(LogSheetName));
            gridAdapter.ClearRange(LogSheetName, 1, Math.Max(existingLastRow, rows.Count + 1), 1, Headers.Length);
            gridAdapter.SetRangeNumberFormat(LogSheetName, 1, rows.Count + 1, 1, Headers.Length, TextNumberFormat);
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
                result[row + 1, 2] = NormalizeChangeMode(entry.ChangeMode);
                result[row + 1, 3] = entry.NewValue ?? string.Empty;
                result[row + 1, 4] = entry.OldValue ?? string.Empty;
                result[row + 1, 5] = (entry.ChangedAt == default(DateTime) ? now : entry.ChangedAt)
                    .ToString(TimestampFormat, CultureInfo.InvariantCulture);
            }

            return result;
        }

        private string ReadStableCellText(object[,] values, int row, int column, int worksheetRow, int worksheetColumn)
        {
            var value = values[row, column];
            if (value is string textValue)
            {
                return textValue;
            }

            var text = gridAdapter.GetCellText(LogSheetName, worksheetRow, worksheetColumn);
            if (!string.IsNullOrEmpty(text))
            {
                return text;
            }

            return Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
        }

        private static string NormalizeChangeMode(string value)
        {
            var text = value ?? string.Empty;
            if (string.Equals(text, "下载", StringComparison.Ordinal) ||
                string.Equals(text, "Download", StringComparison.OrdinalIgnoreCase))
            {
                return "Download";
            }

            if (string.Equals(text, "上传", StringComparison.Ordinal) ||
                string.Equals(text, "Upload", StringComparison.OrdinalIgnoreCase))
            {
                return "Upload";
            }

            return text;
        }

        private static DateTime ParseTimestamp(object value)
        {
            if (value is DateTime timestamp)
            {
                return timestamp;
            }

            if (value is double serial)
            {
                return ParseOADate(serial);
            }

            if (value is float floatSerial)
            {
                return ParseOADate(floatSerial);
            }

            if (value is decimal decimalSerial)
            {
                return ParseOADate((double)decimalSerial);
            }

            if (value is int intSerial)
            {
                return ParseOADate(intSerial);
            }

            var text = Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
            return DateTime.TryParseExact(
                text,
                TimestampFormat,
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out var parsed)
                ? parsed
                : default(DateTime);
        }

        private static DateTime ParseOADate(double value)
        {
            if (double.IsNaN(value) || double.IsInfinity(value))
            {
                return default(DateTime);
            }

            try
            {
                return DateTime.FromOADate(value);
            }
            catch (ArgumentException)
            {
                return default(DateTime);
            }
        }
    }
}
