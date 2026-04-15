using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;

namespace OfficeAgent.ExcelAddIn.Excel
{
    internal sealed class WorksheetMetadataStore : IWorksheetMetadataStore
    {
        private const string MetadataSheetName = "_Settings";
        private const string BindingsTableName = "SheetBindings";
        private const string FieldMappingsTableName = "SheetFieldMappings";
        private static readonly string[] DefaultFieldMappingHeaders = { "SheetName" };

        private static readonly string[] BindingHeaders =
        {
            "SheetName",
            "SystemKey",
            "ProjectId",
            "ProjectName",
            "HeaderStartRow",
            "HeaderRowCount",
            "DataStartRow",
        };

        private readonly IWorksheetMetadataAdapter adapter;
        private string[] fieldMappingHeaders = DefaultFieldMappingHeaders.ToArray();

        public WorksheetMetadataStore(IWorksheetMetadataAdapter adapter)
        {
            this.adapter = adapter ?? throw new ArgumentNullException(nameof(adapter));
        }

        public void SaveBinding(SheetBinding binding)
        {
            if (binding == null)
            {
                throw new ArgumentNullException(nameof(binding));
            }

            if (string.IsNullOrWhiteSpace(binding.SheetName))
            {
                throw new ArgumentException("Sheet name is required.", nameof(binding));
            }

            adapter.EnsureWorksheet(MetadataSheetName, visible: true);
            var normalizedSheetName = binding.SheetName;
            var rows = adapter.ReadTable(BindingsTableName)?.ToList() ?? new List<string[]>();
            var newRow = new[]
            {
                normalizedSheetName,
                binding.SystemKey ?? string.Empty,
                binding.ProjectId ?? string.Empty,
                binding.ProjectName ?? string.Empty,
                binding.HeaderStartRow.ToString(CultureInfo.InvariantCulture),
                binding.HeaderRowCount.ToString(CultureInfo.InvariantCulture),
                binding.DataStartRow.ToString(CultureInfo.InvariantCulture),
            };

            var existingRowIndex = rows.FindIndex(
                row => row.Length > 0 &&
                       string.Equals(row[0], normalizedSheetName, StringComparison.OrdinalIgnoreCase));

            if (existingRowIndex >= 0)
            {
                rows[existingRowIndex] = newRow;
            }
            else
            {
                rows.Add(newRow);
            }

            adapter.WriteTable(BindingsTableName, BindingHeaders, rows.ToArray());
        }

        public SheetBinding LoadBinding(string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                throw new ArgumentException("Sheet name is required.", nameof(sheetName));
            }

            var rows = adapter.ReadTable(BindingsTableName) ?? Array.Empty<string[]>();

            foreach (var row in rows)
            {
                if (row.Length < 4 || string.IsNullOrWhiteSpace(row[0]))
                {
                    continue;
                }

                if (!string.Equals(row[0], sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                return new SheetBinding
                {
                    SheetName = row[0],
                    SystemKey = row[1],
                    ProjectId = row[2],
                    ProjectName = row[3],
                    HeaderStartRow = ParseIntOrDefault(row, 4, defaultValue: 1),
                    HeaderRowCount = ParseIntOrDefault(row, 5, defaultValue: 2),
                    DataStartRow = ParseIntOrDefault(row, 6, defaultValue: 3),
                };
            }

            throw new InvalidOperationException($"Binding for worksheet '{sheetName}' does not exist.");
        }

        public void SaveFieldMappings(string sheetName, FieldMappingTableDefinition definition, IReadOnlyList<SheetFieldMappingRow> rows)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                throw new ArgumentException("Sheet name is required.", nameof(sheetName));
            }

            if (definition == null)
            {
                throw new ArgumentNullException(nameof(definition));
            }

            adapter.EnsureWorksheet(MetadataSheetName, visible: true);
            var columns = GetValidatedColumns(definition);
            var headers = new[] { "SheetName" }
                .Concat(columns.Select(column => column.ColumnName))
                .ToArray();
            fieldMappingHeaders = headers.ToArray();

            var existingRows = adapter.ReadTable(FieldMappingsTableName)?.ToList() ?? new List<string[]>();
            existingRows.RemoveAll(row =>
                row.Length > 0 &&
                string.Equals(row[0], sheetName, StringComparison.OrdinalIgnoreCase));

            foreach (var mappingRow in rows ?? Array.Empty<SheetFieldMappingRow>())
            {
                var values = new string[columns.Length + 1];
                values[0] = sheetName;

                for (var columnIndex = 0; columnIndex < columns.Length; columnIndex++)
                {
                    var columnName = columns[columnIndex].ColumnName;
                    values[columnIndex + 1] = mappingRow?.Values != null &&
                                              mappingRow.Values.TryGetValue(columnName, out var value)
                        ? value ?? string.Empty
                        : string.Empty;
                }

                existingRows.Add(values);
            }

            adapter.WriteTable(FieldMappingsTableName, headers, existingRows.ToArray());
        }

        public SheetFieldMappingRow[] LoadFieldMappings(string sheetName, FieldMappingTableDefinition definition)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                throw new ArgumentException("Sheet name is required.", nameof(sheetName));
            }

            if (definition == null)
            {
                throw new ArgumentNullException(nameof(definition));
            }

            var columns = GetValidatedColumns(definition);
            var rows = adapter.ReadTable(FieldMappingsTableName) ?? Array.Empty<string[]>();

            return rows
                .Where(row =>
                    row.Length > 0 &&
                    string.Equals(row[0], sheetName, StringComparison.OrdinalIgnoreCase))
                .Select(row =>
                {
                    var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                    for (var columnIndex = 0; columnIndex < columns.Length; columnIndex++)
                    {
                        var columnName = columns[columnIndex].ColumnName;
                        values[columnName] = row.Length > columnIndex + 1
                            ? row[columnIndex + 1]
                            : string.Empty;
                    }

                    return new SheetFieldMappingRow
                    {
                        SheetName = sheetName,
                        Values = values,
                    };
                })
                .ToArray();
        }

        public void ClearFieldMappings(string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                throw new ArgumentException("Sheet name is required.", nameof(sheetName));
            }

            var rows = adapter.ReadTable(FieldMappingsTableName)?.ToList() ?? new List<string[]>();
            var removed = rows.RemoveAll(row =>
                row.Length > 0 &&
                string.Equals(row[0], sheetName, StringComparison.OrdinalIgnoreCase));

            if (removed == 0)
            {
                return;
            }

            var headers = ResolveFieldMappingHeaders(rows);
            adapter.WriteTable(FieldMappingsTableName, headers, rows.ToArray());
        }

        public WorksheetSnapshotCell[] LoadSnapshot(string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                throw new ArgumentException("Sheet name is required.", nameof(sheetName));
            }

            return Array.Empty<WorksheetSnapshotCell>();
        }

        public void SaveSnapshot(string sheetName, WorksheetSnapshotCell[] cells)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                throw new ArgumentException("Sheet name is required.", nameof(sheetName));
            }

            if (cells == null)
            {
                throw new ArgumentNullException(nameof(cells));
            }
        }

        private static int ParseIntOrDefault(IReadOnlyList<string> row, int index, int defaultValue)
        {
            if (row == null || row.Count <= index)
            {
                return defaultValue;
            }

            return int.TryParse(row[index], NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed)
                ? parsed
                : defaultValue;
        }

        private static FieldMappingColumnDefinition[] GetValidatedColumns(FieldMappingTableDefinition definition)
        {
            var columns = definition.Columns ?? Array.Empty<FieldMappingColumnDefinition>();

            for (var index = 0; index < columns.Length; index++)
            {
                if (columns[index] == null || string.IsNullOrWhiteSpace(columns[index].ColumnName))
                {
                    throw new ArgumentException(
                        "Field mapping definition columns must have non-empty ColumnName values.",
                        nameof(definition));
                }
            }

            return columns;
        }

        private string[] ResolveFieldMappingHeaders(IReadOnlyList<string[]> rows)
        {
            if (fieldMappingHeaders != null && fieldMappingHeaders.Length > 0)
            {
                return fieldMappingHeaders.ToArray();
            }

            var maxColumns = rows?.Count > 0
                ? Math.Max(rows.Max(row => row?.Length ?? 0), 1)
                : 1;
            var headers = new string[maxColumns];
            headers[0] = "SheetName";

            for (var index = 1; index < maxColumns; index++)
            {
                headers[index] = "Column" + index.ToString(CultureInfo.InvariantCulture);
            }

            return headers;
        }
    }
}
