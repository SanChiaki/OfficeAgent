using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;

namespace OfficeAgent.ExcelAddIn.Excel
{
    internal sealed class WorksheetMetadataStore : IWorksheetMetadataStore, IWorksheetTemplateBindingStore
    {
        private const string MetadataSheetName = "ISDP_Setting";
        private const string TemplateBindingsTableName = "TemplateBindings";
        private const string BindingsTableName = "SheetBindings";
        private const string FieldMappingsTableName = "SheetFieldMappings";
        private static readonly string[] DefaultFieldMappingHeaders = { "SheetName" };

        private static readonly string[] TemplateBindingHeaders =
        {
            "SheetName",
            "TemplateName",
            "TemplateRevision",
            "TemplateOrigin",
            "TemplateId",
            "AppliedFingerprint",
            "TemplateLastAppliedAt",
            "DerivedFromTemplateId",
            "DerivedFromTemplateRevision",
        };

        private static readonly string[] BindingHeaders =
        {
            "SheetName",
            "ProjectId",
            "ProjectName",
            "HeaderStartRow",
            "HeaderRowCount",
            "DataStartRow",
            "SystemKey",
        };

        private readonly IWorksheetMetadataAdapter adapter;
        private string[] fieldMappingHeaders = DefaultFieldMappingHeaders.ToArray();
        private string[][] templateBindingRowsCache;
        private bool templateBindingRowsCacheLoaded;
        private string[][] bindingRowsCache;
        private bool bindingRowsCacheLoaded;
        private string[][] fieldMappingRowsCache;
        private bool fieldMappingRowsCacheLoaded;
        private string workbookScopeKey = string.Empty;

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

            EnsureWorkbookScope();
            adapter.EnsureWorksheet(MetadataSheetName, visible: true);
            var normalizedSheetName = binding.SheetName;
            var rows = GetBindingRows().ToList();
            var newRow = new[]
            {
                normalizedSheetName,
                binding.ProjectId ?? string.Empty,
                binding.ProjectName ?? string.Empty,
                binding.HeaderStartRow.ToString(CultureInfo.InvariantCulture),
                binding.HeaderRowCount.ToString(CultureInfo.InvariantCulture),
                binding.DataStartRow.ToString(CultureInfo.InvariantCulture),
                binding.SystemKey ?? string.Empty,
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
            bindingRowsCache = CloneRows(rows);
            bindingRowsCacheLoaded = true;
        }

        public void SaveTemplateBinding(SheetTemplateBinding binding)
        {
            if (binding == null)
            {
                throw new ArgumentNullException(nameof(binding));
            }

            if (string.IsNullOrWhiteSpace(binding.SheetName))
            {
                throw new ArgumentException("Sheet name is required.", nameof(binding));
            }

            EnsureWorkbookScope();
            adapter.EnsureWorksheet(MetadataSheetName, visible: true);
            var normalizedSheetName = binding.SheetName;
            var rows = GetTemplateBindingRows().ToList();
            var newRow = new[]
            {
                normalizedSheetName,
                binding.TemplateName ?? string.Empty,
                binding.TemplateRevision?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
                binding.TemplateOrigin ?? string.Empty,
                binding.TemplateId ?? string.Empty,
                binding.AppliedFingerprint ?? string.Empty,
                binding.TemplateLastAppliedAt?.ToString("o", CultureInfo.InvariantCulture) ?? string.Empty,
                binding.DerivedFromTemplateId ?? string.Empty,
                binding.DerivedFromTemplateRevision?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
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

            adapter.WriteTable(TemplateBindingsTableName, TemplateBindingHeaders, rows.ToArray());
            templateBindingRowsCache = CloneRows(rows);
            templateBindingRowsCacheLoaded = true;
        }

        public SheetBinding LoadBinding(string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                throw new ArgumentException("Sheet name is required.", nameof(sheetName));
            }

            EnsureWorkbookScope();
            var binding = GetBindingRows()
                .Select(ParseBindingRow)
                .FirstOrDefault(candidate =>
                    candidate != null &&
                    string.Equals(candidate.SheetName, sheetName, StringComparison.OrdinalIgnoreCase));

            if (binding != null)
            {
                return binding;
            }

            throw new InvalidOperationException($"Binding for worksheet '{sheetName}' does not exist.");
        }

        public SheetTemplateBinding LoadTemplateBinding(string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                throw new ArgumentException("Sheet name is required.", nameof(sheetName));
            }

            EnsureWorkbookScope();
            var binding = GetTemplateBindingRows()
                .Select(ParseTemplateBindingRow)
                .FirstOrDefault(candidate =>
                    candidate != null &&
                    string.Equals(candidate.SheetName, sheetName, StringComparison.OrdinalIgnoreCase));

            if (binding != null)
            {
                return binding;
            }

            throw new InvalidOperationException($"Template binding for worksheet '{sheetName}' does not exist.");
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

            EnsureWorkbookScope();
            adapter.EnsureWorksheet(MetadataSheetName, visible: true);
            var columns = GetDistinctFieldMappingColumns(definition);
            var headers = new[] { "SheetName" }
                .Concat(columns.Select(column => column.ColumnName))
                .ToArray();
            fieldMappingHeaders = headers.ToArray();

            var existingRows = GetFieldMappingRows().ToList();
            existingRows.RemoveAll(row =>
                row.Length > 0 &&
                string.Equals(row[0], sheetName, StringComparison.OrdinalIgnoreCase));

            foreach (var mappingRow in rows ?? Array.Empty<SheetFieldMappingRow>())
            {
                var values = new string[columns.Length + 1];
                values[0] = sheetName;

                for (var columnIndex = 0; columnIndex < columns.Length; columnIndex++)
                {
                    var columnName = GetColumnValueKey(columns[columnIndex]);
                    values[columnIndex + 1] = mappingRow?.Values != null &&
                                              mappingRow.Values.TryGetValue(columnName, out var value)
                        ? value ?? string.Empty
                        : string.Empty;
                }

                existingRows.Add(values);
            }

            adapter.WriteTable(FieldMappingsTableName, headers, existingRows.ToArray());
            fieldMappingRowsCache = CloneRows(existingRows);
            fieldMappingRowsCacheLoaded = true;
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

            EnsureWorkbookScope();
            var columns = GetDistinctFieldMappingColumns(definition);
            var rows = GetFieldMappingRows();

            return rows
                .Where(row =>
                    row.Length > 0 &&
                    string.Equals(row[0], sheetName, StringComparison.OrdinalIgnoreCase))
                .Select(row =>
                {
                    var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                    for (var columnIndex = 0; columnIndex < columns.Length; columnIndex++)
                    {
                        var columnName = GetColumnValueKey(columns[columnIndex]);
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

            EnsureWorkbookScope();
            var rows = GetFieldMappingRows().ToList();
            var removed = rows.RemoveAll(row =>
                row.Length > 0 &&
                string.Equals(row[0], sheetName, StringComparison.OrdinalIgnoreCase));

            if (removed == 0)
            {
                return;
            }

            var headers = ResolveFieldMappingHeaders(rows);
            adapter.WriteTable(FieldMappingsTableName, headers, rows.ToArray());
            fieldMappingRowsCache = CloneRows(rows);
            fieldMappingRowsCacheLoaded = true;
        }

        public void ClearTemplateBinding(string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                throw new ArgumentException("Sheet name is required.", nameof(sheetName));
            }

            EnsureWorkbookScope();
            var rows = GetTemplateBindingRows().ToList();
            var removed = rows.RemoveAll(row =>
                row.Length > 0 &&
                string.Equals(row[0], sheetName, StringComparison.OrdinalIgnoreCase));

            if (removed == 0)
            {
                return;
            }

            adapter.WriteTable(TemplateBindingsTableName, TemplateBindingHeaders, rows.ToArray());
            templateBindingRowsCache = CloneRows(rows);
            templateBindingRowsCacheLoaded = true;
        }

        public WorksheetSnapshotCell[] LoadSnapshot(string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                throw new ArgumentException("Sheet name is required.", nameof(sheetName));
            }

            EnsureWorkbookScope();
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

            EnsureWorkbookScope();
        }

        internal void InvalidateCache()
        {
            templateBindingRowsCache = null;
            templateBindingRowsCacheLoaded = false;
            bindingRowsCache = null;
            bindingRowsCacheLoaded = false;
            fieldMappingRowsCache = null;
            fieldMappingRowsCacheLoaded = false;
            fieldMappingHeaders = DefaultFieldMappingHeaders.ToArray();
        }

        private void EnsureWorkbookScope()
        {
            var currentWorkbookScopeKey = adapter.GetWorkbookScopeKey() ?? string.Empty;
            if (string.Equals(workbookScopeKey, currentWorkbookScopeKey, StringComparison.Ordinal))
            {
                return;
            }

            InvalidateCache();
            workbookScopeKey = currentWorkbookScopeKey;
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

        private static int? ParseNullableInt(IReadOnlyList<string> row, int index)
        {
            if (row == null || row.Count <= index || string.IsNullOrWhiteSpace(row[index]))
            {
                return null;
            }

            return int.TryParse(row[index], NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed)
                ? parsed
                : (int?)null;
        }

        private static DateTime? ParseNullableDateTime(IReadOnlyList<string> row, int index)
        {
            if (row == null || row.Count <= index || string.IsNullOrWhiteSpace(row[index]))
            {
                return null;
            }

            return DateTime.TryParse(
                row[index],
                CultureInfo.InvariantCulture,
                DateTimeStyles.RoundtripKind,
                out var parsed)
                ? parsed
                : (DateTime?)null;
        }

        private static SheetTemplateBinding ParseTemplateBindingRow(IReadOnlyList<string> row)
        {
            if (row == null || row.Count == 0 || string.IsNullOrWhiteSpace(row[0]))
            {
                return null;
            }

            return new SheetTemplateBinding
            {
                SheetName = row[0],
                TemplateName = row.Count > 1 ? row[1] : string.Empty,
                TemplateRevision = ParseNullableInt(row, 2),
                TemplateOrigin = row.Count > 3 ? row[3] : string.Empty,
                TemplateId = row.Count > 4 ? row[4] : string.Empty,
                AppliedFingerprint = row.Count > 5 ? row[5] : string.Empty,
                TemplateLastAppliedAt = ParseNullableDateTime(row, 6),
                DerivedFromTemplateId = row.Count > 7 ? row[7] : string.Empty,
                DerivedFromTemplateRevision = ParseNullableInt(row, 8),
            };
        }

        private static SheetBinding ParseBindingRow(IReadOnlyList<string> row)
        {
            if (row == null || row.Count < 4 || string.IsNullOrWhiteSpace(row[0]))
            {
                return null;
            }

            return new SheetBinding
            {
                SheetName = row[0],
                ProjectId = row[1],
                ProjectName = row[2],
                HeaderStartRow = ParseIntOrDefault(row, 3, defaultValue: 1),
                HeaderRowCount = ParseIntOrDefault(row, 4, defaultValue: 2),
                DataStartRow = ParseIntOrDefault(row, 5, defaultValue: 3),
                SystemKey = row.Count > 6 ? row[6] : string.Empty,
            };
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

        private static FieldMappingColumnDefinition[] GetDistinctFieldMappingColumns(FieldMappingTableDefinition definition)
        {
            var columns = GetValidatedColumns(definition);
            var distinct = new List<FieldMappingColumnDefinition>(columns.Length);
            var seenKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var column in columns)
            {
                var valueKey = GetColumnValueKey(column);
                if (!seenKeys.Add(valueKey))
                {
                    continue;
                }

                distinct.Add(column);
            }

            return distinct.ToArray();
        }

        private static string GetColumnValueKey(FieldMappingColumnDefinition column)
        {
            if (column == null)
            {
                return string.Empty;
            }

            return string.IsNullOrWhiteSpace(column.RoleKey)
                ? column.ColumnName ?? string.Empty
                : column.RoleKey;
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

        private IReadOnlyList<string[]> GetTemplateBindingRows()
        {
            if (!templateBindingRowsCacheLoaded)
            {
                templateBindingRowsCache = adapter.ReadTable(TemplateBindingsTableName) ?? Array.Empty<string[]>();
                templateBindingRowsCacheLoaded = true;
            }

            return templateBindingRowsCache ?? Array.Empty<string[]>();
        }

        private IReadOnlyList<string[]> GetBindingRows()
        {
            if (!bindingRowsCacheLoaded)
            {
                bindingRowsCache = adapter.ReadTable(BindingsTableName) ?? Array.Empty<string[]>();
                bindingRowsCacheLoaded = true;
            }

            return bindingRowsCache ?? Array.Empty<string[]>();
        }

        private IReadOnlyList<string[]> GetFieldMappingRows()
        {
            if (!fieldMappingRowsCacheLoaded)
            {
                fieldMappingRowsCache = adapter.ReadTable(FieldMappingsTableName) ?? Array.Empty<string[]>();
                fieldMappingRowsCacheLoaded = true;
            }

            return fieldMappingRowsCache ?? Array.Empty<string[]>();
        }

        private static string[][] CloneRows(IEnumerable<string[]> rows)
        {
            return (rows ?? Array.Empty<string[]>())
                .Select(row => row?.ToArray() ?? Array.Empty<string>())
                .ToArray();
        }
    }
}
