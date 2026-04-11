using System;
using System.Collections.Generic;
using System.Linq;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;

namespace OfficeAgent.ExcelAddIn.Excel
{
    internal sealed class WorksheetMetadataStore : IWorksheetMetadataStore
    {
        private const string MetadataSheetName = "_OfficeAgentMetadata";
        private const string BindingsTableName = "SheetBindings";
        private const string SnapshotsTableName = "SheetSnapshots";

        private static readonly string[] BindingHeaders =
        {
            "SheetName",
            "SystemKey",
            "ProjectId",
            "ProjectName",
        };

        private static readonly string[] SnapshotHeaders =
        {
            "SheetName",
            "RowId",
            "ApiFieldKey",
            "Value",
        };

        private readonly IWorksheetMetadataAdapter adapter;

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

            adapter.EnsureWorksheet(MetadataSheetName, visible: true);
            var normalizedSheetName = binding.SheetName ?? string.Empty;
            var rows = adapter.ReadTable(BindingsTableName)?.ToList() ?? new List<string[]>();
            var newRow = new[]
            {
                normalizedSheetName,
                binding.SystemKey ?? string.Empty,
                binding.ProjectId ?? string.Empty,
                binding.ProjectName ?? string.Empty,
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

            adapter.EnsureWorksheet(MetadataSheetName, visible: true);
            var rows = adapter.ReadTable(BindingsTableName) ?? Array.Empty<string[]>();

            foreach (var row in rows)
            {
                if (row.Length < BindingHeaders.Length)
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
                };
            }

            throw new InvalidOperationException($"Binding for worksheet '{sheetName}' does not exist.");
        }

        public WorksheetSnapshotCell[] LoadSnapshot(string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                throw new ArgumentException("Sheet name is required.", nameof(sheetName));
            }

            adapter.EnsureWorksheet(MetadataSheetName, visible: true);
            var rows = adapter.ReadTable(SnapshotsTableName) ?? Array.Empty<string[]>();
            var result = new List<WorksheetSnapshotCell>();

            foreach (var row in rows)
            {
                if (row.Length < SnapshotHeaders.Length)
                {
                    continue;
                }

                if (!string.Equals(row[0], sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                result.Add(new WorksheetSnapshotCell
                {
                    SheetName = row[0],
                    RowId = row[1],
                    ApiFieldKey = row[2],
                    Value = row[3],
                });
            }

            return result.ToArray();
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

            adapter.EnsureWorksheet(MetadataSheetName, visible: true);
            var rows = adapter.ReadTable(SnapshotsTableName)?.ToList() ?? new List<string[]>();
            rows.RemoveAll(row =>
                row.Length > 0 &&
                string.Equals(row[0], sheetName, StringComparison.OrdinalIgnoreCase));

            var replacementRows = cells.Select(cell => new[]
            {
                sheetName,
                cell.RowId ?? string.Empty,
                cell.ApiFieldKey ?? string.Empty,
                cell.Value ?? string.Empty,
            });

            rows.AddRange(replacementRows);

            adapter.WriteTable(SnapshotsTableName, SnapshotHeaders, rows.ToArray());
        }
    }
}
