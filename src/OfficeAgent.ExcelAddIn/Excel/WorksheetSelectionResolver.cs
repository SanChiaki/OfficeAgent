using System;
using System.Collections.Generic;
using System.Linq;
using OfficeAgent.Core.Models;

namespace OfficeAgent.ExcelAddIn.Excel
{
    internal sealed class WorksheetSelectionResolver
    {
        public ResolvedSelection Resolve(
            WorksheetSchema schema,
            IReadOnlyList<SelectedVisibleCell> visibleCells,
            Func<int, string> rowIdAccessor)
        {
            if (schema is null)
            {
                throw new ArgumentNullException(nameof(schema));
            }

            if (visibleCells is null)
            {
                throw new ArgumentNullException(nameof(visibleCells));
            }

            if (rowIdAccessor is null)
            {
                throw new ArgumentNullException(nameof(rowIdAccessor));
            }

            var columns = schema.Columns ?? Array.Empty<WorksheetColumnBinding>();
            var columnsByIndex = new Dictionary<int, WorksheetColumnBinding>(columns.Length);
            foreach (var column in columns)
            {
                if (column is null)
                {
                    continue;
                }

                columnsByIndex[column.ColumnIndex] = column;
            }

            var resolvedCells = new List<(SelectedVisibleCell Cell, WorksheetColumnBinding Column, string RowId)>(visibleCells.Count);
            foreach (var cell in visibleCells)
            {
                if (cell is null)
                {
                    continue;
                }

                if (!columnsByIndex.TryGetValue(cell.Column, out var column))
                {
                    continue;
                }

                var rowId = rowIdAccessor(cell.Row);
                if (string.IsNullOrWhiteSpace(rowId))
                {
                    continue;
                }

                resolvedCells.Add((cell, column, rowId));
            }

            var rowIds = resolvedCells
                .Select(entry => entry.RowId)
                .Distinct(StringComparer.Ordinal)
                .ToArray();

            var apiFieldKeys = resolvedCells
                .Where(entry => !entry.Column.IsIdColumn)
                .Select(entry => entry.Column.ApiFieldKey)
                .Where(key => !string.IsNullOrWhiteSpace(key))
                .Distinct(StringComparer.Ordinal)
                .ToArray();

            return new ResolvedSelection
            {
                RowIds = rowIds,
                ApiFieldKeys = apiFieldKeys,
                TargetCells = resolvedCells.Select(entry => entry.Cell).ToArray(),
            };
        }
    }
}
