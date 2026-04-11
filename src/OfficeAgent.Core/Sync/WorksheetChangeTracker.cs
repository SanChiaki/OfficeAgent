using System;
using System.Collections.Generic;
using System.Linq;
using OfficeAgent.Core.Models;

namespace OfficeAgent.Core.Sync
{
    public sealed class WorksheetChangeTracker
    {
        public CellChange[] GetDirtyCells(
            string sheetName,
            IReadOnlyList<WorksheetSnapshotCell> snapshot,
            IReadOnlyList<CellChange> currentCells)
        {
            var baseline = snapshot.ToDictionary(
                item => $"{item.RowId}|{item.ApiFieldKey}",
                item => item.Value,
                StringComparer.Ordinal);

            return currentCells
                .Where(item => string.Equals(item.SheetName, sheetName, StringComparison.Ordinal))
                .Where(item =>
                    baseline.TryGetValue($"{item.RowId}|{item.ApiFieldKey}", out var oldValue) &&
                    !string.Equals(oldValue, item.NewValue, StringComparison.Ordinal))
                .ToArray();
        }
    }
}
