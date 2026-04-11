using System;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Sync;
using Xunit;

namespace OfficeAgent.Core.Tests
{
    public sealed class WorksheetChangeTrackerTests
    {
        [Fact]
        public void GetDirtyCellsReturnsOnlyChangedCellsForExistingIds()
        {
            var tracker = new WorksheetChangeTracker();

            var dirty = tracker.GetDirtyCells(
                "Sync-performance",
                new[]
                {
                    new WorksheetSnapshotCell { SheetName = "Sync-performance", RowId = "row-1", ApiFieldKey = "name", Value = "旧值" },
                    new WorksheetSnapshotCell { SheetName = "Sync-performance", RowId = "row-1", ApiFieldKey = "start_12345678", Value = "2026-01-01" },
                },
                new[]
                {
                    new CellChange { SheetName = "Sync-performance", RowId = "row-1", ApiFieldKey = "name", OldValue = "旧值", NewValue = "新值" },
                    new CellChange { SheetName = "Sync-performance", RowId = "row-1", ApiFieldKey = "start_12345678", OldValue = "2026-01-01", NewValue = "2026-01-01" },
                });

            var changed = Assert.Single(dirty);
            Assert.Equal("name", changed.ApiFieldKey);
            Assert.Equal("新值", changed.NewValue);
        }
    }
}
