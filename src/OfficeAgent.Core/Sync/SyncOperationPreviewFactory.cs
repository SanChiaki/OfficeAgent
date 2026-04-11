using System;
using System.Collections.Generic;
using System.Linq;
using OfficeAgent.Core.Models;

namespace OfficeAgent.Core.Sync
{
    public sealed class SyncOperationPreviewFactory
    {
        public SyncOperationPreview CreateUploadPreview(string operationName, IReadOnlyList<CellChange> changes)
        {
            var changeList = changes ?? Array.Empty<CellChange>();

            var details = changeList
                .Take(3)
                .Select(item => $"{item.RowId} / {item.ApiFieldKey}: {item.OldValue} -> {item.NewValue}")
                .ToArray();

            return new SyncOperationPreview
            {
                OperationName = operationName ?? string.Empty,
                Summary = $"Upload {changeList.Count} changed cell(s).",
                Details = details,
                Changes = changeList.ToArray(),
            };
        }
    }
}
