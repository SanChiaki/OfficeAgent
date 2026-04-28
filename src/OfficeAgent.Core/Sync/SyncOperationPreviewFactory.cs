using System;
using System.Collections.Generic;
using System.Linq;
using OfficeAgent.Core.Models;

namespace OfficeAgent.Core.Sync
{
    public sealed class SyncOperationPreviewFactory
    {
        public SyncOperationPreview CreateUploadPreview(
            string operationName,
            IReadOnlyList<CellChange> changes,
            IReadOnlyList<SkippedCellChange> skippedChanges = null)
        {
            var changeList = changes ?? Array.Empty<CellChange>();
            var skippedList = skippedChanges ?? Array.Empty<SkippedCellChange>();

            var uploadedDetails = changeList
                .Take(3)
                .Select(item => $"{item.RowId} / {item.ApiFieldKey}: {item.OldValue} -> {item.NewValue}")
                .ToArray();
            var skippedDetails = skippedList
                .Take(Math.Max(0, 10 - uploadedDetails.Length))
                .Select(item => $"{item.Change?.RowId ?? string.Empty} / {item.Change?.ApiFieldKey ?? string.Empty}: 已跳过，{item.Reason ?? string.Empty}")
                .ToArray();
            var details = uploadedDetails.Concat(skippedDetails).ToArray();
            var summary = skippedList.Count == 0
                ? $"Upload {changeList.Count} changed cell(s)."
                : $"{operationName ?? string.Empty}将上传 {changeList.Count} 个单元格，跳过 {skippedList.Count} 个单元格。";

            return new SyncOperationPreview
            {
                OperationName = operationName ?? string.Empty,
                Summary = summary,
                Details = details,
                Changes = changeList.ToArray(),
                SkippedChanges = skippedList.ToArray(),
            };
        }
    }
}
