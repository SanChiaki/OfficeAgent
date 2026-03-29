using System;
using System.Collections.Generic;
using OfficeAgent.Core.Models;

namespace OfficeAgent.Core.Services
{
    public static class SelectionContextFactory
    {
        private const int MaxPreviewColumns = 5;
        private const int MaxSampleRows = 3;

        public static SelectionContext Create(
            string workbookName,
            string sheetName,
            string address,
            int rowCount,
            int columnCount,
            int areaCount,
            string[,] previewValues)
        {
            var context = new SelectionContext
            {
                HasSelection = true,
                WorkbookName = workbookName ?? string.Empty,
                SheetName = sheetName ?? string.Empty,
                Address = address ?? string.Empty,
                RowCount = Math.Max(rowCount, 0),
                ColumnCount = Math.Max(columnCount, 0),
                IsContiguous = areaCount <= 1,
            };

            if (!context.IsContiguous)
            {
                context.WarningMessage = "Multiple selection areas are not supported yet.";
                return context;
            }

            if (previewValues == null || previewValues.Length == 0)
            {
                return context;
            }

            var previewColumnCount = Math.Min(previewValues.GetLength(1), MaxPreviewColumns);
            var previewRowCount = previewValues.GetLength(0);

            var headerPreview = new string[previewColumnCount];
            for (var columnIndex = 0; columnIndex < previewColumnCount; columnIndex++)
            {
                headerPreview[columnIndex] = previewValues[0, columnIndex] ?? string.Empty;
            }

            var sampleRows = new List<string[]>();
            var sampleRowCount = Math.Min(Math.Max(previewRowCount - 1, 0), MaxSampleRows);
            for (var rowIndex = 0; rowIndex < sampleRowCount; rowIndex++)
            {
                var sampleRow = new string[previewColumnCount];
                for (var columnIndex = 0; columnIndex < previewColumnCount; columnIndex++)
                {
                    sampleRow[columnIndex] = previewValues[rowIndex + 1, columnIndex] ?? string.Empty;
                }

                sampleRows.Add(sampleRow);
            }

            context.HeaderPreview = headerPreview;
            context.SampleRows = sampleRows.ToArray();
            return context;
        }
    }
}
