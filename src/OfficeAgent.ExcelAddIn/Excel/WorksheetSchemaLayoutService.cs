using System.Collections.Generic;
using System.Linq;
using OfficeAgent.Core.Models;

namespace OfficeAgent.ExcelAddIn.Excel
{
    internal sealed class WorksheetSchemaLayoutService
    {
        public HeaderCellPlan[] BuildHeaderPlan(WorksheetSchema schema)
        {
            var cells = new List<HeaderCellPlan>();

            foreach (var column in schema.Columns.Where(column => column.ColumnKind == WorksheetColumnKind.Single))
            {
                cells.Add(new HeaderCellPlan
                {
                    Row = 1,
                    Column = column.ColumnIndex,
                    RowSpan = 2,
                    Text = column.ChildHeaderText,
                });
            }

            var activityGroups = schema.Columns
                .Where(column => column.ColumnKind == WorksheetColumnKind.ActivityProperty)
                .GroupBy(column => GetActivityGroupKey(column))
                .OrderBy(group => group.Min(column => column.ColumnIndex));

            foreach (var group in activityGroups)
            {
                var ordered = group.OrderBy(column => column.ColumnIndex).ToArray();

                cells.Add(new HeaderCellPlan
                {
                    Row = 1,
                    Column = ordered[0].ColumnIndex,
                    ColumnSpan = ordered.Length,
                    Text = ordered[0].ParentHeaderText,
                });

                foreach (var column in ordered)
                {
                    cells.Add(new HeaderCellPlan
                    {
                        Row = 2,
                        Column = column.ColumnIndex,
                        Text = column.ChildHeaderText,
                    });
                }
            }

            return cells
                .OrderBy(cell => cell.Row)
                .ThenBy(cell => cell.Column)
                .ToArray();
        }

        private static string GetActivityGroupKey(WorksheetColumnBinding column)
        {
            if (!string.IsNullOrEmpty(column.ActivityId))
            {
                return $"id:{column.ActivityId}";
            }

            return $"text:{column.ParentHeaderText}";
        }
    }
}
