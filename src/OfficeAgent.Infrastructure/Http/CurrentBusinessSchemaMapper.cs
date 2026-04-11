using System;
using System.Collections.Generic;
using System.Linq;
using OfficeAgent.Core.Models;

namespace OfficeAgent.Infrastructure.Http
{
    public sealed class CurrentBusinessSchemaMapper
    {
        private readonly IReadOnlyDictionary<string, string> propertyNames;

        public CurrentBusinessSchemaMapper(IReadOnlyDictionary<string, string> propertyNames)
        {
            this.propertyNames = propertyNames ?? throw new ArgumentNullException(nameof(propertyNames));
        }

        public WorksheetSchema Build(
            string projectId,
            IReadOnlyList<CurrentBusinessHeadDefinition> headList,
            IReadOnlyList<IDictionary<string, object>> rows)
        {
            if (headList is null)
            {
                throw new ArgumentNullException(nameof(headList));
            }

            if (rows is null)
            {
                throw new ArgumentNullException(nameof(rows));
            }

            var columns = new List<WorksheetColumnBinding>();
            var nextColumn = 1;

            foreach (var head in headList.Where((item) => !string.Equals(item.HeadType, "activity", StringComparison.OrdinalIgnoreCase)))
            {
                columns.Add(new WorksheetColumnBinding
                {
                    ColumnIndex = nextColumn++,
                    ApiFieldKey = head.FieldKey,
                    ColumnKind = WorksheetColumnKind.Single,
                    ParentHeaderText = head.HeaderText,
                    ChildHeaderText = head.HeaderText,
                    IsIdColumn = head.IsId,
                });
            }

            var activityHeads = headList
                .Where((item) => string.Equals(item.HeadType, "activity", StringComparison.OrdinalIgnoreCase))
                .ToDictionary((item) => item.ActivityId, StringComparer.OrdinalIgnoreCase);

            var flatKeys = rows.SelectMany((row) => row.Keys).Distinct(StringComparer.OrdinalIgnoreCase);
            var propertyColumns = new List<(string FlatKey, CurrentBusinessHeadDefinition Activity, string PropertyKey)>();

            foreach (var flatKey in flatKeys)
            {
                var lastIndex = flatKey.LastIndexOf('_');
                if (lastIndex <= 0 || lastIndex == flatKey.Length - 1)
                {
                    continue;
                }

                var propertyKey = flatKey.Substring(0, lastIndex);
                var activityId = flatKey.Substring(lastIndex + 1);

                if (!activityHeads.TryGetValue(activityId, out var activityHead))
                {
                    continue;
                }

                propertyColumns.Add((flatKey, activityHead, propertyKey));
            }

            var orderedProperties = propertyColumns
                .OrderBy(item => item.Activity.ActivityId, StringComparer.OrdinalIgnoreCase)
                .ThenBy(item => item.PropertyKey, StringComparer.OrdinalIgnoreCase);

            foreach (var candidate in orderedProperties)
            {
                var propertyKey = candidate.PropertyKey;
                columns.Add(new WorksheetColumnBinding
                {
                    ColumnIndex = nextColumn++,
                    ApiFieldKey = candidate.FlatKey,
                    ColumnKind = WorksheetColumnKind.ActivityProperty,
                    ParentHeaderText = candidate.Activity.ActivityName,
                    ChildHeaderText = propertyNames.TryGetValue(propertyKey, out var label) ? label : propertyKey,
                    ActivityId = candidate.Activity.ActivityId,
                    ActivityName = candidate.Activity.ActivityName,
                    PropertyKey = propertyKey,
                });
            }

            return new WorksheetSchema
            {
                SystemKey = "current-business-system",
                ProjectId = projectId,
                Columns = columns.ToArray(),
            };
        }
    }
}
