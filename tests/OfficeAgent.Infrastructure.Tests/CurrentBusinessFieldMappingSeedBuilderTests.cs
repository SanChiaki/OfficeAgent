using System;
using System.Collections.Generic;
using System.Linq;
using OfficeAgent.Core.Models;
using OfficeAgent.Infrastructure.Http;
using Xunit;

namespace OfficeAgent.Infrastructure.Tests
{
    public sealed class CurrentBusinessFieldMappingSeedBuilderTests
    {
        [Fact]
        public void BuildCreatesSingleAndActivityRowsUsingCurrentSystemColumnNames()
        {
            var builder = new CurrentBusinessFieldMappingSeedBuilder(new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["name"] = "名称",
                ["start"] = "开始时间",
                ["end"] = "结束时间",
            });

            var headList = new[]
            {
                null,
                new CurrentBusinessHeadDefinition
                {
                    FieldKey = "row_id",
                    HeaderText = "ID",
                    HeadType = "single",
                    IsId = true,
                },
                new CurrentBusinessHeadDefinition
                {
                    FieldKey = "owner_name",
                    HeaderText = "负责人",
                    HeadType = "single",
                    IsId = false,
                },
                new CurrentBusinessHeadDefinition
                {
                    HeadType = "activity",
                    ActivityId = "b2",
                    ActivityName = "活动B",
                },
                new CurrentBusinessHeadDefinition
                {
                    HeadType = "activity",
                    ActivityId = "a1",
                    ActivityName = "活动A",
                },
                new CurrentBusinessHeadDefinition
                {
                    HeadType = "activity",
                    ActivityId = "b2",
                    ActivityName = string.Empty,
                },
            };

            var sampleRows = new List<IDictionary<string, object>>
            {
                null,
                new Dictionary<string, object>
                {
                    ["start_b2"] = "2026-01-02",
                    ["end_b2"] = "2026-01-03",
                    ["name_a1"] = "Alpha",
                    ["custom_name_a1"] = "Custom",
                    ["ignored"] = "x",
                },
                new Dictionary<string, object>
                {
                    ["start_b2"] = "dup",
                },
            };

            var rows = builder.Build("Sheet1", headList, sampleRows).ToList();
            var headerIds = rows.Select(row => row.Values[CurrentBusinessFieldMappingColumns.HeaderId]).ToArray();

            Assert.Equal(
                new[] { "row_id", "owner_name", "custom_name_a1", "name_a1", "end_b2", "start_b2" },
                headerIds);

            var idRow = FindRow(rows, "row_id");
            Assert.Equal("Sheet1", idRow.SheetName);
            Assert.Equal(CurrentBusinessFieldMappingColumns.SingleHeaderType, idRow.Values[CurrentBusinessFieldMappingColumns.HeaderType]);
            Assert.Equal("row_id", idRow.Values[CurrentBusinessFieldMappingColumns.ApiFieldKey]);
            Assert.Equal("true", idRow.Values[CurrentBusinessFieldMappingColumns.IsIdColumn]);
            Assert.Equal("ID", idRow.Values[CurrentBusinessFieldMappingColumns.DefaultLevel1]);
            Assert.Equal("ID", idRow.Values[CurrentBusinessFieldMappingColumns.CurrentLevel1]);
            Assert.Equal(string.Empty, idRow.Values[CurrentBusinessFieldMappingColumns.DefaultLevel2]);
            Assert.Equal(string.Empty, idRow.Values[CurrentBusinessFieldMappingColumns.CurrentLevel2]);

            var activityRow = FindRow(rows, "start_b2");
            Assert.Equal(CurrentBusinessFieldMappingColumns.ActivityPropertyHeaderType, activityRow.Values[CurrentBusinessFieldMappingColumns.HeaderType]);
            Assert.Equal("start_b2", activityRow.Values[CurrentBusinessFieldMappingColumns.ApiFieldKey]);
            Assert.Equal("false", activityRow.Values[CurrentBusinessFieldMappingColumns.IsIdColumn]);
            Assert.Equal("活动B", activityRow.Values[CurrentBusinessFieldMappingColumns.DefaultLevel1]);
            Assert.Equal("活动B", activityRow.Values[CurrentBusinessFieldMappingColumns.CurrentLevel1]);
            Assert.Equal("开始时间", activityRow.Values[CurrentBusinessFieldMappingColumns.DefaultLevel2]);
            Assert.Equal("开始时间", activityRow.Values[CurrentBusinessFieldMappingColumns.CurrentLevel2]);
            Assert.Equal("b2", activityRow.Values[CurrentBusinessFieldMappingColumns.ActivityId]);
            Assert.Equal("start", activityRow.Values[CurrentBusinessFieldMappingColumns.PropertyId]);

            var customRow = FindRow(rows, "custom_name_a1");
            Assert.Equal("custom_name", customRow.Values[CurrentBusinessFieldMappingColumns.PropertyId]);
            Assert.Equal("custom_name", customRow.Values[CurrentBusinessFieldMappingColumns.DefaultLevel2]);
            Assert.Equal("活动A", customRow.Values[CurrentBusinessFieldMappingColumns.CurrentLevel1]);
            Assert.Equal("名称", FindRow(rows, "name_a1").Values[CurrentBusinessFieldMappingColumns.DefaultLevel2]);
        }

        [Fact]
        public void BuildTreatsNullSampleRowsAsEmpty()
        {
            var builder = new CurrentBusinessFieldMappingSeedBuilder(new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase));
            var headList = new[]
            {
                new CurrentBusinessHeadDefinition
                {
                    FieldKey = "row_id",
                    HeaderText = "ID",
                    HeadType = "single",
                    IsId = true,
                },
                new CurrentBusinessHeadDefinition
                {
                    HeadType = "activity",
                    ActivityId = "a1",
                    ActivityName = "活动A",
                },
            };

            var rows = builder.Build("Sheet1", headList, sampleRows: null).ToList();

            Assert.Single(rows);
            Assert.Equal("row_id", rows[0].Values[CurrentBusinessFieldMappingColumns.HeaderId]);
        }

        private static SheetFieldMappingRow FindRow(IEnumerable<SheetFieldMappingRow> rows, string headerId)
        {
            return rows.Single(row => string.Equals(row.Values[CurrentBusinessFieldMappingColumns.HeaderId], headerId, StringComparison.OrdinalIgnoreCase));
        }
    }
}
