using System.Collections.Generic;
using System.Linq;
using OfficeAgent.Core.Models;
using OfficeAgent.Infrastructure.Http;
using Xunit;

namespace OfficeAgent.Infrastructure.Tests
{
    public sealed class CurrentBusinessSchemaMapperTests
    {
        [Fact]
        public void BuildExpandsActivityFieldsIntoMixedColumns()
        {
            var mapper = new CurrentBusinessSchemaMapper(
                new Dictionary<string, string>(System.StringComparer.OrdinalIgnoreCase)
                {
                    ["start"] = "开始时间",
                    ["end"] = "结束时间",
                });

            var schema = mapper.Build(
                "performance",
                new[]
                {
                    new CurrentBusinessHeadDefinition { FieldKey = "id", HeaderText = "ID", IsId = true },
                    new CurrentBusinessHeadDefinition { FieldKey = "name", HeaderText = "项目名称" },
                    new CurrentBusinessHeadDefinition { HeadType = "activity", ActivityId = "12345678", ActivityName = "测试活动111" },
                },
                new[]
                {
                    new Dictionary<string, object>
                    {
                        ["id"] = "row-1",
                        ["name"] = "项目A",
                        ["start_12345678"] = "2026-01-02",
                        ["end_12345678"] = "2026-01-03",
                    },
                });

            Assert.Collection(
                schema.Columns,
                column => Assert.Equal("id", column.ApiFieldKey),
                column => Assert.Equal("name", column.ApiFieldKey),
                column =>
                {
                    Assert.Equal("end_12345678", column.ApiFieldKey);
                    Assert.Equal(WorksheetColumnKind.ActivityProperty, column.ColumnKind);
                    Assert.Equal("测试活动111", column.ParentHeaderText);
                    Assert.Equal("结束时间", column.ChildHeaderText);
                },
                column =>
                {
                    Assert.Equal("start_12345678", column.ApiFieldKey);
                    Assert.Equal(WorksheetColumnKind.ActivityProperty, column.ColumnKind);
                    Assert.Equal("测试活动111", column.ParentHeaderText);
                    Assert.Equal("开始时间", column.ChildHeaderText);
                });
        }

        [Fact]
        public void BuildHandlesPropertiesWithUnderscoresAndOrdersColumns()
        {
            var mapper = new CurrentBusinessSchemaMapper(
                new Dictionary<string, string>(System.StringComparer.OrdinalIgnoreCase)
                {
                    ["foo_bar"] = "FooBar",
                    ["alpha"] = "Alpha",
                });

            var schema = mapper.Build(
                "performance",
                new[]
                {
                    new CurrentBusinessHeadDefinition { FieldKey = "id", HeaderText = "ID", IsId = true },
                    new CurrentBusinessHeadDefinition { HeadType = "activity", ActivityId = "12345678", ActivityName = "活动顺序" },
                },
                new[]
                {
                    new Dictionary<string, object>
                    {
                        ["id"] = "row-1",
                        ["foo_bar_12345678"] = "value-1",
                        ["alpha_12345678"] = "value-2",
                    },
                });

            var activityColumns = schema.Columns.Where(column => column.ColumnKind == WorksheetColumnKind.ActivityProperty);

            Assert.Collection(
                activityColumns,
                column =>
                {
                    Assert.Equal("alpha_12345678", column.ApiFieldKey);
                    Assert.Equal("Alpha", column.ChildHeaderText);
                    Assert.Equal("活动顺序", column.ParentHeaderText);
                },
                column =>
                {
                    Assert.Equal("foo_bar_12345678", column.ApiFieldKey);
                    Assert.Equal("FooBar", column.ChildHeaderText);
                });
        }
    }
}
