using System;
using System.IO;
using System.Linq;
using System.Reflection;
using OfficeAgent.Core.Models;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class WorksheetSchemaLayoutServiceTests
    {
        [Fact]
        public void BuildHeaderPlanMergesSingleColumnsVerticallyAndActivityColumnsHorizontally()
        {
            var service = CreateService();
            var schema = new WorksheetSchema
            {
                Columns = new[]
                {
                    new WorksheetColumnBinding { ColumnIndex = 1, ApiFieldKey = "id", ColumnKind = WorksheetColumnKind.Single, ParentHeaderText = "ID", ChildHeaderText = "ID", IsIdColumn = true },
                    new WorksheetColumnBinding { ColumnIndex = 2, ApiFieldKey = "name", ColumnKind = WorksheetColumnKind.Single, ParentHeaderText = "项目名称", ChildHeaderText = "项目名称" },
                    new WorksheetColumnBinding { ColumnIndex = 3, ApiFieldKey = "start_12345678", ColumnKind = WorksheetColumnKind.ActivityProperty, ParentHeaderText = "测试活动111", ChildHeaderText = "开始时间", ActivityId = "12345678" },
                    new WorksheetColumnBinding { ColumnIndex = 4, ApiFieldKey = "end_12345678", ColumnKind = WorksheetColumnKind.ActivityProperty, ParentHeaderText = "测试活动111", ChildHeaderText = "结束时间", ActivityId = "12345678" },
                },
            };

            var plan = BuildHeaderPlan(service, schema);

            Assert.Contains(plan, cell => cell.Row == 1 && cell.Column == 1 && cell.RowSpan == 2 && cell.Text == "ID");
            Assert.Contains(plan, cell => cell.Row == 1 && cell.Column == 3 && cell.ColumnSpan == 2 && cell.Text == "测试活动111");
            Assert.Contains(plan, cell => cell.Row == 2 && cell.Column == 4 && cell.Text == "结束时间");
        }

        [Fact]
        public void BuildHeaderPlanKeepsSeparateActivitiesEvenWithIdenticalParentText()
        {
            var service = CreateService();
            var schema = new WorksheetSchema
            {
                Columns = new[]
                {
                    new WorksheetColumnBinding { ColumnIndex = 1, ApiFieldKey = "id", ColumnKind = WorksheetColumnKind.Single, ParentHeaderText = "ID", ChildHeaderText = "ID", IsIdColumn = true },
                    new WorksheetColumnBinding { ColumnIndex = 2, ApiFieldKey = "name", ColumnKind = WorksheetColumnKind.Single, ParentHeaderText = "项目名称", ChildHeaderText = "项目名称" },
                    new WorksheetColumnBinding { ColumnIndex = 3, ApiFieldKey = "start_a", ColumnKind = WorksheetColumnKind.ActivityProperty, ParentHeaderText = "活动X", ChildHeaderText = "开始", ActivityId = "activity-1" },
                    new WorksheetColumnBinding { ColumnIndex = 4, ApiFieldKey = "end_a", ColumnKind = WorksheetColumnKind.ActivityProperty, ParentHeaderText = "活动X", ChildHeaderText = "结束", ActivityId = "activity-1" },
                    new WorksheetColumnBinding { ColumnIndex = 5, ApiFieldKey = "start_b", ColumnKind = WorksheetColumnKind.ActivityProperty, ParentHeaderText = "活动X", ChildHeaderText = "开始", ActivityId = "activity-2" },
                    new WorksheetColumnBinding { ColumnIndex = 6, ApiFieldKey = "end_b", ColumnKind = WorksheetColumnKind.ActivityProperty, ParentHeaderText = "活动X", ChildHeaderText = "结束", ActivityId = "activity-2" },
                },
            };

            var plan = BuildHeaderPlan(service, schema);

            var parentCells = plan
                .Where(cell => cell.Row == 1 && cell.Text == "活动X")
                .OrderBy(cell => cell.Column)
                .ToArray();

            Assert.Equal(2, parentCells.Length);
            Assert.Equal(3, parentCells[0].Column);
            Assert.Equal(2, parentCells[0].ColumnSpan);
            Assert.Equal(5, parentCells[1].Column);
            Assert.Equal(2, parentCells[1].ColumnSpan);
        }

        private static object CreateService()
        {
            var assembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var serviceType = assembly.GetType("OfficeAgent.ExcelAddIn.Excel.WorksheetSchemaLayoutService", throwOnError: true);
            var constructor = serviceType.GetConstructor(
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                binder: null,
                types: Type.EmptyTypes,
                modifiers: null);

            return constructor.Invoke(Array.Empty<object>());
        }

        private static HeaderCellPlan[] BuildHeaderPlan(object service, WorksheetSchema schema)
        {
            var method = service.GetType().GetMethod(
                "BuildHeaderPlan",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

            return (HeaderCellPlan[])method.Invoke(service, new object[] { schema })!;
        }

        private static string ResolveAddInAssemblyPath()
        {
            return Path.GetFullPath(
                Path.Combine(
                    AppContext.BaseDirectory,
                    "..",
                    "..",
                    "..",
                    "..",
                    "..",
                    "src",
                    "OfficeAgent.ExcelAddIn",
                    "bin",
                    "Debug",
                    "OfficeAgent.ExcelAddIn.dll"));
        }
    }
}
