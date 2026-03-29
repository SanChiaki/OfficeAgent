using Newtonsoft.Json;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;
using Xunit;

namespace OfficeAgent.Core.Tests
{
    public sealed class SelectionContextFactoryTests
    {
        [Fact]
        public void CreateBuildsHeaderPreviewAndSampleRowsForContiguousSelections()
        {
            var values = new[,]
            {
                { "Name", "Region", "Amount" },
                { "Project A", "CN", "42" },
                { "Project B", "US", "36" },
                { "Project C", "DE", "27" },
            };

            var context = SelectionContextFactory.Create(
                workbookName: "Quarterly Report.xlsx",
                sheetName: "Sheet1",
                address: "A1:C4",
                rowCount: 4,
                columnCount: 3,
                areaCount: 1,
                previewValues: values);

            Assert.True(context.HasSelection);
            Assert.True(context.IsContiguous);
            Assert.Null(context.WarningMessage);
            Assert.Equal(new[] { "Name", "Region", "Amount" }, context.HeaderPreview);
            Assert.Equal(3, context.SampleRows.Length);
            Assert.Equal(new[] { "Project A", "CN", "42" }, context.SampleRows[0]);
        }

        [Fact]
        public void CreateMarksMultiAreaSelectionsAsUnsupported()
        {
            var context = SelectionContextFactory.Create(
                workbookName: "Quarterly Report.xlsx",
                sheetName: "Sheet2",
                address: "A1:A2,C1:C2",
                rowCount: 4,
                columnCount: 2,
                areaCount: 2,
                previewValues: null);

            Assert.True(context.HasSelection);
            Assert.False(context.IsContiguous);
            Assert.Equal("Multiple selection areas are not supported yet.", context.WarningMessage);
            Assert.Empty(context.HeaderPreview);
            Assert.Empty(context.SampleRows);
        }

        [Fact]
        public void SelectionContextSerializesWithCamelCaseProperties()
        {
            var context = SelectionContextFactory.Create(
                workbookName: "Quarterly Report.xlsx",
                sheetName: "Sheet1",
                address: "A1:C4",
                rowCount: 4,
                columnCount: 3,
                areaCount: 1,
                previewValues: new[,]
                {
                    { "Name", "Region", "Amount" },
                    { "Project A", "CN", "42" },
                });

            var json = JsonConvert.SerializeObject(context);

            Assert.Contains("\"hasSelection\":true", json);
            Assert.Contains("\"workbookName\":\"Quarterly Report.xlsx\"", json);
            Assert.DoesNotContain("\"WorkbookName\":", json);
        }
    }
}
