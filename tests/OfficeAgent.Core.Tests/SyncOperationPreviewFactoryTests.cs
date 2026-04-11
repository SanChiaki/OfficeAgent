using OfficeAgent.Core.Models;
using OfficeAgent.Core.Sync;
using Xunit;

namespace OfficeAgent.Core.Tests
{
    public sealed class SyncOperationPreviewFactoryTests
    {
        [Fact]
        public void CreateUploadPreviewSummarizesChanges()
        {
            var factory = new SyncOperationPreviewFactory();
            var changes = new[]
            {
                new CellChange { RowId = "row-1", ApiFieldKey = "name", OldValue = "A", NewValue = "B" },
                new CellChange { RowId = "row-2", ApiFieldKey = "value", OldValue = "1", NewValue = "2" },
                new CellChange { RowId = "row-3", ApiFieldKey = "status", OldValue = "off", NewValue = "on" },
                new CellChange { RowId = "row-4", ApiFieldKey = "extra", OldValue = "X", NewValue = "Y" },
            };

            var preview = factory.CreateUploadPreview("UploadTest", changes);

            Assert.Equal("UploadTest", preview.OperationName);
            Assert.Equal("Upload 4 changed cell(s).", preview.Summary);
            Assert.Equal(3, preview.Details.Length);
            Assert.StartsWith("row-1 / name: A -> B", preview.Details[0]);
            Assert.Equal(4, preview.Changes.Length);
            Assert.Equal("row-4", preview.Changes[3].RowId);
        }
    }
}
