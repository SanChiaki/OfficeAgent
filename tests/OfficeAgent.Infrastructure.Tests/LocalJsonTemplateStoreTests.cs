using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeAgent.Core.Models;
using OfficeAgent.Infrastructure.Storage;
using Xunit;

namespace OfficeAgent.Infrastructure.Tests
{
    public sealed class LocalJsonTemplateStoreTests : IDisposable
    {
        private readonly string tempDirectory;

        public LocalJsonTemplateStoreTests()
        {
            tempDirectory = Path.Combine(Path.GetTempPath(), "OfficeAgent.TemplateStore.Tests", Guid.NewGuid().ToString("N"));
        }

        [Fact]
        public void SaveNewAndLoadRoundTripTemplateDefinition()
        {
            var store = new LocalJsonTemplateStore(tempDirectory);
            var template = CreateTemplate("tpl-performance-a", "条件A", revision: 1);

            var saved = store.SaveNew(template);
            var loaded = store.Load(saved.TemplateId);

            Assert.NotNull(loaded);
            Assert.Equal("条件A", loaded.TemplateName);
            Assert.Equal("current-business-system", loaded.SystemKey);
            Assert.Equal("performance", loaded.ProjectId);
            Assert.Equal("owner_name", loaded.FieldMappingDefinition.Columns[0].ColumnName);
            Assert.Single(loaded.FieldMappings);
            Assert.Equal("负责人", loaded.FieldMappings[0].Values["currentsingledisplayname"]);
            Assert.True(File.Exists(Path.Combine(tempDirectory, "current-business-system", "performance", "tpl-performance-a.json")));
        }

        [Fact]
        public void LoadReturnsNullWhenTemplateDoesNotExist()
        {
            var store = new LocalJsonTemplateStore(tempDirectory);

            var template = store.Load("missing-template");

            Assert.Null(template);
        }

        [Fact]
        public void ListByProjectReturnsOnlyMatchingProjectTemplatesOrderedByUpdatedAtDescending()
        {
            var store = new LocalJsonTemplateStore(tempDirectory);
            store.SaveNew(CreateTemplate(
                "tpl-old",
                "条件旧",
                revision: 1,
                updatedAtUtc: new DateTime(2026, 4, 22, 8, 0, 0, DateTimeKind.Utc)));
            store.SaveNew(CreateTemplate(
                "tpl-new",
                "条件新",
                revision: 1,
                updatedAtUtc: new DateTime(2026, 4, 22, 9, 0, 0, DateTimeKind.Utc)));
            store.SaveNew(CreateTemplate(
                "tpl-other",
                "别的项目",
                revision: 1,
                projectId: "delivery-tracker",
                updatedAtUtc: new DateTime(2026, 4, 22, 10, 0, 0, DateTimeKind.Utc)));

            var templates = store.ListByProject("current-business-system", "performance");

            Assert.Equal(new[] { "tpl-new", "tpl-old" }, templates.Select(template => template.TemplateId).ToArray());
        }

        [Fact]
        public void SaveExistingRejectsRevisionMismatch()
        {
            var store = new LocalJsonTemplateStore(tempDirectory);
            store.SaveNew(CreateTemplate("tpl-performance-a", "条件A", revision: 1));

            var error = Assert.Throws<InvalidOperationException>(() =>
                store.SaveExisting(
                    CreateTemplate("tpl-performance-a", "条件A-更新", revision: 2),
                    expectedRevision: 0));

            Assert.Equal("模板版本已变化。", error.Message);
        }

        [Fact]
        public void SaveExistingOverwritesTemplateWhenExpectedRevisionMatches()
        {
            var store = new LocalJsonTemplateStore(tempDirectory);
            var createdAt = new DateTime(2026, 4, 22, 7, 0, 0, DateTimeKind.Utc);
            store.SaveNew(CreateTemplate("tpl-performance-a", "条件A", revision: 1, createdAtUtc: createdAt));

            var updated = store.SaveExisting(
                CreateTemplate(
                    "tpl-performance-a",
                    "条件A-更新",
                    revision: 2,
                    createdAtUtc: new DateTime(2030, 1, 1, 0, 0, 0, DateTimeKind.Utc),
                    updatedAtUtc: new DateTime(2026, 4, 22, 9, 30, 0, DateTimeKind.Utc)),
                expectedRevision: 1);

            var loaded = store.Load("tpl-performance-a");

            Assert.Equal("条件A-更新", updated.TemplateName);
            Assert.Equal(2, updated.Revision);
            Assert.NotNull(loaded);
            Assert.Equal("条件A-更新", loaded.TemplateName);
            Assert.Equal(2, loaded.Revision);
            Assert.Equal(createdAt, loaded.CreatedAtUtc);
        }

        public void Dispose()
        {
            if (Directory.Exists(tempDirectory))
            {
                Directory.Delete(tempDirectory, recursive: true);
            }
        }

        private static TemplateDefinition CreateTemplate(
            string templateId,
            string templateName,
            int revision,
            string projectId = "performance",
            DateTime? createdAtUtc = null,
            DateTime? updatedAtUtc = null)
        {
            return new TemplateDefinition
            {
                TemplateId = templateId,
                TemplateName = templateName,
                SystemKey = "current-business-system",
                ProjectId = projectId,
                ProjectName = projectId == "performance" ? "绩效项目" : "交付项目",
                HeaderStartRow = 1,
                HeaderRowCount = 2,
                DataStartRow = 3,
                FieldMappingDefinitionFingerprint = "fingerprint-" + projectId,
                FieldMappingDefinition = new FieldMappingTableDefinition
                {
                    SystemKey = "current-business-system",
                    Columns = new[]
                    {
                        new FieldMappingColumnDefinition
                        {
                            ColumnName = "owner_name",
                            Role = FieldMappingSemanticRole.CurrentSingleHeaderText,
                        },
                    },
                },
                FieldMappings = new[]
                {
                    new TemplateFieldMappingRow
                    {
                        Values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                        {
                            ["CurrentSingleDisplayName"] = "负责人",
                        },
                    },
                },
                Revision = revision,
                CreatedAtUtc = createdAtUtc ?? new DateTime(2026, 4, 22, 7, 0, 0, DateTimeKind.Utc),
                UpdatedAtUtc = updatedAtUtc ?? new DateTime(2026, 4, 22, 7, 0, 0, DateTimeKind.Utc),
            };
        }
    }
}
