using System;
using System.Collections.Generic;
using System.Linq;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;
using OfficeAgent.Core.Templates;
using Xunit;

namespace OfficeAgent.Core.Tests
{
    public sealed class WorksheetTemplateCatalogTests
    {
        [Fact]
        public void ApplyTemplateToSheetWritesTemplateBindingAndInjectsCurrentSheetNameIntoMappings()
        {
            var metadataStore = new FakeWorksheetMetadataStore();
            var templateBindingStore = new FakeWorksheetTemplateBindingStore();
            var templateStore = new FakeTemplateStore();
            metadataStore.Bindings["Sheet-Now"] = CreateBinding("Sheet-Now", "system-a", "project-a", "项目A");
            var template = CreateTemplate("template-a", "模板A", "system-a", "project-a");
            templateStore.Templates[template.TemplateId] = template;
            var catalog = CreateCatalog(metadataStore, templateBindingStore, templateStore);

            catalog.ApplyTemplateToSheet("Sheet-Now", "template-a");

            var savedTemplateBinding = templateBindingStore.TemplateBindings["Sheet-Now"];
            Assert.Equal("store-template", savedTemplateBinding.TemplateOrigin);
            Assert.Equal("template-a", savedTemplateBinding.TemplateId);
            Assert.Equal(template.Revision, savedTemplateBinding.TemplateRevision);
            Assert.Equal("Sheet-Now", metadataStore.LastSavedBinding.SheetName);
            Assert.Equal("Sheet-Now", metadataStore.LastSavedFieldMappings[0].SheetName);
            Assert.Equal("owner_name", metadataStore.LastSavedFieldMappings[0].Values["ApiFieldKey"]);
        }

        [Fact]
        public void SaveSheetAsNewTemplateCreatesSheetAgnosticTemplateAndBindsCurrentSheet()
        {
            var metadataStore = new FakeWorksheetMetadataStore();
            var templateBindingStore = new FakeWorksheetTemplateBindingStore();
            var templateStore = new FakeTemplateStore();
            metadataStore.Bindings["Sheet-Now"] = CreateBinding("Sheet-Now", "system-a", "project-a", "项目A");
            metadataStore.FieldMappings["Sheet-Now"] = CreateFieldMappings("Sheet-Now");
            templateBindingStore.TemplateBindings["Sheet-Now"] = new SheetTemplateBinding
            {
                SheetName = "Sheet-Now",
                TemplateId = "template-old",
                TemplateName = "旧模板",
                TemplateRevision = 5,
                TemplateOrigin = "store-template",
            };
            var catalog = CreateCatalog(metadataStore, templateBindingStore, templateStore);

            catalog.SaveSheetAsNewTemplate("Sheet-Now", "模板新版本");

            var savedTemplate = Assert.Single(templateStore.SaveNewCalls);
            Assert.Equal("模板新版本", savedTemplate.TemplateName);
            Assert.DoesNotContain(savedTemplate.FieldMappings, row => row.Values.ContainsKey("SheetName"));
            var savedTemplateBinding = templateBindingStore.TemplateBindings["Sheet-Now"];
            Assert.Equal("store-template", savedTemplateBinding.TemplateOrigin);
            Assert.Equal(savedTemplate.TemplateId, savedTemplateBinding.TemplateId);
            Assert.Equal("template-old", savedTemplateBinding.DerivedFromTemplateId);
            Assert.Equal(5, savedTemplateBinding.DerivedFromTemplateRevision);
        }

        [Fact]
        public void GetSheetStateMarksDirtyWhenFingerprintDiffersFromAppliedFingerprint()
        {
            var metadataStore = new FakeWorksheetMetadataStore();
            var templateBindingStore = new FakeWorksheetTemplateBindingStore();
            var templateStore = new FakeTemplateStore();
            metadataStore.Bindings["Sheet-Now"] = CreateBinding("Sheet-Now", "system-a", "project-a", "项目A");
            metadataStore.FieldMappings["Sheet-Now"] = CreateFieldMappings("Sheet-Now");
            templateBindingStore.TemplateBindings["Sheet-Now"] = new SheetTemplateBinding
            {
                SheetName = "Sheet-Now",
                TemplateId = "template-a",
                TemplateName = "模板A",
                TemplateRevision = 1,
                TemplateOrigin = "store-template",
                AppliedFingerprint = "fingerprint-old",
            };
            templateStore.Templates["template-a"] = CreateTemplate("template-a", "模板A", "system-a", "project-a");
            var catalog = CreateCatalog(metadataStore, templateBindingStore, templateStore);

            var state = catalog.GetSheetState("Sheet-Now");

            Assert.True(state.IsDirty);
        }

        [Fact]
        public void SaveSheetToExistingTemplateRejectsRevisionMismatchWhenOverwriteDisabled()
        {
            var metadataStore = new FakeWorksheetMetadataStore();
            var templateBindingStore = new FakeWorksheetTemplateBindingStore();
            var templateStore = new FakeTemplateStore();
            metadataStore.Bindings["Sheet-Now"] = CreateBinding("Sheet-Now", "system-a", "project-a", "项目A");
            metadataStore.FieldMappings["Sheet-Now"] = CreateFieldMappings("Sheet-Now");
            templateBindingStore.TemplateBindings["Sheet-Now"] = new SheetTemplateBinding
            {
                SheetName = "Sheet-Now",
                TemplateId = "template-a",
                TemplateName = "模板A",
                TemplateRevision = 1,
                TemplateOrigin = "store-template",
            };
            templateStore.Templates["template-a"] = CreateTemplate("template-a", "模板A", "system-a", "project-a", 2);
            var catalog = CreateCatalog(metadataStore, templateBindingStore, templateStore);

            var exception = Assert.Throws<InvalidOperationException>(() =>
                catalog.SaveSheetToExistingTemplate("Sheet-Now", "template-a", 1, false));

            Assert.Contains("revision", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Empty(templateStore.SaveExistingCalls);
        }

        [Fact]
        public void GetSheetStateMarksTemplateMissingWhenReferencedTemplateNoLongerExists()
        {
            var metadataStore = new FakeWorksheetMetadataStore();
            var templateBindingStore = new FakeWorksheetTemplateBindingStore();
            var templateStore = new FakeTemplateStore();
            metadataStore.Bindings["Sheet-Now"] = CreateBinding("Sheet-Now", "system-a", "project-a", "项目A");
            metadataStore.FieldMappings["Sheet-Now"] = CreateFieldMappings("Sheet-Now");
            templateBindingStore.TemplateBindings["Sheet-Now"] = new SheetTemplateBinding
            {
                SheetName = "Sheet-Now",
                TemplateId = "template-missing",
                TemplateName = "丢失模板",
                TemplateRevision = 7,
                TemplateOrigin = "store-template",
                AppliedFingerprint = "fingerprint-old",
            };
            var catalog = CreateCatalog(metadataStore, templateBindingStore, templateStore);

            var state = catalog.GetSheetState("Sheet-Now");

            Assert.True(state.TemplateMissing);
            Assert.True(state.CanApplyTemplate);
            Assert.True(state.CanSaveAsTemplate);
            Assert.False(state.CanSaveTemplate);
            Assert.Equal(7, state.TemplateRevision);
            Assert.Null(state.StoredTemplateRevision);
        }

        [Fact]
        public void ApplyTemplateToSheetRejectsIncompatibleTemplateDefinition()
        {
            var metadataStore = new FakeWorksheetMetadataStore();
            var templateBindingStore = new FakeWorksheetTemplateBindingStore();
            var templateStore = new FakeTemplateStore();
            metadataStore.Bindings["Sheet-Now"] = CreateBinding("Sheet-Now", "system-a", "project-a", "项目A");
            var incompatibleTemplate = CreateTemplate("template-a", "模板A", "system-a", "project-a");
            incompatibleTemplate.FieldMappingDefinition = new FieldMappingTableDefinition
            {
                SystemKey = "system-a",
                Columns = new[]
                {
                    new FieldMappingColumnDefinition
                    {
                        ColumnName = "LegacyField",
                        Role = FieldMappingSemanticRole.ApiFieldKey,
                    },
                },
            };
            incompatibleTemplate.FieldMappingDefinitionFingerprint =
                TemplateFingerprintBuilder.BuildFieldMappingDefinitionFingerprint(incompatibleTemplate.FieldMappingDefinition);
            templateStore.Templates[incompatibleTemplate.TemplateId] = incompatibleTemplate;
            var catalog = CreateCatalog(metadataStore, templateBindingStore, templateStore);

            var exception = Assert.Throws<InvalidOperationException>(() =>
                catalog.ApplyTemplateToSheet("Sheet-Now", incompatibleTemplate.TemplateId));

            Assert.Contains("compatible", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void GetSheetStateDisablesSaveAndAvoidsDirtyForIncompatibleTemplateDefinition()
        {
            var metadataStore = new FakeWorksheetMetadataStore();
            var templateBindingStore = new FakeWorksheetTemplateBindingStore();
            var templateStore = new FakeTemplateStore();
            metadataStore.Bindings["Sheet-Now"] = CreateBinding("Sheet-Now", "system-a", "project-a", "项目A");
            metadataStore.FieldMappings["Sheet-Now"] = CreateFieldMappings("Sheet-Now");
            templateBindingStore.TemplateBindings["Sheet-Now"] = new SheetTemplateBinding
            {
                SheetName = "Sheet-Now",
                TemplateId = "template-a",
                TemplateName = "模板A",
                TemplateRevision = 1,
                TemplateOrigin = "store-template",
                AppliedFingerprint = "outdated-fingerprint",
            };
            var incompatibleTemplate = CreateTemplate("template-a", "模板A", "system-a", "project-a");
            incompatibleTemplate.FieldMappingDefinition = new FieldMappingTableDefinition
            {
                SystemKey = "system-a",
                Columns = new[]
                {
                    new FieldMappingColumnDefinition
                    {
                        ColumnName = "LegacyField",
                        Role = FieldMappingSemanticRole.ApiFieldKey,
                    },
                },
            };
            incompatibleTemplate.FieldMappingDefinitionFingerprint =
                TemplateFingerprintBuilder.BuildFieldMappingDefinitionFingerprint(incompatibleTemplate.FieldMappingDefinition);
            templateStore.Templates[incompatibleTemplate.TemplateId] = incompatibleTemplate;
            var catalog = CreateCatalog(metadataStore, templateBindingStore, templateStore);

            var state = catalog.GetSheetState("Sheet-Now");

            Assert.False(state.CanSaveTemplate);
            Assert.False(state.IsDirty);
            Assert.False(state.TemplateMissing);
        }

        [Fact]
        public void SaveSheetToExistingTemplatePassesExpectedRevisionToStore()
        {
            var metadataStore = new FakeWorksheetMetadataStore();
            var templateBindingStore = new FakeWorksheetTemplateBindingStore();
            var templateStore = new FakeTemplateStore();
            metadataStore.Bindings["Sheet-Now"] = CreateBinding("Sheet-Now", "system-a", "project-a", "项目A");
            metadataStore.FieldMappings["Sheet-Now"] = CreateFieldMappings("Sheet-Now");
            templateBindingStore.TemplateBindings["Sheet-Now"] = new SheetTemplateBinding
            {
                SheetName = "Sheet-Now",
                TemplateId = "template-a",
                TemplateName = "模板A",
                TemplateRevision = 1,
                TemplateOrigin = "store-template",
            };
            templateStore.Templates["template-a"] = CreateTemplate("template-a", "模板A", "system-a", "project-a", 1);
            var catalog = CreateCatalog(metadataStore, templateBindingStore, templateStore);

            catalog.SaveSheetToExistingTemplate("Sheet-Now", "template-a", 1, false);

            var call = Assert.Single(templateStore.SaveExistingCalls);
            Assert.Equal(1, call.ExpectedRevision);
        }

        [Fact]
        public void SaveSheetToExistingTemplateUsesLatestStoredRevisionWhenOverwriteEnabled()
        {
            var metadataStore = new FakeWorksheetMetadataStore();
            var templateBindingStore = new FakeWorksheetTemplateBindingStore();
            var templateStore = new FakeTemplateStore();
            metadataStore.Bindings["Sheet-Now"] = CreateBinding("Sheet-Now", "system-a", "project-a", "项目A");
            metadataStore.FieldMappings["Sheet-Now"] = CreateFieldMappings("Sheet-Now");
            templateBindingStore.TemplateBindings["Sheet-Now"] = new SheetTemplateBinding
            {
                SheetName = "Sheet-Now",
                TemplateId = "template-a",
                TemplateName = "模板A",
                TemplateRevision = 1,
                TemplateOrigin = "store-template",
            };
            templateStore.Templates["template-a"] = CreateTemplate("template-a", "模板A", "system-a", "project-a", 4);
            var catalog = CreateCatalog(metadataStore, templateBindingStore, templateStore);

            catalog.SaveSheetToExistingTemplate("Sheet-Now", "template-a", 1, true);

            var call = Assert.Single(templateStore.SaveExistingCalls);
            Assert.Equal(4, call.ExpectedRevision);
        }

        [Fact]
        public void SaveSheetToExistingTemplateRejectsIncompatibleTemplateDefinition()
        {
            var metadataStore = new FakeWorksheetMetadataStore();
            var templateBindingStore = new FakeWorksheetTemplateBindingStore();
            var templateStore = new FakeTemplateStore();
            metadataStore.Bindings["Sheet-Now"] = CreateBinding("Sheet-Now", "system-a", "project-a", "项目A");
            metadataStore.FieldMappings["Sheet-Now"] = CreateFieldMappings("Sheet-Now");
            templateBindingStore.TemplateBindings["Sheet-Now"] = new SheetTemplateBinding
            {
                SheetName = "Sheet-Now",
                TemplateId = "template-a",
                TemplateName = "模板A",
                TemplateRevision = 1,
                TemplateOrigin = "store-template",
            };
            var incompatibleTemplate = CreateTemplate("template-a", "模板A", "system-a", "project-a", 1);
            incompatibleTemplate.FieldMappingDefinition = new FieldMappingTableDefinition
            {
                SystemKey = "system-a",
                Columns = new[]
                {
                    new FieldMappingColumnDefinition
                    {
                        ColumnName = "LegacyField",
                        Role = FieldMappingSemanticRole.ApiFieldKey,
                    },
                },
            };
            incompatibleTemplate.FieldMappingDefinitionFingerprint =
                TemplateFingerprintBuilder.BuildFieldMappingDefinitionFingerprint(incompatibleTemplate.FieldMappingDefinition);
            templateStore.Templates["template-a"] = incompatibleTemplate;
            var catalog = CreateCatalog(metadataStore, templateBindingStore, templateStore);

            var exception = Assert.Throws<InvalidOperationException>(() =>
                catalog.SaveSheetToExistingTemplate("Sheet-Now", "template-a", 1, false));

            Assert.Contains("compatible", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Empty(templateStore.SaveExistingCalls);
        }

        private static WorksheetTemplateCatalog CreateCatalog(
            FakeWorksheetMetadataStore metadataStore,
            FakeWorksheetTemplateBindingStore templateBindingStore,
            FakeTemplateStore templateStore)
        {
            return new WorksheetTemplateCatalog(
                new SystemConnectorRegistry(new[] { new FakeSystemConnector("system-a") }),
                metadataStore,
                templateBindingStore,
                templateStore);
        }

        private static SheetBinding CreateBinding(string sheetName, string systemKey, string projectId, string projectName)
        {
            return new SheetBinding
            {
                SheetName = sheetName,
                SystemKey = systemKey,
                ProjectId = projectId,
                ProjectName = projectName,
                HeaderStartRow = 1,
                HeaderRowCount = 2,
                DataStartRow = 3,
            };
        }

        private static SheetFieldMappingRow[] CreateFieldMappings(string sheetName)
        {
            return new[]
            {
                new SheetFieldMappingRow
                {
                    SheetName = sheetName,
                    Values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["ApiFieldKey"] = "owner_name",
                        ["HeaderType"] = "single",
                        ["Excel L1"] = "负责人",
                    },
                },
            };
        }

        private static TemplateDefinition CreateTemplate(
            string templateId,
            string templateName,
            string systemKey,
            string projectId,
            int revision = 1)
        {
            var template = new TemplateDefinition
            {
                TemplateId = templateId,
                TemplateName = templateName,
                SystemKey = systemKey,
                ProjectId = projectId,
                ProjectName = "项目A",
                HeaderStartRow = 1,
                HeaderRowCount = 2,
                DataStartRow = 3,
                FieldMappingDefinition = new FieldMappingTableDefinition
                {
                    SystemKey = systemKey,
                    Columns = new[]
                    {
                        new FieldMappingColumnDefinition
                        {
                            ColumnName = "ApiFieldKey",
                            Role = FieldMappingSemanticRole.ApiFieldKey,
                        },
                        new FieldMappingColumnDefinition
                        {
                            ColumnName = "HeaderType",
                            Role = FieldMappingSemanticRole.HeaderType,
                        },
                        new FieldMappingColumnDefinition
                        {
                            ColumnName = "Excel L1",
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
                            ["ApiFieldKey"] = "owner_name",
                            ["HeaderType"] = "single",
                            ["Excel L1"] = "负责人",
                        },
                    },
                },
                Revision = revision,
                CreatedAtUtc = DateTime.UtcNow.AddDays(-1),
                UpdatedAtUtc = DateTime.UtcNow,
            };
            template.FieldMappingDefinitionFingerprint =
                TemplateFingerprintBuilder.BuildFieldMappingDefinitionFingerprint(template.FieldMappingDefinition);
            return template;
        }

        private sealed class FakeSystemConnector : ISystemConnector
        {
            public FakeSystemConnector(string systemKey)
            {
                SystemKey = systemKey;
                FieldMappingDefinition = new FieldMappingTableDefinition
                {
                    SystemKey = systemKey,
                    Columns = new[]
                    {
                        new FieldMappingColumnDefinition
                        {
                            ColumnName = "ApiFieldKey",
                            Role = FieldMappingSemanticRole.ApiFieldKey,
                        },
                        new FieldMappingColumnDefinition
                        {
                            ColumnName = "HeaderType",
                            Role = FieldMappingSemanticRole.HeaderType,
                        },
                        new FieldMappingColumnDefinition
                        {
                            ColumnName = "Excel L1",
                            Role = FieldMappingSemanticRole.CurrentSingleHeaderText,
                        },
                    },
                };
            }

            public string SystemKey { get; }

            public FieldMappingTableDefinition FieldMappingDefinition { get; }

            public IReadOnlyList<ProjectOption> GetProjects()
            {
                return Array.Empty<ProjectOption>();
            }

            public SheetBinding CreateBindingSeed(string sheetName, ProjectOption project)
            {
                throw new NotSupportedException();
            }

            public FieldMappingTableDefinition GetFieldMappingDefinition(string projectId)
            {
                return FieldMappingDefinition;
            }

            public IReadOnlyList<SheetFieldMappingRow> BuildFieldMappingSeed(string sheetName, string projectId)
            {
                throw new NotSupportedException();
            }

            public WorksheetSchema GetSchema(string projectId)
            {
                throw new NotSupportedException();
            }

            public IReadOnlyList<IDictionary<string, object>> Find(string projectId, IReadOnlyList<string> rowIds, IReadOnlyList<string> fieldKeys)
            {
                throw new NotSupportedException();
            }

            public void BatchSave(string projectId, IReadOnlyList<CellChange> changes)
            {
                throw new NotSupportedException();
            }
        }

        private sealed class FakeWorksheetMetadataStore : IWorksheetMetadataStore
        {
            public Dictionary<string, SheetBinding> Bindings { get; } = new Dictionary<string, SheetBinding>(StringComparer.OrdinalIgnoreCase);

            public Dictionary<string, SheetFieldMappingRow[]> FieldMappings { get; } = new Dictionary<string, SheetFieldMappingRow[]>(StringComparer.OrdinalIgnoreCase);

            public SheetBinding LastSavedBinding { get; private set; }

            public FieldMappingTableDefinition LastSavedFieldMappingDefinition { get; private set; }

            public SheetFieldMappingRow[] LastSavedFieldMappings { get; private set; } = Array.Empty<SheetFieldMappingRow>();

            public void SaveBinding(SheetBinding binding)
            {
                LastSavedBinding = binding;
                Bindings[binding.SheetName] = binding;
            }

            public SheetBinding LoadBinding(string sheetName)
            {
                if (!Bindings.TryGetValue(sheetName, out var binding))
                {
                    throw new InvalidOperationException("No binding.");
                }

                return binding;
            }

            public void SaveFieldMappings(string sheetName, FieldMappingTableDefinition definition, IReadOnlyList<SheetFieldMappingRow> rows)
            {
                LastSavedFieldMappingDefinition = definition;
                LastSavedFieldMappings = (rows ?? Array.Empty<SheetFieldMappingRow>()).ToArray();
                FieldMappings[sheetName] = LastSavedFieldMappings;
            }

            public SheetFieldMappingRow[] LoadFieldMappings(string sheetName, FieldMappingTableDefinition definition)
            {
                return FieldMappings.TryGetValue(sheetName, out var rows)
                    ? rows
                    : Array.Empty<SheetFieldMappingRow>();
            }

            public void ClearFieldMappings(string sheetName)
            {
                FieldMappings.Remove(sheetName);
            }

            public WorksheetSnapshotCell[] LoadSnapshot(string sheetName)
            {
                return Array.Empty<WorksheetSnapshotCell>();
            }

            public void SaveSnapshot(string sheetName, WorksheetSnapshotCell[] cells)
            {
            }
        }

        private sealed class FakeWorksheetTemplateBindingStore : IWorksheetTemplateBindingStore
        {
            public Dictionary<string, SheetTemplateBinding> TemplateBindings { get; } = new Dictionary<string, SheetTemplateBinding>(StringComparer.OrdinalIgnoreCase);

            public void SaveTemplateBinding(SheetTemplateBinding binding)
            {
                TemplateBindings[binding.SheetName] = binding;
            }

            public SheetTemplateBinding LoadTemplateBinding(string sheetName)
            {
                return TemplateBindings.TryGetValue(sheetName, out var binding)
                    ? binding
                    : null;
            }

            public void ClearTemplateBinding(string sheetName)
            {
                TemplateBindings.Remove(sheetName);
            }
        }

        private sealed class FakeTemplateStore : ITemplateStore
        {
            public sealed class SaveExistingCall
            {
                public TemplateDefinition Template { get; set; }

                public int ExpectedRevision { get; set; }
            }

            public Dictionary<string, TemplateDefinition> Templates { get; } = new Dictionary<string, TemplateDefinition>(StringComparer.OrdinalIgnoreCase);

            public List<TemplateDefinition> SaveNewCalls { get; } = new List<TemplateDefinition>();

            public List<SaveExistingCall> SaveExistingCalls { get; } = new List<SaveExistingCall>();

            public IReadOnlyList<TemplateDefinition> ListByProject(string systemKey, string projectId)
            {
                return Templates.Values
                    .Where(template => string.Equals(template.SystemKey, systemKey, StringComparison.Ordinal)
                        && string.Equals(template.ProjectId, projectId, StringComparison.Ordinal))
                    .ToArray();
            }

            public TemplateDefinition Load(string templateId)
            {
                return Templates.TryGetValue(templateId, out var template)
                    ? template
                    : null;
            }

            public TemplateDefinition SaveNew(TemplateDefinition template)
            {
                SaveNewCalls.Add(template);
                Templates[template.TemplateId] = template;
                return template;
            }

            public TemplateDefinition SaveExisting(TemplateDefinition template, int expectedRevision)
            {
                SaveExistingCalls.Add(
                    new SaveExistingCall
                    {
                        Template = template,
                        ExpectedRevision = expectedRevision,
                    });
                Templates[template.TemplateId] = template;
                return template;
            }
        }
    }
}
