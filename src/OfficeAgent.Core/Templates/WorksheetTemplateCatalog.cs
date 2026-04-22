using System;
using System.Collections.Generic;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;

namespace OfficeAgent.Core.Templates
{
    public sealed class WorksheetTemplateCatalog : ITemplateCatalog
    {
        private const string StoreTemplateOrigin = "store-template";
        private const string AdHocTemplateOrigin = "ad-hoc";

        private readonly ISystemConnectorRegistry systemConnectorRegistry;
        private readonly IWorksheetMetadataStore worksheetMetadataStore;
        private readonly IWorksheetTemplateBindingStore worksheetTemplateBindingStore;
        private readonly ITemplateStore templateStore;
        private readonly TemplateDefinitionNormalizer templateDefinitionNormalizer;
        private readonly TemplateFingerprintBuilder templateFingerprintBuilder;

        public WorksheetTemplateCatalog(
            ISystemConnectorRegistry systemConnectorRegistry,
            IWorksheetMetadataStore worksheetMetadataStore,
            IWorksheetTemplateBindingStore worksheetTemplateBindingStore,
            ITemplateStore templateStore)
        {
            this.systemConnectorRegistry = systemConnectorRegistry ?? throw new ArgumentNullException(nameof(systemConnectorRegistry));
            this.worksheetMetadataStore = worksheetMetadataStore ?? throw new ArgumentNullException(nameof(worksheetMetadataStore));
            this.worksheetTemplateBindingStore = worksheetTemplateBindingStore ?? throw new ArgumentNullException(nameof(worksheetTemplateBindingStore));
            this.templateStore = templateStore ?? throw new ArgumentNullException(nameof(templateStore));
            templateDefinitionNormalizer = new TemplateDefinitionNormalizer();
            templateFingerprintBuilder = new TemplateFingerprintBuilder();
        }

        public IReadOnlyList<TemplateDefinition> ListTemplates(string sheetName)
        {
            var binding = worksheetMetadataStore.LoadBinding(sheetName);
            return templateStore.ListByProject(binding.SystemKey, binding.ProjectId);
        }

        public SheetTemplateState GetSheetState(string sheetName)
        {
            var state = new SheetTemplateState();
            if (!TryLoadBinding(sheetName, out var binding))
            {
                return state;
            }

            state.HasProjectBinding = true;
            state.CanApplyTemplate = true;
            state.CanSaveAsTemplate = true;
            state.ProjectDisplayName = binding.ProjectName ?? string.Empty;

            if (!TryLoadTemplateBinding(sheetName, out var templateBinding))
            {
                return state;
            }

            state.TemplateId = templateBinding.TemplateId ?? string.Empty;
            state.TemplateName = templateBinding.TemplateName ?? string.Empty;
            state.TemplateRevision = templateBinding.TemplateRevision;
            state.TemplateOrigin = templateBinding.TemplateOrigin ?? string.Empty;

            if (!string.Equals(templateBinding.TemplateOrigin, StoreTemplateOrigin, StringComparison.Ordinal)
                || string.IsNullOrWhiteSpace(templateBinding.TemplateId))
            {
                return state;
            }

            var storedTemplate = templateStore.Load(templateBinding.TemplateId);
            if (storedTemplate == null)
            {
                state.TemplateMissing = true;
                state.CanSaveTemplate = false;
                state.IsDirty = IsDirty(sheetName, binding, templateBinding);
                return state;
            }

            state.StoredTemplateRevision = storedTemplate.Revision;
            state.TemplateMissing = false;
            var isCompatible = IsTemplateCompatibleWithCurrentDefinition(binding, storedTemplate);
            state.CanSaveTemplate = isCompatible;
            state.IsDirty = isCompatible && IsDirty(sheetName, binding, templateBinding);
            return state;
        }

        public void ApplyTemplateToSheet(string sheetName, string templateId)
        {
            var binding = worksheetMetadataStore.LoadBinding(sheetName);
            var template = templateStore.Load(templateId)
                ?? throw new InvalidOperationException("Template was not found.");
            EnsureTemplateProjectCompatibility(binding, template);
            EnsureTemplateDefinitionCompatibility(binding, template);

            var nextBinding = templateDefinitionNormalizer.ToSheetBinding(template, sheetName);
            var nextFieldMappings = templateDefinitionNormalizer.ToSheetFieldMappings(template, sheetName);
            worksheetMetadataStore.SaveBinding(nextBinding);
            worksheetMetadataStore.SaveFieldMappings(sheetName, template.FieldMappingDefinition, nextFieldMappings);

            worksheetTemplateBindingStore.SaveTemplateBinding(
                new SheetTemplateBinding
                {
                    SheetName = sheetName,
                    TemplateId = template.TemplateId,
                    TemplateName = template.TemplateName,
                    TemplateRevision = template.Revision,
                    TemplateOrigin = StoreTemplateOrigin,
                    AppliedFingerprint = templateFingerprintBuilder.Build(template),
                    TemplateLastAppliedAt = DateTime.UtcNow,
                });
        }

        public void SaveSheetToExistingTemplate(string sheetName, string templateId, int expectedRevision, bool overwriteRevisionConflict)
        {
            var binding = worksheetMetadataStore.LoadBinding(sheetName);
            var templateBinding = worksheetTemplateBindingStore.LoadTemplateBinding(sheetName);
            if (templateBinding == null
                || !string.Equals(templateBinding.TemplateOrigin, StoreTemplateOrigin, StringComparison.Ordinal))
            {
                throw new InvalidOperationException("Current sheet is not bound to a store template.");
            }

            if (!string.Equals(templateBinding.TemplateId, templateId, StringComparison.Ordinal))
            {
                throw new InvalidOperationException("Current sheet is bound to a different template.");
            }

            var storedTemplate = templateStore.Load(templateId)
                ?? throw new InvalidOperationException("Template was not found.");
            EnsureTemplateProjectCompatibility(binding, storedTemplate);
            EnsureTemplateDefinitionCompatibility(binding, storedTemplate);

            var boundRevision = templateBinding.TemplateRevision;
            if (!overwriteRevisionConflict && (!boundRevision.HasValue || storedTemplate.Revision != boundRevision.Value))
            {
                throw new InvalidOperationException("Template revision conflict.");
            }

            var storeExpectedRevision = overwriteRevisionConflict
                ? storedTemplate.Revision
                : boundRevision.Value;

            var timestamp = DateTime.UtcNow;
            var snapshot = BuildSnapshot(
                sheetName,
                binding,
                templateId,
                storedTemplate.TemplateName,
                storedTemplate.Revision + 1,
                storedTemplate.CreatedAtUtc == default(DateTime) ? timestamp : storedTemplate.CreatedAtUtc,
                timestamp);
            var savedTemplate = templateStore.SaveExisting(snapshot, storeExpectedRevision);

            worksheetTemplateBindingStore.SaveTemplateBinding(
                new SheetTemplateBinding
                {
                    SheetName = sheetName,
                    TemplateId = savedTemplate.TemplateId,
                    TemplateName = savedTemplate.TemplateName,
                    TemplateRevision = savedTemplate.Revision,
                    TemplateOrigin = StoreTemplateOrigin,
                    AppliedFingerprint = templateFingerprintBuilder.Build(savedTemplate),
                    TemplateLastAppliedAt = timestamp,
                    DerivedFromTemplateId = templateBinding.DerivedFromTemplateId ?? string.Empty,
                    DerivedFromTemplateRevision = templateBinding.DerivedFromTemplateRevision,
                });
        }

        public void SaveSheetAsNewTemplate(string sheetName, string templateName)
        {
            var binding = worksheetMetadataStore.LoadBinding(sheetName);
            var timestamp = DateTime.UtcNow;
            var snapshot = BuildSnapshot(
                sheetName,
                binding,
                Guid.NewGuid().ToString("N"),
                templateName,
                1,
                timestamp,
                timestamp);
            var savedTemplate = templateStore.SaveNew(snapshot);
            TryLoadTemplateBinding(sheetName, out var previousBinding);

            var derivedFromTemplateId = previousBinding?.DerivedFromTemplateId ?? string.Empty;
            var derivedFromTemplateRevision = previousBinding?.DerivedFromTemplateRevision;
            if (string.IsNullOrWhiteSpace(derivedFromTemplateId) && !string.IsNullOrWhiteSpace(previousBinding?.TemplateId))
            {
                derivedFromTemplateId = previousBinding.TemplateId;
                derivedFromTemplateRevision = previousBinding.TemplateRevision;
            }

            worksheetTemplateBindingStore.SaveTemplateBinding(
                new SheetTemplateBinding
                {
                    SheetName = sheetName,
                    TemplateId = savedTemplate.TemplateId,
                    TemplateName = savedTemplate.TemplateName,
                    TemplateRevision = savedTemplate.Revision,
                    TemplateOrigin = StoreTemplateOrigin,
                    AppliedFingerprint = templateFingerprintBuilder.Build(savedTemplate),
                    TemplateLastAppliedAt = timestamp,
                    DerivedFromTemplateId = derivedFromTemplateId,
                    DerivedFromTemplateRevision = derivedFromTemplateRevision,
                });
        }

        public void DetachTemplate(string sheetName)
        {
            worksheetTemplateBindingStore.SaveTemplateBinding(
                new SheetTemplateBinding
                {
                    SheetName = sheetName,
                    TemplateOrigin = AdHocTemplateOrigin,
                    TemplateLastAppliedAt = DateTime.UtcNow,
                });
        }

        private TemplateDefinition BuildSnapshot(
            string sheetName,
            SheetBinding binding,
            string templateId,
            string templateName,
            int revision,
            DateTime createdAtUtc,
            DateTime updatedAtUtc)
        {
            var fieldMappingDefinition = GetFieldMappingDefinition(binding);
            var fieldMappings = worksheetMetadataStore.LoadFieldMappings(sheetName, fieldMappingDefinition);
            return templateDefinitionNormalizer.Normalize(
                templateId,
                templateName,
                binding,
                fieldMappingDefinition,
                fieldMappings,
                revision,
                createdAtUtc,
                updatedAtUtc);
        }

        private FieldMappingTableDefinition GetFieldMappingDefinition(SheetBinding binding)
        {
            var connector = systemConnectorRegistry.GetRequiredConnector(binding.SystemKey);
            return connector.GetFieldMappingDefinition(binding.ProjectId);
        }

        private bool IsDirty(string sheetName, SheetBinding binding, SheetTemplateBinding templateBinding)
        {
            if (string.IsNullOrWhiteSpace(templateBinding.AppliedFingerprint))
            {
                return false;
            }

            var currentSnapshot = BuildSnapshot(
                sheetName,
                binding,
                templateBinding.TemplateId,
                templateBinding.TemplateName,
                templateBinding.TemplateRevision ?? 0,
                DateTime.MinValue,
                DateTime.MinValue);
            var currentFingerprint = templateFingerprintBuilder.Build(currentSnapshot);
            return !string.Equals(templateBinding.AppliedFingerprint, currentFingerprint, StringComparison.Ordinal);
        }

        private static bool IsTemplateInSameProject(SheetBinding binding, TemplateDefinition template)
        {
            return string.Equals(binding.SystemKey, template.SystemKey, StringComparison.Ordinal)
                && string.Equals(binding.ProjectId, template.ProjectId, StringComparison.Ordinal);
        }

        private static void EnsureTemplateProjectCompatibility(SheetBinding binding, TemplateDefinition template)
        {
            if (!IsTemplateInSameProject(binding, template))
            {
                throw new InvalidOperationException("Template project does not match current sheet binding.");
            }
        }

        private void EnsureTemplateDefinitionCompatibility(SheetBinding binding, TemplateDefinition template)
        {
            if (!IsTemplateCompatibleWithCurrentDefinition(binding, template))
            {
                throw new InvalidOperationException("Template definition is not compatible with current connector definition.");
            }
        }

        private bool IsTemplateCompatibleWithCurrentDefinition(SheetBinding binding, TemplateDefinition template)
        {
            if (!IsTemplateInSameProject(binding, template))
            {
                return false;
            }

            var currentDefinition = GetFieldMappingDefinition(binding);
            var currentFingerprint = TemplateFingerprintBuilder.BuildFieldMappingDefinitionFingerprint(currentDefinition);
            var templateFingerprint = string.IsNullOrWhiteSpace(template.FieldMappingDefinitionFingerprint)
                ? TemplateFingerprintBuilder.BuildFieldMappingDefinitionFingerprint(template.FieldMappingDefinition)
                : template.FieldMappingDefinitionFingerprint;

            return string.Equals(currentFingerprint, templateFingerprint, StringComparison.Ordinal);
        }

        private bool TryLoadBinding(string sheetName, out SheetBinding binding)
        {
            try
            {
                binding = worksheetMetadataStore.LoadBinding(sheetName);
                return binding != null;
            }
            catch (InvalidOperationException)
            {
                binding = null;
                return false;
            }
        }

        private bool TryLoadTemplateBinding(string sheetName, out SheetTemplateBinding templateBinding)
        {
            try
            {
                templateBinding = worksheetTemplateBindingStore.LoadTemplateBinding(sheetName);
                return templateBinding != null;
            }
            catch (InvalidOperationException)
            {
                templateBinding = null;
                return false;
            }
        }
    }
}
