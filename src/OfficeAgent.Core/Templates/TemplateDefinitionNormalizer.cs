using System;
using System.Collections.Generic;
using System.Linq;
using OfficeAgent.Core.Models;

namespace OfficeAgent.Core.Templates
{
    public sealed class TemplateDefinitionNormalizer
    {
        public TemplateDefinition Normalize(
            string templateId,
            string templateName,
            SheetBinding binding,
            FieldMappingTableDefinition fieldMappingDefinition,
            IReadOnlyList<SheetFieldMappingRow> fieldMappings,
            int revision,
            DateTime createdAtUtc,
            DateTime updatedAtUtc)
        {
            if (binding == null)
            {
                throw new ArgumentNullException(nameof(binding));
            }

            if (fieldMappingDefinition == null)
            {
                throw new ArgumentNullException(nameof(fieldMappingDefinition));
            }

            var clonedDefinition = CloneFieldMappingDefinition(fieldMappingDefinition);

            return new TemplateDefinition
            {
                TemplateId = templateId ?? string.Empty,
                TemplateName = templateName ?? string.Empty,
                SystemKey = binding.SystemKey ?? string.Empty,
                ProjectId = binding.ProjectId ?? string.Empty,
                ProjectName = binding.ProjectName ?? string.Empty,
                HeaderStartRow = binding.HeaderStartRow,
                HeaderRowCount = binding.HeaderRowCount,
                DataStartRow = binding.DataStartRow,
                FieldMappingDefinition = clonedDefinition,
                FieldMappingDefinitionFingerprint = TemplateFingerprintBuilder.BuildFieldMappingDefinitionFingerprint(clonedDefinition),
                FieldMappings = NormalizeFieldMappings(fieldMappings),
                Revision = revision,
                CreatedAtUtc = createdAtUtc,
                UpdatedAtUtc = updatedAtUtc,
            };
        }

        public SheetBinding ToSheetBinding(TemplateDefinition template, string sheetName)
        {
            if (template == null)
            {
                throw new ArgumentNullException(nameof(template));
            }

            return new SheetBinding
            {
                SheetName = sheetName ?? string.Empty,
                SystemKey = template.SystemKey ?? string.Empty,
                ProjectId = template.ProjectId ?? string.Empty,
                ProjectName = template.ProjectName ?? string.Empty,
                HeaderStartRow = template.HeaderStartRow,
                HeaderRowCount = template.HeaderRowCount,
                DataStartRow = template.DataStartRow,
            };
        }

        public SheetFieldMappingRow[] ToSheetFieldMappings(TemplateDefinition template, string sheetName)
        {
            if (template == null)
            {
                throw new ArgumentNullException(nameof(template));
            }

            return (template.FieldMappings ?? Array.Empty<TemplateFieldMappingRow>())
                .Select(row => new SheetFieldMappingRow
                {
                    SheetName = sheetName ?? string.Empty,
                    Values = new Dictionary<string, string>(
                        row?.Values ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase),
                        StringComparer.OrdinalIgnoreCase),
                })
                .ToArray();
        }

        private static TemplateFieldMappingRow[] NormalizeFieldMappings(IReadOnlyList<SheetFieldMappingRow> fieldMappings)
        {
            if (fieldMappings == null || fieldMappings.Count == 0)
            {
                return Array.Empty<TemplateFieldMappingRow>();
            }

            return fieldMappings
                .Select(row => new TemplateFieldMappingRow
                {
                    Values = (row?.Values ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase))
                        .Where(pair => !string.Equals(pair.Key, "SheetName", StringComparison.OrdinalIgnoreCase))
                        .ToDictionary(
                            pair => pair.Key,
                            pair => pair.Value,
                            StringComparer.OrdinalIgnoreCase),
                })
                .ToArray();
        }

        private static FieldMappingTableDefinition CloneFieldMappingDefinition(FieldMappingTableDefinition definition)
        {
            return new FieldMappingTableDefinition
            {
                SystemKey = definition.SystemKey ?? string.Empty,
                Columns = (definition.Columns ?? Array.Empty<FieldMappingColumnDefinition>())
                    .Select(column => new FieldMappingColumnDefinition
                    {
                        ColumnName = column?.ColumnName ?? string.Empty,
                        Role = column?.Role ?? default(FieldMappingSemanticRole),
                        RoleKey = column?.RoleKey ?? string.Empty,
                    })
                    .ToArray(),
            };
        }
    }
}
