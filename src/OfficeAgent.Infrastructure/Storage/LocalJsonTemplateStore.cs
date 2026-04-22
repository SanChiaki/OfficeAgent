using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;

namespace OfficeAgent.Infrastructure.Storage
{
    public sealed class LocalJsonTemplateStore : ITemplateStore
    {
        private readonly string rootPath;

        public LocalJsonTemplateStore(string rootPath)
        {
            this.rootPath = rootPath ?? throw new ArgumentNullException(nameof(rootPath));
        }

        public IReadOnlyList<TemplateDefinition> ListByProject(string systemKey, string projectId)
        {
            var directory = GetProjectDirectory(systemKey, projectId);
            if (!Directory.Exists(directory))
            {
                return Array.Empty<TemplateDefinition>();
            }

            return Directory.GetFiles(directory, "*.json", SearchOption.TopDirectoryOnly)
                .Select(ReadTemplateFromPath)
                .Where(template => template != null)
                .OrderByDescending(template => template.UpdatedAtUtc)
                .ToArray();
        }

        public TemplateDefinition Load(string templateId)
        {
            if (string.IsNullOrWhiteSpace(templateId) || !Directory.Exists(rootPath))
            {
                return null;
            }

            foreach (var path in Directory.GetFiles(rootPath, "*.json", SearchOption.AllDirectories))
            {
                var template = ReadTemplateFromPath(path);
                if (template != null &&
                    string.Equals(template.TemplateId, templateId, StringComparison.Ordinal))
                {
                    return template;
                }
            }

            return null;
        }

        public TemplateDefinition SaveNew(TemplateDefinition template)
        {
            var normalized = NormalizeForSave(template);
            normalized.Revision = Math.Max(1, normalized.Revision);

            if (normalized.CreatedAtUtc == default(DateTime))
            {
                normalized.CreatedAtUtc = DateTime.UtcNow;
            }

            if (normalized.UpdatedAtUtc == default(DateTime))
            {
                normalized.UpdatedAtUtc = normalized.CreatedAtUtc;
            }

            WriteTemplate(normalized);
            return Clone(normalized);
        }

        public TemplateDefinition SaveExisting(TemplateDefinition template, int expectedRevision)
        {
            if (template == null)
            {
                throw new ArgumentNullException(nameof(template));
            }

            var existing = Load(template.TemplateId);
            if (existing == null)
            {
                throw new InvalidOperationException("未找到模板。");
            }

            if (existing.Revision != expectedRevision)
            {
                throw new InvalidOperationException("模板版本已变化。");
            }

            var normalized = NormalizeForSave(template);
            normalized.CreatedAtUtc = existing.CreatedAtUtc;

            if (normalized.UpdatedAtUtc == default(DateTime))
            {
                normalized.UpdatedAtUtc = DateTime.UtcNow;
            }

            if (normalized.Revision <= 0)
            {
                normalized.Revision = existing.Revision;
            }

            WriteTemplate(normalized);
            return Clone(normalized);
        }

        private void WriteTemplate(TemplateDefinition template)
        {
            var directory = GetProjectDirectory(template.SystemKey, template.ProjectId);
            Directory.CreateDirectory(directory);

            var path = GetTemplatePath(directory, template.TemplateId);
            File.WriteAllText(path, JsonConvert.SerializeObject(template, Formatting.Indented));
        }

        private TemplateDefinition ReadTemplateFromPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                return null;
            }

            var json = File.ReadAllText(path);
            var persisted = JsonConvert.DeserializeObject<TemplateDefinition>(json);
            return persisted == null ? null : Clone(persisted);
        }

        private string GetProjectDirectory(string systemKey, string projectId)
        {
            return Path.Combine(rootPath, systemKey ?? string.Empty, projectId ?? string.Empty);
        }

        private static string GetTemplatePath(string directory, string templateId)
        {
            return Path.Combine(directory, (templateId ?? string.Empty) + ".json");
        }

        private static TemplateDefinition NormalizeForSave(TemplateDefinition template)
        {
            if (template == null)
            {
                throw new ArgumentNullException(nameof(template));
            }

            return Clone(template);
        }

        private static TemplateDefinition Clone(TemplateDefinition template)
        {
            if (template == null)
            {
                return null;
            }

            return new TemplateDefinition
            {
                TemplateId = template.TemplateId ?? string.Empty,
                TemplateName = template.TemplateName ?? string.Empty,
                SystemKey = template.SystemKey ?? string.Empty,
                ProjectId = template.ProjectId ?? string.Empty,
                ProjectName = template.ProjectName ?? string.Empty,
                HeaderStartRow = template.HeaderStartRow,
                HeaderRowCount = template.HeaderRowCount,
                DataStartRow = template.DataStartRow,
                FieldMappingDefinitionFingerprint = template.FieldMappingDefinitionFingerprint ?? string.Empty,
                FieldMappingDefinition = Clone(template.FieldMappingDefinition),
                FieldMappings = (template.FieldMappings ?? Array.Empty<TemplateFieldMappingRow>())
                    .Select(Clone)
                    .ToArray(),
                Revision = template.Revision,
                CreatedAtUtc = template.CreatedAtUtc,
                UpdatedAtUtc = template.UpdatedAtUtc,
            };
        }

        private static FieldMappingTableDefinition Clone(FieldMappingTableDefinition definition)
        {
            if (definition == null)
            {
                return new FieldMappingTableDefinition();
            }

            return new FieldMappingTableDefinition
            {
                SystemKey = definition.SystemKey ?? string.Empty,
                Columns = (definition.Columns ?? Array.Empty<FieldMappingColumnDefinition>())
                    .Select(Clone)
                    .ToArray(),
            };
        }

        private static FieldMappingColumnDefinition Clone(FieldMappingColumnDefinition column)
        {
            if (column == null)
            {
                return new FieldMappingColumnDefinition();
            }

            return new FieldMappingColumnDefinition
            {
                ColumnName = column.ColumnName ?? string.Empty,
                Role = column.Role,
                RoleKey = column.RoleKey ?? string.Empty,
            };
        }

        private static TemplateFieldMappingRow Clone(TemplateFieldMappingRow row)
        {
            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var entry in row?.Values ?? new Dictionary<string, string>())
            {
                values[entry.Key ?? string.Empty] = entry.Value ?? string.Empty;
            }

            return new TemplateFieldMappingRow
            {
                Values = values,
            };
        }
    }
}
