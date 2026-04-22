using System;

namespace OfficeAgent.Core.Models
{
    public sealed class TemplateDefinition
    {
        public string TemplateId { get; set; } = string.Empty;

        public string TemplateName { get; set; } = string.Empty;

        public string SystemKey { get; set; } = string.Empty;

        public string ProjectId { get; set; } = string.Empty;

        public string ProjectName { get; set; } = string.Empty;

        public int HeaderStartRow { get; set; } = 1;

        public int HeaderRowCount { get; set; } = 2;

        public int DataStartRow { get; set; } = 3;

        public FieldMappingTableDefinition FieldMappingDefinition { get; set; } = new FieldMappingTableDefinition();

        public string FieldMappingDefinitionFingerprint { get; set; } = string.Empty;

        public TemplateFieldMappingRow[] FieldMappings { get; set; } = Array.Empty<TemplateFieldMappingRow>();

        public int Revision { get; set; }

        public DateTime CreatedAtUtc { get; set; }

        public DateTime UpdatedAtUtc { get; set; }
    }
}
