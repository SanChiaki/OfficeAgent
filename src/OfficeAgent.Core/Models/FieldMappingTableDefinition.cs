using System;

namespace OfficeAgent.Core.Models
{
    public sealed class FieldMappingTableDefinition
    {
        public string SystemKey { get; set; } = string.Empty;

        public FieldMappingColumnDefinition[] Columns { get; set; } = Array.Empty<FieldMappingColumnDefinition>();
    }
}
