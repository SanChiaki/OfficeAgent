using System;
using System.Collections.Generic;

namespace OfficeAgent.Core.Models
{
    public sealed class SheetInitializationPlan
    {
        public SheetBinding Binding { get; set; } = new SheetBinding();

        public FieldMappingTableDefinition FieldMappingDefinition { get; set; } = new FieldMappingTableDefinition();

        public IReadOnlyList<SheetFieldMappingRow> FieldMappings { get; set; } = Array.Empty<SheetFieldMappingRow>();
    }
}
