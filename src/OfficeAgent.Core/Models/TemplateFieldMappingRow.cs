using System;
using System.Collections.Generic;

namespace OfficeAgent.Core.Models
{
    public sealed class TemplateFieldMappingRow
    {
        public Dictionary<string, string> Values { get; set; } =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    }
}
