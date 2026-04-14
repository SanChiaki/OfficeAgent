using System;
using System.Collections.Generic;

namespace OfficeAgent.Core.Models
{
    public sealed class SheetFieldMappingRow
    {
        public string SheetName { get; set; } = string.Empty;

        public IReadOnlyDictionary<string, string> Values { get; set; } =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    }
}
