using System;
using System.Linq;
using OfficeAgent.Core.Models;

namespace OfficeAgent.Core.Sync
{
    public sealed class FieldMappingValueAccessor
    {
        public string GetValue(
            FieldMappingTableDefinition definition,
            SheetFieldMappingRow row,
            FieldMappingSemanticRole role)
        {
            if (definition == null)
            {
                throw new ArgumentNullException(nameof(definition));
            }

            if (row == null)
            {
                throw new ArgumentNullException(nameof(row));
            }

            var columnName = (definition.Columns ?? Array.Empty<FieldMappingColumnDefinition>())
                .Where(column => column != null && column.Role == role)
                .Select(column => column.ColumnName)
                .FirstOrDefault(name => !string.IsNullOrWhiteSpace(name));

            if (string.IsNullOrWhiteSpace(columnName) || row.Values == null)
            {
                return string.Empty;
            }

            return row.Values.TryGetValue(columnName, out var value)
                ? value ?? string.Empty
                : string.Empty;
        }

        public bool GetBoolean(
            FieldMappingTableDefinition definition,
            SheetFieldMappingRow row,
            FieldMappingSemanticRole role)
        {
            return string.Equals(GetValue(definition, row, role), "true", StringComparison.OrdinalIgnoreCase);
        }
    }
}
