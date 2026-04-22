using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using OfficeAgent.Core.Models;

namespace OfficeAgent.Core.Templates
{
    public sealed class TemplateFingerprintBuilder
    {
        public string Build(TemplateDefinition template)
        {
            if (template == null)
            {
                throw new ArgumentNullException(nameof(template));
            }

            var canonicalRows = (template.FieldMappings ?? Array.Empty<TemplateFieldMappingRow>())
                .Select(BuildCanonicalRow)
                .OrderBy(value => value, StringComparer.Ordinal)
                .ToArray();

            var payload = EncodeParts(
                "template",
                template.SystemKey ?? string.Empty,
                template.ProjectId ?? string.Empty,
                template.ProjectName ?? string.Empty,
                template.HeaderStartRow.ToString(),
                template.HeaderRowCount.ToString(),
                template.DataStartRow.ToString(),
                BuildFieldMappingDefinitionFingerprint(template.FieldMappingDefinition),
                EncodeParts(canonicalRows));

            return ComputeSha256Hex(payload);
        }

        public static string BuildFieldMappingDefinitionFingerprint(FieldMappingTableDefinition definition)
        {
            if (definition == null)
            {
                return ComputeSha256Hex(string.Empty);
            }

            var columns = definition.Columns ?? Array.Empty<FieldMappingColumnDefinition>();
            var canonicalColumns = columns
                .Select((column, index) =>
                {
                    var value = column ?? new FieldMappingColumnDefinition();
                    return EncodeParts(
                        index.ToString(),
                        value.ColumnName ?? string.Empty,
                        value.Role.ToString(),
                        value.RoleKey ?? string.Empty);
                })
                .ToArray();

            var payload = EncodeParts("definition", definition.SystemKey ?? string.Empty, EncodeParts(canonicalColumns));

            return ComputeSha256Hex(payload);
        }

        private static string BuildCanonicalRow(TemplateFieldMappingRow row)
        {
            var values = row?.Values ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var pairs = values
                .Where(pair => !string.Equals(pair.Key, "SheetName", StringComparison.OrdinalIgnoreCase))
                .OrderBy(pair => pair.Key, StringComparer.Ordinal)
                .Select(pair => EncodeParts(pair.Key ?? string.Empty, pair.Value ?? string.Empty));

            return EncodeParts(pairs);
        }

        private static string EncodeParts(IEnumerable<string> values)
        {
            if (values == null)
            {
                return "0#";
            }

            var builder = new StringBuilder();
            foreach (var value in values)
            {
                var token = value ?? string.Empty;
                builder.Append(token.Length.ToString());
                builder.Append('#');
                builder.Append(token);
            }

            return builder.ToString();
        }

        private static string EncodeParts(params string[] values)
        {
            return EncodeParts((IEnumerable<string>)values ?? Array.Empty<string>());
        }

        private static string ComputeSha256Hex(string value)
        {
            using (var sha = SHA256.Create())
            {
                var bytes = Encoding.UTF8.GetBytes(value ?? string.Empty);
                var hash = sha.ComputeHash(bytes);
                var builder = new StringBuilder(hash.Length * 2);
                foreach (var item in hash)
                {
                    builder.Append(item.ToString("x2"));
                }

                return builder.ToString();
            }
        }
    }
}
