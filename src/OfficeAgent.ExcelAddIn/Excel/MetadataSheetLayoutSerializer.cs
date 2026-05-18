using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeAgent.ExcelAddIn.Excel
{
    internal sealed class MetadataSheetLayoutSerializer
    {
        private static readonly string[] SectionOrder =
        {
            "SheetBindings",
            "SheetFieldMappings",
            "TemplateBindings",
        };

        internal static IReadOnlyList<string> OrderedSectionNames => SectionOrder;

        public string[][] Render(IReadOnlyDictionary<string, MetadataSectionDocument> sections)
        {
            var rendered = new List<string[]>();

            foreach (var sectionName in SectionOrder)
            {
                if (sections == null ||
                    !sections.TryGetValue(sectionName, out var section) ||
                    section == null)
                {
                    continue;
                }

                if (rendered.Count > 0)
                {
                    rendered.Add(Array.Empty<string>());
                    rendered.Add(Array.Empty<string>());
                }

                rendered.Add(new[] { section.Title });
                rendered.Add(section.Headers ?? Array.Empty<string>());
                rendered.AddRange(section.Rows ?? Array.Empty<string[]>());
            }

            return rendered.ToArray();
        }

        public string[][] ReadTable(string tableName, string[][] sheetRows)
        {
            var section = ReadSection(tableName, sheetRows);
            return section?.Rows ?? Array.Empty<string[]>();
        }

        public MetadataSectionDocument ReadSection(string tableName, string[][] sheetRows)
        {
            if (string.IsNullOrWhiteSpace(tableName) || sheetRows == null)
            {
                return null;
            }

            for (var rowIndex = 0; rowIndex < sheetRows.Length; rowIndex++)
            {
                var row = sheetRows[rowIndex] ?? Array.Empty<string>();
                if (row.Length == 0 ||
                    !string.Equals(row[0], tableName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var headers = rowIndex + 1 < sheetRows.Length
                    ? (sheetRows[rowIndex + 1] ?? Array.Empty<string>())
                    : Array.Empty<string>();
                var result = new List<string[]>();
                for (var dataRowIndex = rowIndex + 2; dataRowIndex < sheetRows.Length; dataRowIndex++)
                {
                    var candidate = sheetRows[dataRowIndex] ?? Array.Empty<string>();
                    if (candidate.Length > 0 &&
                        !string.IsNullOrWhiteSpace(candidate[0]) &&
                        Array.IndexOf(SectionOrder, candidate[0]) >= 0)
                    {
                        break;
                    }

                    if (candidate.All(string.IsNullOrEmpty))
                    {
                        break;
                    }

                    result.Add(candidate);
                }

                return new MetadataSectionDocument(tableName, headers, result.ToArray());
            }

            return null;
        }
    }
}
