using System;
using System.Collections.Generic;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Templates;
using Xunit;

namespace OfficeAgent.Core.Tests
{
    public sealed class TemplateFingerprintBuilderTests
    {
        [Fact]
        public void BuildReturnsDifferentHashesForAmbiguousDelimiterValues()
        {
            var builder = new TemplateFingerprintBuilder();
            var templateA = CreateTemplate(new[]
            {
                CreateRow("a", "b=c"),
            });
            var templateB = CreateTemplate(new[]
            {
                CreateRow("a=b", "c"),
            });

            var fingerprintA = builder.Build(templateA);
            var fingerprintB = builder.Build(templateB);

            Assert.NotEqual(fingerprintA, fingerprintB);
        }

        [Fact]
        public void BuildReturnsDifferentHashesForEmbeddedNewlinesAndMultipleRows()
        {
            var builder = new TemplateFingerprintBuilder();
            var templateA = CreateTemplate(new[]
            {
                CreateRow("k", "v1\nk2=v2"),
            });
            var templateB = CreateTemplate(new[]
            {
                CreateRow("k", "v1"),
                CreateRow("k2", "v2"),
            });

            var fingerprintA = builder.Build(templateA);
            var fingerprintB = builder.Build(templateB);

            Assert.NotEqual(fingerprintA, fingerprintB);
        }

        private static TemplateDefinition CreateTemplate(TemplateFieldMappingRow[] rows)
        {
            return new TemplateDefinition
            {
                TemplateId = "template-a",
                TemplateName = "模板A",
                SystemKey = "system-a",
                ProjectId = "project-a",
                ProjectName = "项目A",
                HeaderStartRow = 1,
                HeaderRowCount = 2,
                DataStartRow = 3,
                FieldMappingDefinition = new FieldMappingTableDefinition
                {
                    SystemKey = "system-a",
                    Columns = new[]
                    {
                        new FieldMappingColumnDefinition
                        {
                            ColumnName = "ApiFieldKey",
                            Role = FieldMappingSemanticRole.ApiFieldKey,
                        },
                    },
                },
                FieldMappings = rows ?? Array.Empty<TemplateFieldMappingRow>(),
            };
        }

        private static TemplateFieldMappingRow CreateRow(string key, string value)
        {
            return new TemplateFieldMappingRow
            {
                Values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    [key] = value,
                },
            };
        }
    }
}
