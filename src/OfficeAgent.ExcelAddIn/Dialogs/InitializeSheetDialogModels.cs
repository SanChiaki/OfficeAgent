using System;
using System.Collections.Generic;
using System.Linq;
using OfficeAgent.Core.Models;

namespace OfficeAgent.ExcelAddIn.Dialogs
{
    public enum InitializeSheetMode
    {
        ConfigOnly,
        TemplateImport,
    }

    public sealed class InitializeSheetDialogRequest
    {
        public string ProjectDisplayName { get; set; } = string.Empty;

        public bool IsBlankSheet { get; set; }

        public bool SupportsTemplateImport { get; set; }
    }

    public sealed class InitializeSheetTemplateLoadResult
    {
        private InitializeSheetTemplateLoadResult(
            bool isSupported,
            bool isSuccess,
            string disabledReason,
            IReadOnlyList<BusinessExportTemplateOption> templates)
        {
            IsSupported = isSupported;
            IsSuccess = isSuccess;
            DisabledReason = disabledReason ?? string.Empty;
            Templates = templates ?? Array.Empty<BusinessExportTemplateOption>();
        }

        public bool IsSupported { get; }

        public bool IsSuccess { get; }

        public bool IsFailed => IsSupported && !IsSuccess;

        public string DisabledReason { get; }

        public IReadOnlyList<BusinessExportTemplateOption> Templates { get; }

        public static InitializeSheetTemplateLoadResult Supported(
            IEnumerable<BusinessExportTemplateOption> templates)
        {
            return Success(templates);
        }

        public static InitializeSheetTemplateLoadResult Success(
            IEnumerable<BusinessExportTemplateOption> templates)
        {
            return new InitializeSheetTemplateLoadResult(
                isSupported: true,
                isSuccess: true,
                disabledReason: string.Empty,
                templates: NormalizeTemplates(templates));
        }

        public static InitializeSheetTemplateLoadResult Failed(string disabledReason)
        {
            return new InitializeSheetTemplateLoadResult(
                isSupported: true,
                isSuccess: false,
                disabledReason: disabledReason,
                templates: Array.Empty<BusinessExportTemplateOption>());
        }

        public static InitializeSheetTemplateLoadResult Unsupported(string disabledReason)
        {
            return new InitializeSheetTemplateLoadResult(
                isSupported: false,
                isSuccess: false,
                disabledReason: disabledReason,
                templates: Array.Empty<BusinessExportTemplateOption>());
        }

        private static IReadOnlyList<BusinessExportTemplateOption> NormalizeTemplates(
            IEnumerable<BusinessExportTemplateOption> templates)
        {
            return (templates ?? Array.Empty<BusinessExportTemplateOption>())
                .Where(template => template != null && !string.IsNullOrWhiteSpace(template.TemplateId))
                .ToArray();
        }
    }

    public sealed class InitializeSheetDialogResult
    {
        public InitializeSheetMode Mode { get; set; }

        public BusinessExportTemplateOption SelectedTemplate { get; set; }

        public string SelectedTemplateId => SelectedTemplate?.TemplateId ?? string.Empty;
    }
}
