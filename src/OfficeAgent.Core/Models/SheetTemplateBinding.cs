using System;

namespace OfficeAgent.Core.Models
{
    public sealed class SheetTemplateBinding
    {
        public string SheetName { get; set; } = string.Empty;

        public string TemplateId { get; set; } = string.Empty;

        public string TemplateName { get; set; } = string.Empty;

        public int? TemplateRevision { get; set; }

        public string TemplateOrigin { get; set; } = string.Empty;

        public string AppliedFingerprint { get; set; } = string.Empty;

        public DateTime? TemplateLastAppliedAt { get; set; }

        public string DerivedFromTemplateId { get; set; } = string.Empty;

        public int? DerivedFromTemplateRevision { get; set; }
    }
}
