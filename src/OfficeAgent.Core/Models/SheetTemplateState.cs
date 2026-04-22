namespace OfficeAgent.Core.Models
{
    public sealed class SheetTemplateState
    {
        public bool HasProjectBinding { get; set; }

        public bool CanApplyTemplate { get; set; }

        public bool CanSaveTemplate { get; set; }

        public bool CanSaveAsTemplate { get; set; }

        public bool IsDirty { get; set; }

        public bool TemplateMissing { get; set; }

        public string ProjectDisplayName { get; set; } = string.Empty;

        public string TemplateId { get; set; } = string.Empty;

        public string TemplateName { get; set; } = string.Empty;

        public int? TemplateRevision { get; set; }

        public int? StoredTemplateRevision { get; set; }

        public string TemplateOrigin { get; set; } = string.Empty;
    }
}
