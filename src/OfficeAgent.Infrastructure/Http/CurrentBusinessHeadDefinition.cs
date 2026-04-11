namespace OfficeAgent.Infrastructure.Http
{
    public sealed class CurrentBusinessHeadDefinition
    {
        public string FieldKey { get; set; } = string.Empty;
        public string HeaderText { get; set; } = string.Empty;
        public string HeadType { get; set; } = string.Empty;
        public string ActivityId { get; set; } = string.Empty;
        public string ActivityName { get; set; } = string.Empty;
        public bool IsId { get; set; }
    }
}
