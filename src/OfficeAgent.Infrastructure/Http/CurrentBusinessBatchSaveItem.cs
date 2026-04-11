namespace OfficeAgent.Infrastructure.Http
{
    public sealed class CurrentBusinessBatchSaveItem
    {
        public string ProjectId { get; set; } = string.Empty;
        public string Id { get; set; } = string.Empty;
        public string FieldKey { get; set; } = string.Empty;
        public string Value { get; set; } = string.Empty;
    }
}
