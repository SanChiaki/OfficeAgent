namespace OfficeAgent.Core.Models
{
    public sealed class FieldMappingColumnDefinition
    {
        public string ColumnName { get; set; } = string.Empty;

        public FieldMappingSemanticRole Role { get; set; }

        public string RoleKey { get; set; } = string.Empty;
    }
}
