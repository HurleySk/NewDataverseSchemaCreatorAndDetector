namespace DataverseSchemaManager.Models
{
    public class SchemaDefinition
    {
        // Required Excel columns
        public string TableName { get; set; } = string.Empty;
        public string ColumnName { get; set; } = string.Empty;
        public string ColumnType { get; set; } = string.Empty;

        // Optional Excel columns - Choice fields
        public string? ChoiceOptions { get; set; }

        // Optional Excel columns - Phase 1 enhancements
        public string? LogicalName { get; set; }
        public string? TableLogicalName { get; set; }
        public string? TableDisplayCollectionName { get; set; }
        public string? Description { get; set; }
        public string? Required { get; set; }

        // Runtime properties
        public bool TableExistsInDataverse { get; set; }
        public bool ColumnExistsInDataverse { get; set; }
        public string? ErrorMessage { get; set; }
    }
}