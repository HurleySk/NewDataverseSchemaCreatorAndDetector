namespace DataverseSchemaManager.Models
{
    public class SchemaDefinition
    {
        // Required for DETECTION (Phase 1 - Assess)
        public string LogicalName { get; set; } = string.Empty;
        public string TableLogicalName { get; set; } = string.Empty;

        // Required for CREATION (Phase 2 - Create)
        public string? TableName { get; set; }
        public string? ColumnName { get; set; }
        public string? ColumnType { get; set; }

        // Optional Excel columns - Choice fields
        public string? ChoiceOptions { get; set; }

        // Optional Excel columns - Lookup/Customer fields
        public string? LookupTargetTable { get; set; }
        public string? LookupRelationshipName { get; set; }
        public string? CustomerTargetTables { get; set; }

        // Optional Excel columns - Metadata
        public string? TableDisplayCollectionName { get; set; }
        public string? Description { get; set; }
        public string? Required { get; set; }

        // Runtime properties
        public bool TableExistsInDataverse { get; set; }
        public bool ColumnExistsInDataverse { get; set; }
        public string? ErrorMessage { get; set; }
    }
}