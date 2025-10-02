namespace DataverseSchemaManager.Models
{
    public class AppConfiguration
    {
        // Input/Output file paths
        public string? ExcelFilePath { get; set; }
        public string? ConnectionString { get; set; }
        public string? OutputCsvPath { get; set; } = "create_template.csv";

        // Required column mappings for ASSESS (detection)
        public string? LogicalNameColumn { get; set; } = "Column Logical Name";
        public string? TableLogicalNameColumn { get; set; } = "Table Logical Name";

        // Optional column mappings for ASSESS (passthrough to template)
        public string? ColumnTypeColumn { get; set; } = "Column Type";

        // Optional column mappings for CREATE (if reading from full CSV)
        public string? TableNameColumn { get; set; } = "Table Name";
        public string? ColumnNameColumn { get; set; } = "Column Name";

        // Optional column mappings - Choice fields
        public string? ChoiceOptionsColumn { get; set; } = "Choice Options";

        // Optional column mappings - Lookup/Customer fields
        public string? LookupTargetTableColumn { get; set; }
        public string? LookupRelationshipNameColumn { get; set; }
        public string? CustomerTargetTablesColumn { get; set; }

        // Optional column mappings - Metadata
        public string? TableDisplayCollectionNameColumn { get; set; }
        public string? DescriptionColumn { get; set; }
        public string? RequiredColumn { get; set; }

        // Optional column mappings - Filtering
        public string? IncludeColumn { get; set; }
    }
}