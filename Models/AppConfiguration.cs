namespace DataverseSchemaManager.Models
{
    public class AppConfiguration
    {
        public string? ExcelFilePath { get; set; }
        public string? ConnectionString { get; set; }

        // Required column mappings
        public string? TableNameColumn { get; set; } = "Table Name";
        public string? ColumnNameColumn { get; set; } = "Column Name";
        public string? ColumnTypeColumn { get; set; } = "Column Type";

        // Optional column mappings - Existing
        public string? ChoiceOptionsColumn { get; set; } = "Choice Options";

        // Optional column mappings - Phase 1 enhancements
        public string? LogicalNameColumn { get; set; }
        public string? TableLogicalNameColumn { get; set; }
        public string? TableDisplayCollectionNameColumn { get; set; }
        public string? DescriptionColumn { get; set; }
        public string? RequiredColumn { get; set; }

        public string? OutputCsvPath { get; set; }
    }
}