namespace DataverseSchemaManager.Models
{
    public class AppConfiguration
    {
        public string? ExcelFilePath { get; set; }
        public string? ConnectionString { get; set; }
        public string? TableNameColumn { get; set; } = "Table Name";
        public string? ColumnNameColumn { get; set; } = "Column Name";
        public string? ColumnTypeColumn { get; set; } = "Column Type";
        public string? ChoiceOptionsColumn { get; set; } = "Choice Options";
        public string? OutputCsvPath { get; set; }
    }
}