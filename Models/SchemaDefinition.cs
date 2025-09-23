namespace DataverseSchemaManager.Models
{
    public class SchemaDefinition
    {
        public string TableName { get; set; } = string.Empty;
        public string ColumnName { get; set; } = string.Empty;
        public string ColumnType { get; set; } = string.Empty;
        public string? ChoiceOptions { get; set; }
        public bool TableExistsInDataverse { get; set; }
        public bool ColumnExistsInDataverse { get; set; }
        public string? ErrorMessage { get; set; }
    }
}