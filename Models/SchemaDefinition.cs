namespace DataverseSchemaManager.Models
{
    public class SchemaDefinition
    {
        public string TableName { get; set; } = string.Empty;
        public string ColumnName { get; set; } = string.Empty;
        public string ColumnType { get; set; } = string.Empty;
        public bool ExistsInDataverse { get; set; }
        public string? ErrorMessage { get; set; }
    }
}