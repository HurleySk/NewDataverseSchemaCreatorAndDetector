# Dataverse Schema Creator and Detector

A .NET console application for detecting existing schema and creating new schema in Microsoft Dataverse environments based on Excel file definitions.

## Features

- **Schema Detection**: Automatically detects whether specified columns exist in Dataverse tables
- **Bulk Schema Creation**: Create multiple columns across different tables in a single operation
- **Excel Import**: Read schema definitions from Excel files with configurable column mappings
- **CSV Export**: Export detected new schema to CSV files for review
- **Solution Management**: Create schema within specific Dataverse solutions with publisher prefix validation
- **Multiple Data Types**: Support for various column types including:
  - Text/String
  - Number/Integer
  - Decimal/Money/Currency
  - Date/DateTime
  - Boolean/Bit

## Prerequisites

- .NET 9.0 SDK or later
- Access to a Microsoft Dataverse environment
- Valid Dataverse credentials (OAuth, Client ID/Secret, or interactive authentication)
- Excel file with schema definitions

## Installation

1. Clone the repository:
```bash
git clone https://github.com/yourusername/NewDataverseSchemaCreatorAndDetector.git
cd NewDataverseSchemaCreatorAndDetector
```

2. Copy the configuration template:
```bash
cp appsettings.template.json appsettings.json
```

3. Build the project:
```bash
dotnet build
```

## Configuration

Edit `appsettings.json` to configure the application:

```json
{
  "ExcelFilePath": "path/to/your/schema.xlsx",
  "ConnectionString": "your-dataverse-connection-string",
  "TableNameColumn": "Table Name",
  "ColumnNameColumn": "Column Name",
  "ColumnTypeColumn": "Column Type",
  "OutputCsvPath": "new_schemas.csv"
}
```

### Connection String Format

The connection string should follow this format:
```
AuthType=OAuth;Url=https://yourorg.crm.dynamics.com;Username=user@org.com;Password=yourpassword;RequireNewInstance=true
```

For other authentication types, refer to the [Microsoft Dataverse documentation](https://docs.microsoft.com/en-us/power-apps/developer/data-platform/xrm-tooling/use-connection-strings-xrm-tooling-connect).

## Excel File Format

Your Excel file should have the following columns (column names are configurable):

| Table Name | Column Name | Column Type |
|------------|-------------|-------------|
| account    | new_field1  | text        |
| contact    | new_field2  | number      |
| lead       | new_field3  | boolean     |

### Supported Column Types

- `text` or `string` - Single line of text (max 100 characters)
- `number`, `int`, or `integer` - Whole number
- `decimal`, `money`, or `currency` - Currency/decimal values
- `date` or `datetime` - Date and time values
- `boolean`, `bool`, or `bit` - Yes/No (two options)

## Usage

### Running the Application

```bash
dotnet run
```

The application will:
1. Connect to your Dataverse environment
2. Load schema definitions from the Excel file
3. Check which schemas already exist
4. Present a menu with options to:
   - Export new schemas to CSV
   - Create new schemas in Dataverse
   - Both export and create
   - Re-scan schemas
   - Exit

### Generate Sample Excel File

To generate a sample Excel file for testing:

```bash
dotnet run -- --generate-sample
```

This creates `sample_schema.xlsx` with example schema definitions.

### Menu Options

1. **Export new schemas to CSV**: Saves all non-existing schemas to a CSV file for review
2. **Create new schemas in Dataverse**:
   - Lists available unmanaged solutions
   - Validates publisher prefix
   - Creates the columns in the selected solution
   - Publishes customizations
3. **Export and Create**: Performs both operations sequentially
4. **Re-scan schemas**: Re-checks the existence of all schemas
5. **Exit**: Closes the application

## Project Structure

```
├── Models/
│   ├── AppConfiguration.cs     # Configuration model
│   └── SchemaDefinition.cs     # Schema data model
├── Services/
│   ├── DataverseService.cs     # Dataverse connection and operations
│   ├── ExcelReaderService.cs   # Excel file reading
│   └── CsvExportService.cs     # CSV export functionality
├── Utils/
│   └── GenerateSampleExcel.cs  # Sample file generator
├── Program.cs                   # Main application entry point
├── appsettings.json            # Application configuration (user-specific)
├── appsettings.template.json   # Configuration template
└── DataverseSchemaManager.csproj # Project file
```

## Error Handling

The application handles common errors including:
- Missing or invalid Excel files
- Connection failures to Dataverse
- Missing required columns in Excel
- Tables that don't exist in Dataverse
- Invalid column types
- Publisher prefix mismatches

## Security Considerations

- Never commit `appsettings.json` with actual credentials
- Use environment variables or Azure Key Vault for production deployments
- Consider using managed identities when running in Azure
- The `.gitignore` file excludes sensitive configuration files

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## Troubleshooting

### Connection Issues
- Verify your connection string format
- Ensure your account has appropriate permissions in Dataverse
- Check network connectivity to the Dataverse environment

### Schema Creation Failures
- Verify the table names exist in Dataverse
- Check that column names follow Dataverse naming conventions
- Ensure you have customization permissions
- Verify the solution is unmanaged

### Excel Reading Issues
- Ensure column headers match the configuration
- Check that the Excel file is not open in another application
- Verify the file path is correct

## License

This project is provided as-is for educational and development purposes.

## Support

For issues or questions, please create an issue in the GitHub repository.