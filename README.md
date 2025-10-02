# Dataverse Schema Creator and Detector

A .NET console tool for detecting and creating Dataverse schema (tables and columns) from Excel definitions.

## Features

- **Detect** existing schema in Dataverse tables
- **Create** multiple columns across tables in bulk
- **Auto-create** missing tables when needed (with user confirmation)
- **Configure** logical names, descriptions, required levels, and pluralization
- **Export** new schema to CSV for review
- **Solution-aware** creation with publisher prefix validation

## Quick Start

**Prerequisites**: .NET 9.0 SDK, Dataverse access, Excel file with schema definitions

```bash
# Clone and build
git clone https://github.com/yourusername/NewDataverseSchemaCreatorAndDetector.git
cd NewDataverseSchemaCreatorAndDetector
dotnet build

# Configure
cp appsettings.template.json appsettings.json
# Edit appsettings.json with your settings

# Run
dotnet run
```

**Generate sample Excel:**
```bash
dotnet run -- --generate-sample
```

## Configuration

Edit `appsettings.json`:

```json
{
  "ExcelFilePath": "path/to/schema.xlsx",
  "ConnectionString": "AuthType=OAuth;Username=user@org.com;Url=https://yourorg.crm.dynamics.com/;AppId=51f81489-12ee-4a9e-aaae-a2591f45987d;RedirectUri=http://localhost;TokenCacheStorePath=C:\\path\\to\\tokencache",
  "TableNameColumn": "Table Name",
  "ColumnNameColumn": "Column Name",
  "ColumnTypeColumn": "Column Type",
  "OutputCsvPath": "new_schemas.csv"
}
```

### Optional Enhancement Columns

Add these to `appsettings.json` if your Excel has these columns:

```json
{
  "ChoiceOptionsColumn": "Choice Options",
  "LogicalNameColumn": "Logical Name",
  "TableLogicalNameColumn": "Table Logical Name",
  "TableDisplayCollectionNameColumn": "Display Plural",
  "DescriptionColumn": "Description",
  "RequiredColumn": "Required"
}
```

**Connection String Notes:**
- Include `Username` for OAuth token caching (avoids re-authentication)
- See [Dataverse connection strings](https://docs.microsoft.com/en-us/power-apps/developer/data-platform/xrm-tooling/use-connection-strings-xrm-tooling-connect) for other auth types

## Excel File Format

### Minimum Required Columns

| Table Name | Column Name | Column Type |
|------------|-------------|-------------|
| account    | status      | text        |
| contact    | score       | number      |
| lead       | priority    | choice      |

### All Supported Columns

| Column | Required | Description | Example |
|--------|----------|-------------|---------|
| **Table Name** | ✓ | Target Dataverse table | `account` |
| **Column Name** | ✓ | Display name for column | `Status` |
| **Column Type** | ✓ | Data type (see below) | `text` |
| Choice Options | | Semicolon-separated choices | `New;Active;Closed` or `1:New;2:Active` |
| Logical Name | | Explicit column logical name (without prefix) | `custom_status` |
| Table Logical Name | | Explicit table logical name if creating table | `custom_account` |
| Display Plural | | Table plural display name | `People` (not "Persons") |
| Description | | Column description | `Customer status field` |
| Required | | Required level | `None`, `Optional`, `Required`, `Recommended` |

**Defaults if optional columns not provided:**
- Logical names: Auto-generated from display names
- Display Plural: `{TableName}s`
- Description: Blank
- Required: `None`

### Supported Column Types

| Type | Aliases | Notes |
|------|---------|-------|
| **text** | string, single line of text | Max 100 chars |
| **memo** | multiple lines of text, multiline | Max 2000 chars |
| **number** | int, integer, whole number | Int32 |
| **decimal** | money, currency | 2 decimal precision |
| **float** | double, floating point number | Double precision |
| **date** | datetime, date and time | With time |
| **date only** | dateonly | Without time |
| **boolean** | bool, bit, yes/no, y/n | Two options |
| **choice** | picklist | Requires Choice Options |

**Choice Options Format:**
- Simple: `Low;Medium;High` (auto-assigns 1,2,3...)
- Explicit: `100:Low;200:Medium;300:High`

## Usage

### Main Menu

After scanning schemas, the application presents:

1. **Export to CSV** - Save new schemas for review
2. **Create in Dataverse** - Select solution and create schema (requires "yes" confirmation)
3. **Export and Create** - Both operations
4. **Re-scan** - Refresh schema status
5. **Exit**

### Schema Creation Flow

When creating schema:
1. Lists available unmanaged solutions
2. Shows publisher prefix
3. **Displays what will be created** including any missing tables
4. **Waits for explicit "yes" or "y" confirmation**
5. Creates columns in selected solution
6. Publishes customizations

**Critical**: The app will NOT create any schema without explicit user confirmation.

## Project Structure

```
├── Models/            # AppConfiguration, SchemaDefinition
├── Services/          # DataverseService, ExcelReaderService, CsvExportService
├── Constants/         # DataverseConstants
├── Interfaces/        # Service interfaces
├── Utils/             # GenerateSampleExcel
└── Program.cs         # Main application
```

## Troubleshooting

| Issue | Solution |
|-------|----------|
| **Connection fails** | Verify connection string, permissions, network access |
| **Schema creation fails** | Confirm table exists, naming conventions valid, unmanaged solution |
| **Excel read errors** | Check column headers match config, file not locked, path correct |
| **Missing table warning** | App auto-creates tables with confirmation - this is expected |
| **Publisher prefix mismatch** | Select correct solution matching your prefix |

## Security

- Never commit `appsettings.json` with credentials
- Use environment variables or Azure Key Vault in production
- Consider managed identities in Azure environments
- `.gitignore` excludes sensitive files by default

## Error Handling

Handles:
- Missing/invalid Excel files
- Dataverse connection failures
- Missing required columns
- Non-existent tables (auto-creates with confirmation)
- Invalid column types
- Publisher prefix mismatches
- File locking issues (retries with exponential backoff)

## Support

Create an issue in the GitHub repository for questions or bug reports.
