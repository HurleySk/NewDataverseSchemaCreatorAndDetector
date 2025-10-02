# Dataverse Schema Manager

A .NET console tool for detecting existing Dataverse schema and creating new schema from simple input files.

## Two-Phase Architecture

### Phase 1: ASSESS (Detection + Template Generation)
- **Input**: Simple Excel/CSV with 2 required columns
- **Process**: Detect what exists in Dataverse
- **Output**: CREATE template CSV (only NEW schemas)

### Phase 2: CREATE (Schema Creation)
- **Input**: Filled CREATE CSV
- **Process**: Validate + Create schema
- **Output**: Schema in Dataverse

---

## Phase 1: ASSESS

### Input Requirements

**Required Columns:**
- **Table Logical Name** - Logical name of table (e.g., `account`, `new_customtable`)
- **Column Logical Name** - Logical name of column (e.g., `new_field`, `custom_status`)

**Optional Columns (for convenience):**
- **Column Type** - Passes through to template (e.g., `text`, `number`, `choice`)
- **Choice Options** - For choice detection + passes through to template
- **Include** - Filter rows (`Yes`/`Y`/`True`/`1` to process, others skipped)

**Supported File Formats:** Excel (.xlsx) or CSV

### What Happens

1. Connects to Dataverse
2. Checks if table + column exist using **logical names only**
3. Identifies NEW schemas (not in Dataverse)
4. Generates CREATE template CSV with:
   - Pre-filled: Logical names, Column Type (if provided), Choice Options (if provided)
   - **BLANK fields YOU MUST FILL**: Table Name, Column Name, and type-specific fields

### Output: CREATE Template CSV

Shows ONLY new schemas.
Contains ALL columns needed for creation.
Pre-fills what it can, leaves blank what you must provide.

**Template columns:**
- Table Logical Name ✓ (pre-filled)
- Column Logical Name ✓ (pre-filled)
- Table Name (BLANK - you fill)
- Column Name (BLANK - you fill)
- Column Type (pre-filled if provided, else BLANK)
- Choice Options (pre-filled if provided, else BLANK)
- Lookup Target Table (BLANK)
- Lookup Relationship Name (BLANK)
- Customer Target Tables (BLANK)
- Display Plural (BLANK)
- Description (BLANK)
- Required (BLANK)

---

## Phase 2: CREATE

### Input: CREATE CSV Requirements

**Required Columns (must be filled):**
- Table Logical Name
- Column Logical Name
- **Table Name** (display name) - REQUIRED
- **Column Name** (display name) - REQUIRED
- **Column Type** - REQUIRED

**Type-Specific REQUIRED Fields:**
- IF type = `choice` or `picklist` → **Choice Options REQUIRED**
- IF type = `lookup` → **Lookup Target Table REQUIRED**
- IF type = `customer` → **Customer Target Tables REQUIRED** (comma-separated)

**Optional Fields:**
- Description
- Required (None/Optional/Required/Recommended)
- Display Plural
- Lookup Relationship Name
- Customer Target Tables

### What Happens

1. Reads CREATE CSV
2. Validates ALL required fields are present
3. Validates logical names format
4. Validates lookup/customer target tables exist in Dataverse
5. Connects to Dataverse
6. Shows detailed preview of what will be created
7. **Waits for explicit confirmation ("yes" or "y")**
8. Creates schema in selected solution
9. Publishes customizations

---

## Supported Column Types

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
| **choice** | picklist | **Requires Choice Options** |
| **lookup** | | **Requires Lookup Target Table** |
| **customer** | | **Requires Customer Target Tables** |

### Choice Options Format
- Simple: `Low;Medium;High` (auto-assigns values 1,2,3...)
- Explicit: `100:Low;200:Medium;300:High`

### Lookup Columns
- Specify single target table logical name
- Creates one-to-many relationship automatically
- Target table must exist in Dataverse

### Customer Columns (Polymorphic)
- Specify comma-separated target tables (e.g., `account,contact`)
- Creates polymorphic relationships to each target
- All target tables must exist in Dataverse

---

## Quick Start

### Prerequisites
- .NET 9.0 SDK
- Dataverse environment access
- Input file (Excel or CSV)

### Installation

```bash
git clone https://github.com/yourusername/NewDataverseSchemaCreatorAndDetector.git
cd NewDataverseSchemaCreatorAndDetector
dotnet build
```

### Configuration

1. Copy template:
```bash
cp appsettings.template.json appsettings.json
```

2. Edit `appsettings.json`:
```json
{
  "ExcelFilePath": "path/to/input.xlsx",
  "ConnectionString": "AuthType=OAuth;Username=user@org.com;Url=https://yourorg.crm.dynamics.com/;AppId=51f81489-12ee-4a9e-aaae-a2591f45987d;RedirectUri=http://localhost",
  "OutputCsvPath": "create_template.csv",

  "TableLogicalNameColumn": "Table Logical Name",
  "LogicalNameColumn": "Column Logical Name",
  "ChoiceOptionsColumn": "Choice Options",
  "ColumnTypeColumn": "Column Type",
  "IncludeColumn": "Include"
}
```

### Run

```bash
dotnet run
```

The app will:
1. Read your input file
2. Detect existing schemas
3. Generate `create_template.csv` with new schemas
4. Show main menu for next steps

---

## Usage Flow

### Step 1: Prepare Input File

Create Excel or CSV with minimum 2 columns:

| Table Logical Name | Column Logical Name | Column Type | Choice Options | Include |
|-------------------|-------------------|-------------|----------------|---------|
| account | new_customfield | text | | Yes |
| contact | custom_score | number | | Yes |
| lead | custom_priority | choice | Low;Medium;High | Yes |

### Step 2: Run Assessment

```bash
dotnet run
```

Output:
```
Loaded 3 schema definitions from Excel.
=== Schema Detection Results ===
Total schemas: 3
Existing in Dataverse: 0
New schemas to create: 3

CREATE template generated: create_template.csv
Fill in the required fields (Table Name, Column Name, Column Type) and use this file to create schemas.
```

### Step 3: Fill CREATE Template

Open `create_template.csv` and fill in:
- Table Name (display)
- Column Name (display)
- Any missing type-specific fields

Example:
| Table Logical Name | Column Logical Name | Table Name | Column Name | Column Type | Choice Options |
|-------------------|-------------------|------------|-------------|-------------|----------------|
| account | new_customfield | Account | Custom Field | text | |
| contact | custom_score | Contact | Custom Score | number | |
| lead | custom_priority | Lead | Priority Level | choice | Low;Medium;High |

### Step 4: Create Schemas

From main menu:
```
=== Main Menu ===
1. Export new schemas to CSV
2. Create new schemas in Dataverse
3. Export new schemas to CSV and create in Dataverse
4. Re-scan schemas
5. Exit
```

Select option 2 or 3, choose solution, confirm creation.

---

## Validation

### Detection Phase (Assess)
- Validates logical name format (lowercase, alphanumeric, underscores)
- No other validation needed

### Creation Phase (Create)
- **Table Name** (display) - must be filled
- **Column Name** (display) - must be filled
- **Column Type** - must be filled
- **Logical names** - format validation (lowercase, alphanumeric, underscores only)
- **Prefix validation** - ensures logical names don't already contain prefix
- **Conflict detection** - checks for duplicate tables/columns
- **Type-specific validation**:
  - Choice → Choice Options required
  - Lookup → Lookup Target Table required and must exist
  - Customer → Customer Target Tables required and all must exist

All validation happens **BEFORE** any schema creation attempts.

---

## Configuration Reference

### Column Mappings

If your Excel/CSV uses different column names, update these:

```json
{
  "TableLogicalNameColumn": "Your Table Logical Name Header",
  "LogicalNameColumn": "Your Column Logical Name Header",
  "ColumnTypeColumn": "Your Type Header",
  "ChoiceOptionsColumn": "Your Options Header",
  "IncludeColumn": "Your Include Flag Header"
}
```

### Connection String

OAuth with token caching (recommended):
```
AuthType=OAuth;Username=user@org.com;Url=https://yourorg.crm.dynamics.com/;AppId=51f81489-12ee-4a9e-aaae-a2591f45987d;RedirectUri=http://localhost;TokenCacheStorePath=C:\\path\\to\\tokencache
```

See [Dataverse connection strings](https://docs.microsoft.com/en-us/power-apps/developer/data-platform/xrm-tooling/use-connection-strings-xrm-tooling-connect) for other auth types.

---

## Troubleshooting

| Issue | Solution |
|-------|----------|
| **"Required columns not found"** | Ensure your Excel/CSV has "Table Logical Name" and "Column Logical Name" columns |
| **"Table Name (display) is REQUIRED for creation"** | Fill in Table Name column in CREATE CSV before creating |
| **"Choice Options are REQUIRED for choice columns"** | Add Choice Options for any choice/picklist columns |
| **"Lookup Target Table does not exist"** | Ensure target table exists in Dataverse before creating lookup |
| **Connection fails** | Verify connection string, permissions, network access |
| **File locked error** | Close Excel file or wait for OneDrive sync |

---

## Security

- Never commit `appsettings.json` with credentials
- Use environment variables or Azure Key Vault in production
- Consider managed identities in Azure environments
- `.gitignore` excludes sensitive files by default

---

## Project Structure

```
├── Models/               # SchemaDefinition, AppConfiguration
├── Services/             # DataverseService, ExcelReaderService, CsvReaderService, TemplateGeneratorService
├── Constants/            # DataverseConstants
├── Interfaces/           # Service interfaces
└── Program.cs            # Main application with two-phase flow
```

---

## Support

Create an issue in the GitHub repository for questions or bug reports.
