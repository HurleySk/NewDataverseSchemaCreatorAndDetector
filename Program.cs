using DataverseSchemaManager.Constants;
using DataverseSchemaManager.Interfaces;
using DataverseSchemaManager.Models;
using DataverseSchemaManager.Services;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Serilog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace DataverseSchemaManager
{
    class Program
    {
        static async Task<int> Main(string[] args)
        {
            // Build configuration
            var configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile(DataverseConstants.Files.DefaultConfigFile, optional: true)
                .Build();

            // Configure Serilog
            Log.Logger = new LoggerConfiguration()
                .ReadFrom.Configuration(configuration)
                .CreateLogger();

            try
            {
                Log.Information("===========================================");
                Log.Information("  Dataverse Schema Creator and Detector   ");
                Log.Information("===========================================");

                // Handle special command-line arguments
                if (args.Length > 0 && args[0] == "--generate-sample")
                {
                    Log.Information("Creating sample Excel file...");
                    GenerateSampleExcel.CreateSampleFile(DataverseConstants.Files.SampleExcelFileName);
                    Log.Information("Sample file created successfully!");
                    return 0;
                }

                // Build and run the host
                var host = CreateHostBuilder(args, configuration).Build();

                // Run the application
                var exitCode = await host.Services.GetRequiredService<Application>().RunAsync();

                return exitCode;
            }
            catch (Exception ex)
            {
                Log.Fatal(ex, "Application terminated unexpectedly");
                return 1;
            }
            finally
            {
                await Log.CloseAndFlushAsync();
            }
        }

        static IHostBuilder CreateHostBuilder(string[] args, IConfiguration configuration) =>
            Host.CreateDefaultBuilder(args)
                .UseSerilog()
                .ConfigureServices((context, services) =>
                {
                    // Register configuration
                    var appConfig = new AppConfiguration();
                    configuration.Bind(appConfig);
                    services.AddSingleton(appConfig);

                    // Register services
                    services.AddSingleton<IDataverseService, DataverseService>();
                    services.AddSingleton<IExcelReaderService, ExcelReaderService>();
                    services.AddSingleton<ICsvReaderService, CsvReaderService>();
                    services.AddSingleton<ICsvExportService, CsvExportService>();
                    services.AddSingleton<ITemplateGeneratorService, TemplateGeneratorService>();

                    // Register the main application
                    services.AddTransient<Application>();
                });
    }

    /// <summary>
    /// Main application logic with dependency injection support.
    /// </summary>
    public class Application
    {
        private readonly AppConfiguration _config;
        private readonly IDataverseService _dataverseService;
        private readonly IExcelReaderService _excelReader;
        private readonly ICsvExportService _csvExporter;
        private readonly ITemplateGeneratorService _templateGenerator;
        private readonly ILogger<Application> _logger;
        private bool _exitRequested = false;

        public Application(
            AppConfiguration config,
            IDataverseService dataverseService,
            IExcelReaderService excelReader,
            ICsvExportService csvExporter,
            ITemplateGeneratorService templateGenerator,
            ILogger<Application> logger)
        {
            _config = config ?? throw new ArgumentNullException(nameof(config));
            _dataverseService = dataverseService ?? throw new ArgumentNullException(nameof(dataverseService));
            _excelReader = excelReader ?? throw new ArgumentNullException(nameof(excelReader));
            _csvExporter = csvExporter ?? throw new ArgumentNullException(nameof(csvExporter));
            _templateGenerator = templateGenerator ?? throw new ArgumentNullException(nameof(templateGenerator));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        public async Task<int> RunAsync(CancellationToken cancellationToken = default)
        {
            try
            {
                if (!await ConnectToDataverseAsync(cancellationToken))
                {
                    return 1;
                }

                var schemas = await LoadSchemaDefinitionsAsync(cancellationToken);
                if (schemas == null || schemas.Count == 0)
                {
                    _logger.LogWarning("No schema definitions found in the Excel file");
                    Console.WriteLine("No schema definitions found in the Excel file.");
                    return 1;
                }

                _logger.LogInformation("Loaded {Count} schema definitions from Excel", schemas.Count);
                Console.WriteLine($"\nLoaded {schemas.Count} schema definitions from Excel.");

                await _dataverseService.CheckSchemaExistsAsync(schemas, cancellationToken);

                DisplaySchemaStatus(schemas);

                // Generate CREATE template for new schemas
                var newSchemas = schemas.Where(s => !s.ColumnExistsInDataverse).ToList();
                if (newSchemas.Count > 0)
                {
                    string templatePath = _config.OutputCsvPath ?? "create_template.csv";
                    await _templateGenerator.GenerateCreateTemplateAsync(newSchemas, templatePath, cancellationToken);
                    Console.WriteLine($"\nCREATE template generated: {templatePath}");
                    Console.WriteLine("Fill in the required fields (Table Name, Column Name, Column Type) and use this file to create schemas.");
                }

                await ShowMainMenuAsync(schemas, cancellationToken);

                return 0;
            }
            catch (OperationCanceledException)
            {
                _logger.LogWarning("Operation cancelled by user");
                Console.WriteLine("\nOperation cancelled.");
                return 130; // Standard exit code for SIGINT
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Unexpected error in application");
                Console.WriteLine($"\nError: {ex.Message}");
                return 1;
            }
            finally
            {
                _dataverseService?.Dispose();
                if (!_exitRequested)
                {
                    Console.WriteLine("\nPress any key to exit...");
                    Console.ReadKey();
                }
            }
        }

        private async Task<bool> ConnectToDataverseAsync(CancellationToken cancellationToken)
        {
            Console.WriteLine("Connecting to Dataverse...");

            string? connectionString = _config.ConnectionString;
            if (string.IsNullOrEmpty(connectionString))
            {
                Console.WriteLine("\nEnter Dataverse connection string:");
                Console.WriteLine("(Format: AuthType=OAuth;Url=https://yourorg.crm.dynamics.com;Username=user@org.com;Password=pass;RequireNewInstance=true)");
                connectionString = Console.ReadLine();
            }

            if (string.IsNullOrEmpty(connectionString))
            {
                _logger.LogError("Connection string is required");
                Console.WriteLine("Connection string is required.");
                return false;
            }

            var connected = await _dataverseService.ConnectAsync(connectionString, cancellationToken);

            if (connected)
            {
                Console.WriteLine("Successfully connected to Dataverse!");
                return true;
            }
            else
            {
                Console.WriteLine("Failed to connect to Dataverse. Please check your connection string.");
                return false;
            }
        }

        private async Task<List<SchemaDefinition>?> LoadSchemaDefinitionsAsync(CancellationToken cancellationToken)
        {
            string? excelPath = _config.ExcelFilePath;

            if (string.IsNullOrEmpty(excelPath) || !File.Exists(excelPath))
            {
                Console.WriteLine("\nEnter the path to the Excel file containing schema definitions:");
                excelPath = Console.ReadLine()?.Trim('"');
            }

            if (string.IsNullOrEmpty(excelPath) || !File.Exists(excelPath))
            {
                _logger.LogError("Excel file not found: {Path}", excelPath ?? "null");
                Console.WriteLine("Excel file not found.");
                return null;
            }

            try
            {
                return await _excelReader.ReadSchemaDefinitionsAsync(excelPath, _config, cancellationToken);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error reading Excel file");
                Console.WriteLine($"Error reading Excel file: {ex.Message}");
                return null;
            }
        }

        private void DisplaySchemaStatus(List<SchemaDefinition> schemas)
        {
            var existingCount = schemas.Count(s => s.ColumnExistsInDataverse);
            var newCount = schemas.Count(s => !s.ColumnExistsInDataverse);
            var missingTableErrors = schemas.Count(s => s.ErrorMessage?.Contains("does not exist") == true);
            var otherErrors = schemas.Count(s => !string.IsNullOrEmpty(s.ErrorMessage) && !s.ErrorMessage.Contains("does not exist"));

            Console.WriteLine("\n=== Schema Detection Results ===");
            Console.WriteLine($"Total schemas: {schemas.Count}");
            Console.WriteLine($"Existing in Dataverse: {existingCount}");
            Console.WriteLine($"New schemas to create: {newCount}");

            if (missingTableErrors > 0)
            {
                var missingTables = schemas
                    .Where(s => s.ErrorMessage?.Contains("does not exist") == true)
                    .Select(s => s.TableName.ToLower())
                    .Distinct()
                    .Count();
                Console.WriteLine($"Missing tables (will be created): {missingTables}");
            }

            if (otherErrors > 0)
            {
                Console.WriteLine($"Other errors encountered: {otherErrors}");
            }

            if (newCount > 0)
            {
                Console.WriteLine("\nNew schemas to be created:");
                foreach (var schema in schemas.Where(s => !s.ColumnExistsInDataverse))
                {
                    Console.WriteLine($"  - Table: {schema.TableName}, Column: {schema.ColumnName}, Type: {schema.ColumnType}");
                    if (!string.IsNullOrEmpty(schema.ErrorMessage) && !schema.ErrorMessage.Contains("does not exist"))
                    {
                        Console.WriteLine($"    Error: {schema.ErrorMessage}");
                    }
                }
            }
        }

        private async Task ShowMainMenuAsync(List<SchemaDefinition> schemas, CancellationToken cancellationToken)
        {
            while (!cancellationToken.IsCancellationRequested)
            {
                Console.WriteLine("\n=== Main Menu ===");
                Console.WriteLine("1. Export new schemas to CSV");
                Console.WriteLine("2. Create new schemas in Dataverse");
                Console.WriteLine("3. Export new schemas to CSV and create in Dataverse");
                Console.WriteLine("4. Re-scan schemas");
                Console.WriteLine("5. Exit");
                Console.Write("\nSelect an option (1-5): ");

                var choice = Console.ReadLine();

                try
                {
                    switch (choice)
                    {
                        case "1":
                            await ExportToCsvAsync(schemas, cancellationToken);
                            break;
                        case "2":
                            await CreateSchemasInDataverseAsync(schemas, cancellationToken);
                            break;
                        case "3":
                            await ExportToCsvAsync(schemas, cancellationToken);
                            await CreateSchemasInDataverseAsync(schemas, cancellationToken);
                            break;
                        case "4":
                            await _dataverseService.CheckSchemaExistsAsync(schemas, cancellationToken);
                            DisplaySchemaStatus(schemas);
                            break;
                        case "5":
                            _exitRequested = true;
                            return;
                        default:
                            Console.WriteLine("Invalid option. Please try again.");
                            break;
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error processing menu option {Choice}", choice);
                    Console.WriteLine($"Error: {ex.Message}");
                }
            }
        }

        private async Task ExportToCsvAsync(List<SchemaDefinition> schemas, CancellationToken cancellationToken)
        {
            var newSchemas = schemas.Where(s => !s.ColumnExistsInDataverse).ToList();

            if (newSchemas.Count == 0)
            {
                Console.WriteLine("No new schemas to export.");
                return;
            }

            string? outputPath = _config.OutputCsvPath;

            if (string.IsNullOrEmpty(outputPath))
            {
                Console.Write("\nEnter output CSV file path (or press Enter for default 'new_schemas.csv'): ");
                outputPath = Console.ReadLine();

                if (string.IsNullOrEmpty(outputPath))
                {
                    outputPath = DataverseConstants.Files.DefaultOutputCsvName;
                }
            }

            try
            {
                await _csvExporter.ExportNewSchemaAsync(schemas, outputPath, cancellationToken);
                Console.WriteLine($"Successfully exported {newSchemas.Count} new schemas to {outputPath}");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error exporting to CSV");
                Console.WriteLine($"Error exporting to CSV: {ex.Message}");
            }
        }

        private async Task CreateSchemasInDataverseAsync(List<SchemaDefinition> schemas, CancellationToken cancellationToken)
        {
            var newSchemas = schemas.Where(s => !s.ColumnExistsInDataverse).ToList();

            if (newSchemas.Count == 0)
            {
                Console.WriteLine("No new schemas to create.");
                return;
            }

            try
            {
                var solutions = await _dataverseService.GetSolutionsAsync(cancellationToken);

                if (solutions.Count == 0)
                {
                    Console.WriteLine("No unmanaged solutions found.");
                    return;
                }

                Console.WriteLine("\n=== Available Solutions ===");
                for (int i = 0; i < solutions.Count; i++)
                {
                    var friendlyName = solutions[i].GetAttributeValue<string>("friendlyname");
                    var uniqueName = solutions[i].GetAttributeValue<string>("uniquename");
                    Console.WriteLine($"{i + 1}. {friendlyName} ({uniqueName})");
                }

                Console.Write("\nSelect a solution (enter number): ");
                if (!int.TryParse(Console.ReadLine(), out int solutionIndex) ||
                    solutionIndex < 1 || solutionIndex > solutions.Count)
                {
                    Console.WriteLine("Invalid selection.");
                    return;
                }

                var selectedSolution = solutions[solutionIndex - 1];
                var solutionName = selectedSolution.GetAttributeValue<string>("uniquename");
                var publisherId = selectedSolution.GetAttributeValue<Microsoft.Xrm.Sdk.EntityReference>("publisherid").Id;

                var publisherPrefix = await _dataverseService.GetPublisherPrefixAsync(publisherId, cancellationToken);

                if (string.IsNullOrEmpty(publisherPrefix))
                {
                    Console.WriteLine("Could not retrieve publisher prefix.");
                    return;
                }

                Console.WriteLine($"\nSolution: {solutionName}");
                Console.WriteLine($"Publisher prefix: {publisherPrefix}");

                // Validate schemas before creation
                Console.WriteLine("\nValidating schemas...");
                var validationErrors = await ValidateSchemasAsync(newSchemas, publisherPrefix, cancellationToken);

                if (validationErrors.Any())
                {
                    Console.WriteLine($"\n=== VALIDATION ERRORS ({validationErrors.Count}) ===");
                    foreach (var error in validationErrors)
                    {
                        Console.WriteLine($"  ✗ {error}");
                    }
                    Console.WriteLine("\nPlease fix the errors in your Excel file and try again.");
                    return;
                }

                Console.WriteLine("✓ All schemas validated successfully");

                // Show preview of what will be created
                await ShowSchemaPreviewAsync(newSchemas, publisherPrefix, cancellationToken);

                var missingTables = newSchemas
                    .Where(s => s.ErrorMessage?.Contains("does not exist") == true)
                    .Select(s => s.TableName.ToLower())
                    .Distinct()
                    .ToList();

                if (missingTables.Any())
                {
                    Console.WriteLine($"\n=== WARNING: Missing Tables ===");
                    Console.WriteLine($"The following {missingTables.Count} table(s) do not exist and will be created:");
                    foreach (var table in missingTables)
                    {
                        Console.WriteLine($"  - {table}");
                    }
                }

                Console.WriteLine($"\nTotal new schemas to create: {newSchemas.Count}");
                Console.Write("\nDo you want to proceed with schema creation? (yes/no): ");
                var confirmation = Console.ReadLine()?.Trim().ToLower();

                if (confirmation != "yes" && confirmation != "y")
                {
                    _logger.LogInformation("Schema creation cancelled by user");
                    Console.WriteLine("Schema creation cancelled.");
                    return;
                }

                Console.WriteLine("\nCreating schemas...");
                await _dataverseService.CreateSchemaAsync(newSchemas, solutionName, publisherPrefix, cancellationToken);

                Console.WriteLine("\nSchema creation completed!");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating schemas");
                Console.WriteLine($"Error creating schemas: {ex.Message}");
            }
        }

        private async Task<List<string>> ValidateSchemasAsync(List<SchemaDefinition> schemas, string publisherPrefix, CancellationToken cancellationToken)
        {
            var errors = new List<string>();

            foreach (var schema in schemas)
            {
                // Validate required fields for CREATION
                if (string.IsNullOrWhiteSpace(schema.TableName))
                {
                    errors.Add($"Table Name (display) is REQUIRED for creation. Row: Table Logical Name='{schema.TableLogicalName}'");
                }

                if (string.IsNullOrWhiteSpace(schema.ColumnName))
                {
                    errors.Add($"Column Name (display) is REQUIRED for creation. Row: Column Logical Name='{schema.LogicalName}'");
                }

                if (string.IsNullOrWhiteSpace(schema.ColumnType))
                {
                    errors.Add($"Column Type is REQUIRED for creation. Row: Column Logical Name='{schema.LogicalName}'");
                }

                // Validate required logical names are not empty
                if (string.IsNullOrWhiteSpace(schema.LogicalName))
                {
                    errors.Add($"Column Logical Name is REQUIRED and cannot be empty");
                }
                else
                {
                    var (isValid, errorMessage) = _dataverseService.ValidateLogicalName(schema.LogicalName);
                    if (!isValid)
                    {
                        errors.Add($"Column '{schema.ColumnName}': {errorMessage}");
                    }

                    // Check for prefix mismatch
                    var (isPrefixValid, prefixError) = _dataverseService.ValidatePrefixMismatch(schema.LogicalName, publisherPrefix);
                    if (!isPrefixValid)
                    {
                        errors.Add($"Column '{schema.ColumnName}': {prefixError}");
                    }
                }

                if (string.IsNullOrWhiteSpace(schema.TableLogicalName))
                {
                    errors.Add($"Table '{schema.TableName}': Table Logical Name is REQUIRED and cannot be empty");
                }
                else
                {
                    var (isValid, errorMessage) = _dataverseService.ValidateLogicalName(schema.TableLogicalName);
                    if (!isValid)
                    {
                        errors.Add($"Table '{schema.TableName}': {errorMessage}");
                    }

                    // Check for prefix mismatch
                    var (isPrefixValid, prefixError) = _dataverseService.ValidatePrefixMismatch(schema.TableLogicalName, publisherPrefix);
                    if (!isPrefixValid)
                    {
                        errors.Add($"Table '{schema.TableName}': {prefixError}");
                    }
                }

                // Validate type-specific required fields
                if (!string.IsNullOrWhiteSpace(schema.ColumnType))
                {
                    var columnType = schema.ColumnType.Trim().ToLower();

                    // Choice columns MUST have Choice Options
                    if (columnType == "choice" || columnType == "picklist")
                    {
                        if (string.IsNullOrWhiteSpace(schema.ChoiceOptions))
                        {
                            errors.Add($"Column '{schema.ColumnName}': Choice Options are REQUIRED for choice/picklist columns");
                        }
                    }

                    // Lookup columns MUST have Lookup Target Table
                    if (columnType.StartsWith("lookup"))
                    {
                        if (string.IsNullOrWhiteSpace(schema.LookupTargetTable))
                        {
                            errors.Add($"Column '{schema.ColumnName}': Lookup Target Table is REQUIRED for lookup columns");
                        }
                        else
                        {
                            try
                            {
                                var targetExists = await _dataverseService.TableExistsAsync(schema.LookupTargetTable, cancellationToken);
                                if (!targetExists)
                                {
                                    errors.Add($"Lookup column '{schema.ColumnName}': Target table '{schema.LookupTargetTable}' does not exist in Dataverse");
                                }
                            }
                            catch (Exception ex)
                            {
                                _logger.LogWarning(ex, "Could not validate lookup target table: {Error}", ex.Message);
                            }
                        }
                    }

                    // Customer columns MUST have Customer Target Tables
                    if (columnType.StartsWith("customer"))
                    {
                        if (string.IsNullOrWhiteSpace(schema.CustomerTargetTables))
                        {
                            errors.Add($"Column '{schema.ColumnName}': Customer Target Tables are REQUIRED for customer columns");
                        }
                        else
                        {
                            var targetTables = schema.CustomerTargetTables.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                            foreach (var targetTable in targetTables)
                            {
                                try
                                {
                                    var targetExists = await _dataverseService.TableExistsAsync(targetTable, cancellationToken);
                                    if (!targetExists)
                                    {
                                        errors.Add($"Customer column '{schema.ColumnName}': Target table '{targetTable}' does not exist in Dataverse");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    _logger.LogWarning(ex, "Could not validate customer target table: {Error}", ex.Message);
                                }
                            }
                        }
                    }
                }

                // Build the final schema names (logical names are now required)
                var columnLogicalName = schema.LogicalName.ToLower().Replace(" ", "_");
                var columnSchemaName = $"{publisherPrefix}_{columnLogicalName}";

                var tableLogicalName = schema.TableLogicalName.ToLower().Replace(" ", "_");
                var tableSchemaName = $"{publisherPrefix}_{tableLogicalName}";

                // Check for conflicts with existing schemas
                if (schema.TableExistsInDataverse)
                {
                    // Table exists, check if column already exists
                    try
                    {
                        var columnExists = await _dataverseService.ColumnExistsAsync(schema.TableName, columnSchemaName, cancellationToken);
                        if (columnExists)
                        {
                            errors.Add($"Column '{columnSchemaName}' already exists in table '{schema.TableName}'");
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, "Could not check if column exists: {Error}", ex.Message);
                    }
                }
                else
                {
                    // Table doesn't exist, check if table schema name conflicts
                    try
                    {
                        var tableExists = await _dataverseService.TableExistsAsync(tableSchemaName, cancellationToken);
                        if (tableExists)
                        {
                            errors.Add($"Table '{tableSchemaName}' already exists but was not detected during initial scan");
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, "Could not check if table exists: {Error}", ex.Message);
                    }
                }
            }

            return errors;
        }

        private async Task ShowSchemaPreviewAsync(List<SchemaDefinition> schemas, string publisherPrefix, CancellationToken cancellationToken)
        {
            Console.WriteLine("\n=== Schema Preview (Exact Names) ===");

            var groupedByTable = schemas.GroupBy(s => s.TableName.ToLower()).OrderBy(g => g.Key);

            foreach (var tableGroup in groupedByTable)
            {
                var firstSchema = tableGroup.First();
                var tableLogicalName = !string.IsNullOrWhiteSpace(firstSchema.TableLogicalName)
                    ? firstSchema.TableLogicalName.ToLower().Replace(" ", "_")
                    : firstSchema.TableName.ToLower().Replace(" ", "_");
                var tableSchemaName = $"{publisherPrefix}_{tableLogicalName}";

                var displayCollectionName = !string.IsNullOrWhiteSpace(firstSchema.TableDisplayCollectionName)
                    ? firstSchema.TableDisplayCollectionName
                    : $"{firstSchema.TableName}s";

                if (!firstSchema.TableExistsInDataverse)
                {
                    Console.WriteLine($"\n[NEW TABLE] {firstSchema.TableName}");
                    Console.WriteLine($"  Schema Name: {tableSchemaName}");
                    Console.WriteLine($"  Display Plural: {displayCollectionName}");
                    Console.WriteLine($"  Columns ({tableGroup.Count()}):");
                }
                else
                {
                    Console.WriteLine($"\n[EXISTING TABLE] {firstSchema.TableName}");
                    Console.WriteLine($"  New Columns ({tableGroup.Count()}):");
                }

                foreach (var schema in tableGroup.OrderBy(s => s.ColumnName))
                {
                    var columnLogicalName = !string.IsNullOrWhiteSpace(schema.LogicalName)
                        ? schema.LogicalName.ToLower().Replace(" ", "_")
                        : schema.ColumnName.ToLower().Replace(" ", "_");
                    var columnSchemaName = $"{publisherPrefix}_{columnLogicalName}";

                    var requiredLevel = string.IsNullOrWhiteSpace(schema.Required) ? "None" : schema.Required;
                    var description = string.IsNullOrWhiteSpace(schema.Description) ? "(no description)" : schema.Description;

                    Console.WriteLine($"    • {schema.ColumnName}");
                    Console.WriteLine($"      Schema Name: {columnSchemaName}");
                    Console.WriteLine($"      Type: {schema.ColumnType}");
                    Console.WriteLine($"      Required: {requiredLevel}");
                    if (!string.IsNullOrWhiteSpace(schema.Description))
                    {
                        Console.WriteLine($"      Description: {description}");
                    }
                    if (!string.IsNullOrWhiteSpace(schema.ChoiceOptions))
                    {
                        Console.WriteLine($"      Options: {schema.ChoiceOptions}");
                    }

                    // Show lookup relationship information
                    var columnType = schema.ColumnType.ToLower();
                    if (columnType.StartsWith("lookup") && !string.IsNullOrWhiteSpace(schema.LookupTargetTable))
                    {
                        var relationshipName = !string.IsNullOrWhiteSpace(schema.LookupRelationshipName)
                            ? $"{publisherPrefix}_{schema.LookupRelationshipName.ToLower().Replace(" ", "_")}"
                            : $"{publisherPrefix}_{tableLogicalName}_{columnLogicalName}";

                        Console.WriteLine($"      → References: {schema.LookupTargetTable}");
                        Console.WriteLine($"      Relationship: {relationshipName}");
                    }

                    // Show customer relationship information
                    if (columnType.StartsWith("customer") && !string.IsNullOrWhiteSpace(schema.CustomerTargetTables))
                    {
                        var relationshipName = !string.IsNullOrWhiteSpace(schema.LookupRelationshipName)
                            ? $"{publisherPrefix}_{schema.LookupRelationshipName.ToLower().Replace(" ", "_")}"
                            : $"{publisherPrefix}_{tableLogicalName}_{columnLogicalName}";

                        Console.WriteLine($"      → References: {schema.CustomerTargetTables}");
                        Console.WriteLine($"      Relationship: {relationshipName}_[target]");
                    }
                }
            }

            await Task.CompletedTask;
        }
    }
}
