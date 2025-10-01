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
                    services.AddSingleton<ICsvExportService, CsvExportService>();

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
        private readonly ILogger<Application> _logger;
        private bool _exitRequested = false;

        public Application(
            AppConfiguration config,
            IDataverseService dataverseService,
            IExcelReaderService excelReader,
            ICsvExportService csvExporter,
            ILogger<Application> logger)
        {
            _config = config ?? throw new ArgumentNullException(nameof(config));
            _dataverseService = dataverseService ?? throw new ArgumentNullException(nameof(dataverseService));
            _excelReader = excelReader ?? throw new ArgumentNullException(nameof(excelReader));
            _csvExporter = csvExporter ?? throw new ArgumentNullException(nameof(csvExporter));
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
    }
}
