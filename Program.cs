using DataverseSchemaManager.Models;
using DataverseSchemaManager.Services;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace DataverseSchemaManager
{
    class Program
    {
        private static AppConfiguration _config = new();
        private static DataverseService _dataverseService = new();
        private static ExcelReaderService _excelReader = new();
        private static CsvExportService _csvExporter = new();

        static void Main(string[] args)
        {
            if (args.Length > 0 && args[0] == "--generate-sample")
            {
                Console.WriteLine("Creating sample Excel file...");
                GenerateSampleExcel.CreateSampleFile("sample_schema.xlsx");
                Console.WriteLine("Sample file created successfully!");
                return;
            }

            try
            {
                Console.WriteLine("===========================================");
                Console.WriteLine("  Dataverse Schema Creator and Detector   ");
                Console.WriteLine("===========================================");
                Console.WriteLine();

                LoadConfiguration();

                if (!ConnectToDataverse())
                {
                    return;
                }

                var schemas = LoadSchemaDefinitions();
                if (schemas == null || schemas.Count == 0)
                {
                    Console.WriteLine("No schema definitions found in the Excel file.");
                    return;
                }

                Console.WriteLine($"\nLoaded {schemas.Count} schema definitions from Excel.");

                _dataverseService.CheckSchemaExists(schemas);

                DisplaySchemaStatus(schemas);

                ShowMainMenu(schemas);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\nError: {ex.Message}");
            }
            finally
            {
                _dataverseService?.Dispose();
                Console.WriteLine("\nPress any key to exit...");
                Console.ReadKey();
            }
        }

        private static void LoadConfiguration()
        {
            var configFile = "appsettings.json";
            if (File.Exists(configFile))
            {
                var builder = new ConfigurationBuilder()
                    .SetBasePath(Directory.GetCurrentDirectory())
                    .AddJsonFile(configFile, optional: true);

                var configuration = builder.Build();
                configuration.Bind(_config);
            }
        }

        private static bool ConnectToDataverse()
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
                Console.WriteLine("Connection string is required.");
                return false;
            }

            if (_dataverseService.Connect(connectionString))
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

        private static List<SchemaDefinition>? LoadSchemaDefinitions()
        {
            string? excelPath = _config.ExcelFilePath;

            if (string.IsNullOrEmpty(excelPath) || !File.Exists(excelPath))
            {
                Console.WriteLine("\nEnter the path to the Excel file containing schema definitions:");
                excelPath = Console.ReadLine()?.Trim('"');
            }

            if (string.IsNullOrEmpty(excelPath) || !File.Exists(excelPath))
            {
                Console.WriteLine("Excel file not found.");
                return null;
            }

            try
            {
                return _excelReader.ReadSchemaDefinitions(excelPath, _config);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading Excel file: {ex.Message}");
                return null;
            }
        }

        private static void DisplaySchemaStatus(List<SchemaDefinition> schemas)
        {
            var existingCount = schemas.Count(s => s.ExistsInDataverse);
            var newCount = schemas.Count(s => !s.ExistsInDataverse);
            var errorCount = schemas.Count(s => !string.IsNullOrEmpty(s.ErrorMessage));

            Console.WriteLine("\n=== Schema Detection Results ===");
            Console.WriteLine($"Total schemas: {schemas.Count}");
            Console.WriteLine($"Existing in Dataverse: {existingCount}");
            Console.WriteLine($"New schemas to create: {newCount}");

            if (errorCount > 0)
            {
                Console.WriteLine($"Errors encountered: {errorCount}");
            }

            if (newCount > 0)
            {
                Console.WriteLine("\nNew schemas to be created:");
                foreach (var schema in schemas.Where(s => !s.ExistsInDataverse))
                {
                    Console.WriteLine($"  - Table: {schema.TableName}, Column: {schema.ColumnName}, Type: {schema.ColumnType}");
                    if (!string.IsNullOrEmpty(schema.ErrorMessage))
                    {
                        Console.WriteLine($"    Error: {schema.ErrorMessage}");
                    }
                }
            }
        }

        private static void ShowMainMenu(List<SchemaDefinition> schemas)
        {
            while (true)
            {
                Console.WriteLine("\n=== Main Menu ===");
                Console.WriteLine("1. Export new schemas to CSV");
                Console.WriteLine("2. Create new schemas in Dataverse");
                Console.WriteLine("3. Export new schemas to CSV and create in Dataverse");
                Console.WriteLine("4. Re-scan schemas");
                Console.WriteLine("5. Exit");
                Console.Write("\nSelect an option (1-5): ");

                var choice = Console.ReadLine();

                switch (choice)
                {
                    case "1":
                        ExportToCsv(schemas);
                        break;
                    case "2":
                        CreateSchemasInDataverse(schemas);
                        break;
                    case "3":
                        ExportToCsv(schemas);
                        CreateSchemasInDataverse(schemas);
                        break;
                    case "4":
                        _dataverseService.CheckSchemaExists(schemas);
                        DisplaySchemaStatus(schemas);
                        break;
                    case "5":
                        return;
                    default:
                        Console.WriteLine("Invalid option. Please try again.");
                        break;
                }
            }
        }

        private static void ExportToCsv(List<SchemaDefinition> schemas)
        {
            var newSchemas = schemas.Where(s => !s.ExistsInDataverse).ToList();

            if (newSchemas.Count == 0)
            {
                Console.WriteLine("No new schemas to export.");
                return;
            }

            Console.Write("\nEnter output CSV file path (or press Enter for default 'new_schemas.csv'): ");
            var outputPath = Console.ReadLine();

            if (string.IsNullOrEmpty(outputPath))
            {
                outputPath = "new_schemas.csv";
            }

            try
            {
                _csvExporter.ExportNewSchema(schemas, outputPath);
                Console.WriteLine($"Successfully exported {newSchemas.Count} new schemas to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error exporting to CSV: {ex.Message}");
            }
        }

        private static void CreateSchemasInDataverse(List<SchemaDefinition> schemas)
        {
            var newSchemas = schemas.Where(s => !s.ExistsInDataverse).ToList();

            if (newSchemas.Count == 0)
            {
                Console.WriteLine("No new schemas to create.");
                return;
            }

            try
            {
                var solutions = _dataverseService.GetSolutions();

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

                var publisherPrefix = _dataverseService.GetPublisherPrefix(publisherId);

                if (string.IsNullOrEmpty(publisherPrefix))
                {
                    Console.WriteLine("Could not retrieve publisher prefix.");
                    return;
                }

                Console.WriteLine($"\nSolution: {solutionName}");
                Console.WriteLine($"Publisher prefix: {publisherPrefix}");
                Console.WriteLine($"Creating {newSchemas.Count} new schemas...");

                _dataverseService.CreateSchema(newSchemas, solutionName, publisherPrefix);

                Console.WriteLine("\nSchema creation completed!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating schemas: {ex.Message}");
            }
        }
    }
}
