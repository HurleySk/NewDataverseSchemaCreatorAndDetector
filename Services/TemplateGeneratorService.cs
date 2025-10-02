using CsvHelper;
using CsvHelper.Configuration;
using DataverseSchemaManager.Interfaces;
using DataverseSchemaManager.Models;
using Microsoft.Extensions.Logging;
using System.Globalization;

namespace DataverseSchemaManager.Services
{
    /// <summary>
    /// Provides operations for generating CREATE template CSV files.
    /// </summary>
    public class TemplateGeneratorService : ITemplateGeneratorService
    {
        private readonly ILogger<TemplateGeneratorService> _logger;

        public TemplateGeneratorService(ILogger<TemplateGeneratorService> logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <inheritdoc/>
        public async Task GenerateCreateTemplateAsync(List<SchemaDefinition> newSchemas, string outputPath, CancellationToken cancellationToken = default)
        {
            if (newSchemas == null || newSchemas.Count == 0)
            {
                _logger.LogWarning("No new schemas to generate template for");
                throw new InvalidOperationException("No new schemas provided for template generation.");
            }

            _logger.LogInformation("Generating CREATE template CSV with {Count} schemas to {Path}", newSchemas.Count, outputPath);

            var csvConfig = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                TrimOptions = TrimOptions.Trim
            };

            using var writer = new StreamWriter(outputPath);
            using var csv = new CsvWriter(writer, csvConfig);

            // Write header row with all columns needed for creation
            csv.WriteField("Table Logical Name");
            csv.WriteField("Column Logical Name");
            csv.WriteField("Table Name");
            csv.WriteField("Column Name");
            csv.WriteField("Column Type");
            csv.WriteField("Choice Options");
            csv.WriteField("Lookup Target Table");
            csv.WriteField("Lookup Relationship Name");
            csv.WriteField("Customer Target Tables");
            csv.WriteField("Display Plural");
            csv.WriteField("Description");
            csv.WriteField("Required");
            await csv.NextRecordAsync();

            // Write data rows
            foreach (var schema in newSchemas)
            {
                cancellationToken.ThrowIfCancellationRequested();

                // Pre-fill: Logical names (always)
                csv.WriteField(schema.TableLogicalName);
                csv.WriteField(schema.LogicalName);

                // Leave blank: Display names (user must fill for creation)
                csv.WriteField(schema.TableName ?? string.Empty);
                csv.WriteField(schema.ColumnName ?? string.Empty);

                // Pass through if provided: Column Type (convenience)
                csv.WriteField(schema.ColumnType ?? string.Empty);

                // Pass through if provided: Choice Options
                csv.WriteField(schema.ChoiceOptions ?? string.Empty);

                // Leave blank: Type-specific fields
                csv.WriteField(schema.LookupTargetTable ?? string.Empty);
                csv.WriteField(schema.LookupRelationshipName ?? string.Empty);
                csv.WriteField(schema.CustomerTargetTables ?? string.Empty);

                // Leave blank: Optional metadata
                csv.WriteField(schema.TableDisplayCollectionName ?? string.Empty);
                csv.WriteField(schema.Description ?? string.Empty);
                csv.WriteField(schema.Required ?? string.Empty);

                await csv.NextRecordAsync();
            }

            await writer.FlushAsync();

            _logger.LogInformation("CREATE template CSV generated successfully: {Path}", outputPath);
        }
    }
}
