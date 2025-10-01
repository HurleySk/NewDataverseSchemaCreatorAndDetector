using CsvHelper;
using CsvHelper.Configuration;
using DataverseSchemaManager.Interfaces;
using DataverseSchemaManager.Models;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace DataverseSchemaManager.Services
{
    /// <summary>
    /// Provides operations for exporting schema definitions to CSV format.
    /// </summary>
    public class CsvExportService : ICsvExportService
    {
        private readonly ILogger<CsvExportService> _logger;

        public CsvExportService(ILogger<CsvExportService> logger)
        {
            _logger = logger ?? throw new System.ArgumentNullException(nameof(logger));
        }

        /// <inheritdoc/>
        public async Task ExportNewSchemaAsync(List<SchemaDefinition> schemas, string outputPath, CancellationToken cancellationToken = default)
        {
            var newSchemas = schemas.Where(s => !s.ColumnExistsInDataverse).ToList();

            _logger.LogInformation("Exporting {Count} new schemas to CSV: {OutputPath}", newSchemas.Count, outputPath);

            await ExportToCsvAsync(newSchemas, outputPath, cancellationToken);

            _logger.LogInformation("Successfully exported {Count} new schemas to {OutputPath}", newSchemas.Count, outputPath);
        }

        /// <inheritdoc/>
        public async Task ExportAllSchemaAsync(List<SchemaDefinition> schemas, string outputPath, CancellationToken cancellationToken = default)
        {
            _logger.LogInformation("Exporting {Count} total schemas to CSV: {OutputPath}", schemas.Count, outputPath);

            await ExportToCsvAsync(schemas, outputPath, cancellationToken);

            _logger.LogInformation("Successfully exported {Count} total schemas to {OutputPath}", schemas.Count, outputPath);
        }

        private async Task ExportToCsvAsync(List<SchemaDefinition> schemas, string outputPath, CancellationToken cancellationToken)
        {
            await Task.Run(() =>
            {
                using var writer = new StreamWriter(outputPath);
                using var csv = new CsvWriter(writer, new CsvConfiguration(CultureInfo.InvariantCulture));

                csv.WriteRecords(schemas);
            }, cancellationToken);
        }
    }
}
