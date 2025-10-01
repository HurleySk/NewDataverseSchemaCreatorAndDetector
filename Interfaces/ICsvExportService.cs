using DataverseSchemaManager.Models;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace DataverseSchemaManager.Interfaces
{
    /// <summary>
    /// Defines operations for exporting schema definitions to CSV format.
    /// </summary>
    public interface ICsvExportService
    {
        /// <summary>
        /// Exports new (non-existing) schemas to a CSV file.
        /// </summary>
        /// <param name="schemas">The list of all schema definitions.</param>
        /// <param name="outputPath">The path where the CSV file should be saved.</param>
        /// <param name="cancellationToken">Token to cancel the operation.</param>
        Task ExportNewSchemaAsync(List<SchemaDefinition> schemas, string outputPath, CancellationToken cancellationToken = default);

        /// <summary>
        /// Exports all schemas to a CSV file.
        /// </summary>
        /// <param name="schemas">The list of all schema definitions.</param>
        /// <param name="outputPath">The path where the CSV file should be saved.</param>
        /// <param name="cancellationToken">Token to cancel the operation.</param>
        Task ExportAllSchemaAsync(List<SchemaDefinition> schemas, string outputPath, CancellationToken cancellationToken = default);
    }
}
