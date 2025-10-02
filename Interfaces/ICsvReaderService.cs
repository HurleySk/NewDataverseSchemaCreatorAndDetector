using DataverseSchemaManager.Models;

namespace DataverseSchemaManager.Interfaces
{
    /// <summary>
    /// Provides operations for reading schema definitions from CSV files.
    /// </summary>
    public interface ICsvReaderService
    {
        /// <summary>
        /// Reads schema definitions from a CSV file.
        /// </summary>
        /// <param name="filePath">Path to the CSV file.</param>
        /// <param name="config">Application configuration for column mappings.</param>
        /// <param name="cancellationToken">Token to cancel the operation.</param>
        /// <returns>List of schema definitions read from the CSV.</returns>
        Task<List<SchemaDefinition>> ReadSchemaDefinitionsAsync(string filePath, AppConfiguration config, CancellationToken cancellationToken = default);
    }
}
