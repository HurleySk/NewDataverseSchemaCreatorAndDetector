using DataverseSchemaManager.Models;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace DataverseSchemaManager.Interfaces
{
    /// <summary>
    /// Defines operations for reading schema definitions from Excel files.
    /// </summary>
    public interface IExcelReaderService
    {
        /// <summary>
        /// Reads schema definitions from an Excel file.
        /// </summary>
        /// <param name="filePath">The path to the Excel file.</param>
        /// <param name="config">The application configuration containing column mappings.</param>
        /// <param name="cancellationToken">Token to cancel the operation.</param>
        /// <returns>A list of schema definitions read from the Excel file.</returns>
        Task<List<SchemaDefinition>> ReadSchemaDefinitionsAsync(string filePath, AppConfiguration config, CancellationToken cancellationToken = default);
    }
}
