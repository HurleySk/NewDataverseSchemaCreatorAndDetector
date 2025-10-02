using DataverseSchemaManager.Models;

namespace DataverseSchemaManager.Interfaces
{
    /// <summary>
    /// Provides operations for generating CREATE template CSV files.
    /// </summary>
    public interface ITemplateGeneratorService
    {
        /// <summary>
        /// Generates a CREATE template CSV file with all necessary columns for schema creation.
        /// Pre-fills logical names and any provided optional fields.
        /// </summary>
        /// <param name="newSchemas">List of new schema definitions to include in template.</param>
        /// <param name="outputPath">Path where the CSV template should be created.</param>
        /// <param name="cancellationToken">Token to cancel the operation.</param>
        Task GenerateCreateTemplateAsync(List<SchemaDefinition> newSchemas, string outputPath, CancellationToken cancellationToken = default);
    }
}
