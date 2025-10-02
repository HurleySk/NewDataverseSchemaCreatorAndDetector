using DataverseSchemaManager.Models;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace DataverseSchemaManager.Interfaces
{
    /// <summary>
    /// Defines operations for interacting with Microsoft Dataverse.
    /// </summary>
    public interface IDataverseService : IDisposable
    {
        /// <summary>
        /// Connects to a Dataverse environment using the specified connection string.
        /// </summary>
        /// <param name="connectionString">The connection string for the Dataverse environment.</param>
        /// <param name="cancellationToken">Token to cancel the operation.</param>
        /// <returns>True if connection successful, false otherwise.</returns>
        Task<bool> ConnectAsync(string connectionString, CancellationToken cancellationToken = default);

        /// <summary>
        /// Checks if schemas exist in the Dataverse environment and updates their status.
        /// </summary>
        /// <param name="schemas">The list of schema definitions to check.</param>
        /// <param name="cancellationToken">Token to cancel the operation.</param>
        Task CheckSchemaExistsAsync(List<SchemaDefinition> schemas, CancellationToken cancellationToken = default);

        /// <summary>
        /// Creates schemas (tables and columns) in the specified Dataverse solution.
        /// </summary>
        /// <param name="schemas">The list of schema definitions to create.</param>
        /// <param name="solutionName">The unique name of the solution to add the schemas to.</param>
        /// <param name="publisherPrefix">The publisher prefix for the solution.</param>
        /// <param name="cancellationToken">Token to cancel the operation.</param>
        Task CreateSchemaAsync(List<SchemaDefinition> schemas, string solutionName, string publisherPrefix, CancellationToken cancellationToken = default);

        /// <summary>
        /// Retrieves all unmanaged, visible solutions from the Dataverse environment.
        /// </summary>
        /// <param name="cancellationToken">Token to cancel the operation.</param>
        /// <returns>A list of solution entities.</returns>
        Task<List<Entity>> GetSolutionsAsync(CancellationToken cancellationToken = default);

        /// <summary>
        /// Gets the customization prefix for a specific publisher.
        /// </summary>
        /// <param name="publisherId">The ID of the publisher.</param>
        /// <param name="cancellationToken">Token to cancel the operation.</param>
        /// <returns>The publisher's customization prefix.</returns>
        Task<string?> GetPublisherPrefixAsync(Guid publisherId, CancellationToken cancellationToken = default);

        /// <summary>
        /// Validates a logical name against Dataverse naming rules.
        /// </summary>
        /// <param name="logicalName">The logical name to validate.</param>
        /// <returns>Tuple indicating if valid and error message if not.</returns>
        (bool IsValid, string? ErrorMessage) ValidateLogicalName(string logicalName);

        /// <summary>
        /// Checks if a column schema name already exists in a table.
        /// </summary>
        /// <param name="tableName">The table name to check.</param>
        /// <param name="columnSchemaName">The full column schema name (including prefix).</param>
        /// <param name="cancellationToken">Token to cancel the operation.</param>
        /// <returns>True if column exists, false otherwise.</returns>
        Task<bool> ColumnExistsAsync(string tableName, string columnSchemaName, CancellationToken cancellationToken = default);

        /// <summary>
        /// Checks if a table schema name already exists.
        /// </summary>
        /// <param name="tableSchemaName">The full table schema name (including prefix).</param>
        /// <param name="cancellationToken">Token to cancel the operation.</param>
        /// <returns>True if table exists, false otherwise.</returns>
        Task<bool> TableExistsAsync(string tableSchemaName, CancellationToken cancellationToken = default);

        /// <summary>
        /// Validates that a logical name doesn't already contain a prefix.
        /// </summary>
        /// <param name="explicitLogicalName">The explicit logical name from user.</param>
        /// <param name="publisherPrefix">The publisher prefix.</param>
        /// <returns>Tuple indicating if valid and error message if not.</returns>
        (bool IsValid, string? ErrorMessage) ValidatePrefixMismatch(string? explicitLogicalName, string publisherPrefix);
    }
}
