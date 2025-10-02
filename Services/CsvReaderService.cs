using CsvHelper;
using CsvHelper.Configuration;
using DataverseSchemaManager.Interfaces;
using DataverseSchemaManager.Models;
using Microsoft.Extensions.Logging;
using System.Globalization;

namespace DataverseSchemaManager.Services
{
    /// <summary>
    /// Provides operations for reading schema definitions from CSV files.
    /// </summary>
    public class CsvReaderService : ICsvReaderService
    {
        private readonly ILogger<CsvReaderService> _logger;

        public CsvReaderService(ILogger<CsvReaderService> logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <inheritdoc/>
        public async Task<List<SchemaDefinition>> ReadSchemaDefinitionsAsync(string filePath, AppConfiguration config, CancellationToken cancellationToken = default)
        {
            if (!File.Exists(filePath))
            {
                _logger.LogError("CSV file not found: {FilePath}", filePath);
                throw new FileNotFoundException($"CSV file not found: {filePath}");
            }

            _logger.LogInformation("Reading schema definitions from CSV file: {FilePath}", filePath);

            var schemas = new List<SchemaDefinition>();

            using var reader = new StreamReader(filePath);
            var csvConfig = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                HeaderValidated = null,
                MissingFieldFound = null,
                TrimOptions = TrimOptions.Trim
            };

            using var csv = new CsvReader(reader, csvConfig);

            await csv.ReadAsync();
            csv.ReadHeader();
            var headerRecord = csv.HeaderRecord;

            if (headerRecord == null)
            {
                _logger.LogError("CSV file has no header row");
                throw new InvalidOperationException("CSV file has no header row.");
            }

            // Find column indices
            var tableLogicalNameColumn = FindColumnIndex(headerRecord, config.TableLogicalNameColumn ?? "Table Logical Name");
            var logicalNameColumn = FindColumnIndex(headerRecord, config.LogicalNameColumn ?? "Logical Name");

            // Validate required columns exist
            if (tableLogicalNameColumn == -1 || logicalNameColumn == -1)
            {
                var missing = new List<string>();
                if (tableLogicalNameColumn == -1) missing.Add(config.TableLogicalNameColumn ?? "Table Logical Name");
                if (logicalNameColumn == -1) missing.Add(config.LogicalNameColumn ?? "Logical Name");

                _logger.LogError("Required columns not found: {MissingColumns}", string.Join(", ", missing));
                throw new InvalidOperationException($"Required columns not found in CSV: {string.Join(", ", missing)}.");
            }

            // Optional columns
            var choiceOptionsColumn = !string.IsNullOrEmpty(config.ChoiceOptionsColumn)
                ? FindColumnIndex(headerRecord, config.ChoiceOptionsColumn) : -1;
            var columnTypeColumn = !string.IsNullOrEmpty(config.ColumnTypeColumn)
                ? FindColumnIndex(headerRecord, config.ColumnTypeColumn) : -1;
            var includeColumn = !string.IsNullOrEmpty(config.IncludeColumn)
                ? FindColumnIndex(headerRecord, config.IncludeColumn) : -1;

            // Optional columns for full CREATE CSV
            var tableNameColumn = FindColumnIndex(headerRecord, config.TableNameColumn ?? "Table Name");
            var columnNameColumn = FindColumnIndex(headerRecord, config.ColumnNameColumn ?? "Column Name");
            var lookupTargetTableColumn = !string.IsNullOrEmpty(config.LookupTargetTableColumn)
                ? FindColumnIndex(headerRecord, config.LookupTargetTableColumn) : -1;
            var lookupRelationshipNameColumn = !string.IsNullOrEmpty(config.LookupRelationshipNameColumn)
                ? FindColumnIndex(headerRecord, config.LookupRelationshipNameColumn) : -1;
            var customerTargetTablesColumn = !string.IsNullOrEmpty(config.CustomerTargetTablesColumn)
                ? FindColumnIndex(headerRecord, config.CustomerTargetTablesColumn) : -1;
            var tableDisplayCollectionNameColumn = !string.IsNullOrEmpty(config.TableDisplayCollectionNameColumn)
                ? FindColumnIndex(headerRecord, config.TableDisplayCollectionNameColumn) : -1;
            var descriptionColumn = !string.IsNullOrEmpty(config.DescriptionColumn)
                ? FindColumnIndex(headerRecord, config.DescriptionColumn) : -1;
            var requiredColumn = !string.IsNullOrEmpty(config.RequiredColumn)
                ? FindColumnIndex(headerRecord, config.RequiredColumn) : -1;

            _logger.LogDebug("Found required columns - TableLogicalName: {TableLogicalCol}, LogicalName: {LogicalCol}",
                tableLogicalNameColumn, logicalNameColumn);

            while (await csv.ReadAsync())
            {
                cancellationToken.ThrowIfCancellationRequested();

                // Check include column filter (if configured)
                if (includeColumn != -1)
                {
                    var includeValue = csv.GetField(includeColumn)?.Trim();
                    if (!IsIncluded(includeValue))
                    {
                        continue;
                    }
                }

                // Required columns
                var tableLogicalName = csv.GetField(tableLogicalNameColumn)?.Trim();
                var logicalName = csv.GetField(logicalNameColumn)?.Trim();

                // Skip row if any required field is empty
                if (string.IsNullOrEmpty(tableLogicalName) || string.IsNullOrEmpty(logicalName))
                {
                    continue;
                }

                // Optional columns - passthrough/convenience
                var choiceOptions = choiceOptionsColumn != -1 ? csv.GetField(choiceOptionsColumn)?.Trim() : null;
                var columnType = columnTypeColumn != -1 ? csv.GetField(columnTypeColumn)?.Trim() : null;

                // Optional columns - CREATE CSV fields
                var tableName = tableNameColumn != -1 ? csv.GetField(tableNameColumn)?.Trim() : null;
                var columnName = columnNameColumn != -1 ? csv.GetField(columnNameColumn)?.Trim() : null;
                var lookupTargetTable = lookupTargetTableColumn != -1 ? csv.GetField(lookupTargetTableColumn)?.Trim() : null;
                var lookupRelationshipName = lookupRelationshipNameColumn != -1 ? csv.GetField(lookupRelationshipNameColumn)?.Trim() : null;
                var customerTargetTables = customerTargetTablesColumn != -1 ? csv.GetField(customerTargetTablesColumn)?.Trim() : null;
                var tableDisplayCollectionName = tableDisplayCollectionNameColumn != -1 ? csv.GetField(tableDisplayCollectionNameColumn)?.Trim() : null;
                var description = descriptionColumn != -1 ? csv.GetField(descriptionColumn)?.Trim() : null;
                var required = requiredColumn != -1 ? csv.GetField(requiredColumn)?.Trim() : null;

                schemas.Add(new SchemaDefinition
                {
                    TableLogicalName = tableLogicalName,
                    LogicalName = logicalName,
                    TableName = tableName,
                    ColumnName = columnName,
                    ColumnType = columnType,
                    ChoiceOptions = choiceOptions,
                    LookupTargetTable = lookupTargetTable,
                    LookupRelationshipName = lookupRelationshipName,
                    CustomerTargetTables = customerTargetTables,
                    TableDisplayCollectionName = tableDisplayCollectionName,
                    Description = description,
                    Required = required
                });
            }

            _logger.LogInformation("Successfully read {Count} schema definitions from CSV", schemas.Count);

            return schemas;
        }

        private int FindColumnIndex(string[] headers, string columnName)
        {
            for (int i = 0; i < headers.Length; i++)
            {
                if (headers[i]?.Trim().Equals(columnName, StringComparison.OrdinalIgnoreCase) == true)
                {
                    return i;
                }
            }

            _logger.LogWarning("Column '{ColumnName}' not found in CSV file", columnName);
            return -1;
        }

        private bool IsIncluded(string? includeValue)
        {
            if (string.IsNullOrWhiteSpace(includeValue))
            {
                return false;
            }

            var normalized = includeValue.Trim().ToLowerInvariant();
            return normalized == "yes" || normalized == "y" || normalized == "true" || normalized == "1";
        }
    }
}
