using DataverseSchemaManager.Constants;
using DataverseSchemaManager.Interfaces;
using DataverseSchemaManager.Models;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace DataverseSchemaManager.Services
{
    /// <summary>
    /// Provides operations for reading schema definitions from Excel files.
    /// </summary>
    public class ExcelReaderService : IExcelReaderService
    {
        private readonly ILogger<ExcelReaderService> _logger;

        static ExcelReaderService()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public ExcelReaderService(ILogger<ExcelReaderService> logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <inheritdoc/>
        public async Task<List<SchemaDefinition>> ReadSchemaDefinitionsAsync(string filePath, AppConfiguration config, CancellationToken cancellationToken = default)
        {
            if (!File.Exists(filePath))
            {
                _logger.LogError("Excel file not found: {FilePath}", filePath);
                throw new FileNotFoundException($"Excel file not found: {filePath}");
            }

            _logger.LogInformation("Reading schema definitions from Excel file: {FilePath}", filePath);

            string? tempFilePath = null;

            try
            {
                using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var package = new ExcelPackage(stream);

                return await ReadExcelDataAsync(package, config, cancellationToken);
            }
            catch (IOException ioEx)
            {
                _logger.LogWarning("File is locked, attempting to create temporary copy with retries...");

                return await RetryReadLockedFileAsync(filePath, config, ioEx, cancellationToken);
            }
            finally
            {
                if (tempFilePath != null && File.Exists(tempFilePath))
                {
                    try
                    {
                        File.Delete(tempFilePath);
                        _logger.LogDebug("Temporary file cleaned up: {TempFile}", tempFilePath);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, "Failed to clean up temporary file: {TempFile}", tempFilePath);
                    }
                }
            }
        }

        private async Task<List<SchemaDefinition>> RetryReadLockedFileAsync(
            string filePath,
            AppConfiguration config,
            IOException originalException,
            CancellationToken cancellationToken)
        {
            int retryDelayMs = DataverseConstants.Excel.InitialRetryDelayMs;
            Exception? lastException = originalException;

            for (int attempt = 1; attempt <= DataverseConstants.Excel.MaxFileAccessRetries; attempt++)
            {
                cancellationToken.ThrowIfCancellationRequested();

                try
                {
                    var tempFilePath = Path.Combine(Path.GetTempPath(), $"schema_temp_{Guid.NewGuid()}.xlsx");

                    _logger.LogDebug("Retry attempt {Attempt} of {MaxAttempts}...", attempt, DataverseConstants.Excel.MaxFileAccessRetries);

                    await Task.Run(() => File.Copy(filePath, tempFilePath, overwrite: true), cancellationToken);

                    _logger.LogInformation("Temporary copy created successfully. Reading data...");

                    using var tempStream = new FileStream(tempFilePath, FileMode.Open, FileAccess.Read, FileShare.Read);
                    using var package = new ExcelPackage(tempStream);

                    var result = await ReadExcelDataAsync(package, config, cancellationToken);

                    // Clean up temp file
                    try
                    {
                        File.Delete(tempFilePath);
                        _logger.LogDebug("Temporary file cleaned up.");
                    }
                    catch { }

                    return result;
                }
                catch (IOException copyEx)
                {
                    lastException = copyEx;

                    if (attempt < DataverseConstants.Excel.MaxFileAccessRetries)
                    {
                        _logger.LogWarning(
                            "Retry {Attempt} failed, waiting {DelayMs}ms before next retry...",
                            attempt,
                            retryDelayMs);

                        await Task.Delay(retryDelayMs, cancellationToken);
                        retryDelayMs *= 2; // Exponential backoff
                    }
                }
            }

            _logger.LogError(
                lastException,
                "Cannot access Excel file after {MaxRetries} attempts. The file may be exclusively locked.",
                DataverseConstants.Excel.MaxFileAccessRetries);

            throw new IOException(
                $"Cannot access Excel file after {DataverseConstants.Excel.MaxFileAccessRetries} attempts. " +
                $"The file may be exclusively locked by Excel, OneDrive, or another process. " +
                $"Please close the file and try again. Last error: {lastException?.Message}",
                lastException);
        }

        private async Task<List<SchemaDefinition>> ReadExcelDataAsync(
            ExcelPackage package,
            AppConfiguration config,
            CancellationToken cancellationToken)
        {
            return await Task.Run(() =>
            {
                var schemas = new List<SchemaDefinition>();

                var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                if (worksheet == null)
                {
                    _logger.LogError("No worksheets found in the Excel file");
                    throw new InvalidOperationException("No worksheets found in the Excel file.");
                }

                _logger.LogDebug(
                    "Reading from worksheet '{WorksheetName}' with {RowCount} rows and {ColumnCount} columns",
                    worksheet.Name,
                    worksheet.Dimension?.End.Row ?? 0,
                    worksheet.Dimension?.End.Column ?? 0);

                // Required columns - Logical names (for detection)
                var logicalNameColumn = FindColumnIndex(worksheet, config.LogicalNameColumn ?? "Logical Name");
                var tableLogicalNameColumn = FindColumnIndex(worksheet, config.TableLogicalNameColumn ?? "Table Logical Name");

                // Validate required columns exist
                if (logicalNameColumn == -1 || tableLogicalNameColumn == -1)
                {
                    var missing = new List<string>();
                    if (logicalNameColumn == -1) missing.Add(config.LogicalNameColumn ?? "Logical Name");
                    if (tableLogicalNameColumn == -1) missing.Add(config.TableLogicalNameColumn ?? "Table Logical Name");

                    _logger.LogError("Required columns not found: {MissingColumns}", string.Join(", ", missing));

                    throw new InvalidOperationException(
                        $"Required columns not found in Excel: {string.Join(", ", missing)}. " +
                        $"These columns are REQUIRED for schema detection.");
                }

                // Optional columns - Display names (for creation)
                var tableNameColumn = !string.IsNullOrEmpty(config.TableNameColumn)
                    ? FindColumnIndex(worksheet, config.TableNameColumn) : -1;
                var columnNameColumn = !string.IsNullOrEmpty(config.ColumnNameColumn)
                    ? FindColumnIndex(worksheet, config.ColumnNameColumn) : -1;
                var columnTypeColumn = !string.IsNullOrEmpty(config.ColumnTypeColumn)
                    ? FindColumnIndex(worksheet, config.ColumnTypeColumn) : -1;

                // Optional columns - Choice fields
                var choiceOptionsColumn = FindColumnIndex(worksheet, config.ChoiceOptionsColumn ?? "Choice Options");

                // Optional columns - Lookup/Customer fields
                var lookupTargetTableColumn = !string.IsNullOrEmpty(config.LookupTargetTableColumn)
                    ? FindColumnIndex(worksheet, config.LookupTargetTableColumn) : -1;
                var lookupRelationshipNameColumn = !string.IsNullOrEmpty(config.LookupRelationshipNameColumn)
                    ? FindColumnIndex(worksheet, config.LookupRelationshipNameColumn) : -1;
                var customerTargetTablesColumn = !string.IsNullOrEmpty(config.CustomerTargetTablesColumn)
                    ? FindColumnIndex(worksheet, config.CustomerTargetTablesColumn) : -1;

                // Optional columns - Metadata
                var tableDisplayCollectionNameColumn = !string.IsNullOrEmpty(config.TableDisplayCollectionNameColumn)
                    ? FindColumnIndex(worksheet, config.TableDisplayCollectionNameColumn) : -1;
                var descriptionColumn = !string.IsNullOrEmpty(config.DescriptionColumn)
                    ? FindColumnIndex(worksheet, config.DescriptionColumn) : -1;
                var requiredColumn = !string.IsNullOrEmpty(config.RequiredColumn)
                    ? FindColumnIndex(worksheet, config.RequiredColumn) : -1;

                // Optional columns - Filtering
                var includeColumn = !string.IsNullOrEmpty(config.IncludeColumn)
                    ? FindColumnIndex(worksheet, config.IncludeColumn) : -1;

                _logger.LogDebug(
                    "Found required columns - LogicalName: {LogicalCol}, TableLogicalName: {TableLogicalCol}",
                    logicalNameColumn,
                    tableLogicalNameColumn);

                if (choiceOptionsColumn != -1 || lookupTargetTableColumn != -1 || customerTargetTablesColumn != -1 ||
                    tableDisplayCollectionNameColumn != -1 || descriptionColumn != -1 || requiredColumn != -1 || includeColumn != -1)
                {
                    _logger.LogDebug(
                        "Found optional columns - Choice: {ChoiceCol}, LookupTarget: {LookupCol}, CustomerTargets: {CustomerCol}, " +
                        "TableDisplayCollection: {TableDisplayCol}, Description: {DescCol}, Required: {ReqCol}, Include: {IncludeCol}",
                        choiceOptionsColumn != -1 ? choiceOptionsColumn.ToString() : "Not found",
                        lookupTargetTableColumn != -1 ? lookupTargetTableColumn.ToString() : "Not found",
                        customerTargetTablesColumn != -1 ? customerTargetTablesColumn.ToString() : "Not found",
                        tableDisplayCollectionNameColumn != -1 ? tableDisplayCollectionNameColumn.ToString() : "Not found",
                        descriptionColumn != -1 ? descriptionColumn.ToString() : "Not found",
                        requiredColumn != -1 ? requiredColumn.ToString() : "Not found",
                        includeColumn != -1 ? includeColumn.ToString() : "Not found");
                }

                if (worksheet.Dimension == null)
                {
                    _logger.LogWarning("Worksheet has no dimension (empty worksheet)");
                    return schemas;
                }

                for (int row = DataverseConstants.Excel.DefaultDataStartRow; row <= worksheet.Dimension.End.Row; row++)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    // Check include column filter (if configured)
                    if (includeColumn != -1)
                    {
                        var includeValue = worksheet.Cells[row, includeColumn].Text?.Trim();
                        if (!IsIncluded(includeValue))
                        {
                            continue;
                        }
                    }

                    // Required columns - Logical names
                    var logicalName = worksheet.Cells[row, logicalNameColumn].Text?.Trim();
                    var tableLogicalName = worksheet.Cells[row, tableLogicalNameColumn].Text?.Trim();

                    // Skip row if any required field is empty
                    if (string.IsNullOrEmpty(logicalName) || string.IsNullOrEmpty(tableLogicalName))
                    {
                        continue;
                    }

                    // Optional columns - Display names (for creation)
                    var tableName = tableNameColumn != -1 ? worksheet.Cells[row, tableNameColumn].Text?.Trim() : null;
                    var columnName = columnNameColumn != -1 ? worksheet.Cells[row, columnNameColumn].Text?.Trim() : null;
                    var columnType = columnTypeColumn != -1 ? worksheet.Cells[row, columnTypeColumn].Text?.Trim() : null;

                    // Optional columns - Choice fields
                    var choiceOptions = choiceOptionsColumn != -1
                        ? worksheet.Cells[row, choiceOptionsColumn].Text?.Trim()
                        : null;

                    // Optional columns - Lookup/Customer fields
                    var lookupTargetTable = lookupTargetTableColumn != -1
                        ? worksheet.Cells[row, lookupTargetTableColumn].Text?.Trim()
                        : null;
                    var lookupRelationshipName = lookupRelationshipNameColumn != -1
                        ? worksheet.Cells[row, lookupRelationshipNameColumn].Text?.Trim()
                        : null;
                    var customerTargetTables = customerTargetTablesColumn != -1
                        ? worksheet.Cells[row, customerTargetTablesColumn].Text?.Trim()
                        : null;

                    // Optional columns - Metadata
                    var tableDisplayCollectionName = tableDisplayCollectionNameColumn != -1
                        ? worksheet.Cells[row, tableDisplayCollectionNameColumn].Text?.Trim()
                        : null;
                    var description = descriptionColumn != -1
                        ? worksheet.Cells[row, descriptionColumn].Text?.Trim()
                        : null;
                    var required = requiredColumn != -1
                        ? worksheet.Cells[row, requiredColumn].Text?.Trim()
                        : null;

                    schemas.Add(new SchemaDefinition
                    {
                        TableName = tableName,
                        ColumnName = columnName,
                        ColumnType = columnType,
                        LogicalName = logicalName,
                        TableLogicalName = tableLogicalName,
                        ChoiceOptions = choiceOptions,
                        LookupTargetTable = lookupTargetTable,
                        LookupRelationshipName = lookupRelationshipName,
                        CustomerTargetTables = customerTargetTables,
                        TableDisplayCollectionName = tableDisplayCollectionName,
                        Description = description,
                        Required = required
                    });
                }

                _logger.LogInformation("Successfully read {Count} schema definitions from Excel", schemas.Count);

                return schemas;
            }, cancellationToken);
        }

        private int FindColumnIndex(ExcelWorksheet worksheet, string columnName)
        {
            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                if (worksheet.Cells[DataverseConstants.Excel.DefaultHeaderRow, col].Text?
                    .Trim()
                    .Equals(columnName, StringComparison.OrdinalIgnoreCase) == true)
                {
                    return col;
                }
            }

            _logger.LogWarning("Column '{ColumnName}' not found in Excel file", columnName);
            return -1;
        }

        private bool IsIncluded(string? includeValue)
        {
            // Null/empty means include by default
            if (string.IsNullOrWhiteSpace(includeValue))
            {
                return true;
            }

            var normalized = includeValue.Trim().ToLowerInvariant();

            // Yes values mean include
            if (normalized == "yes" || normalized == "y" || normalized == "true" || normalized == "1")
            {
                return true;
            }

            // No values mean exclude
            if (normalized == "no" || normalized == "n" || normalized == "false" || normalized == "0")
            {
                return false;
            }

            // Default to including if value is unrecognized
            return true;
        }
    }
}
