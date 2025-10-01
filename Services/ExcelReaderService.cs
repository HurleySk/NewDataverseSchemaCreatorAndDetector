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

                var tableNameColumn = FindColumnIndex(worksheet, config.TableNameColumn ?? "Table Name");
                var columnNameColumn = FindColumnIndex(worksheet, config.ColumnNameColumn ?? "Column Name");
                var columnTypeColumn = FindColumnIndex(worksheet, config.ColumnTypeColumn ?? "Column Type");
                var choiceOptionsColumn = FindColumnIndex(worksheet, config.ChoiceOptionsColumn ?? "Choice Options");

                if (tableNameColumn == -1 || columnNameColumn == -1 || columnTypeColumn == -1)
                {
                    _logger.LogError(
                        "Required columns not found. Looking for: TableName='{TableName}', ColumnName='{ColumnName}', ColumnType='{ColumnType}'",
                        config.TableNameColumn,
                        config.ColumnNameColumn,
                        config.ColumnTypeColumn);

                    throw new InvalidOperationException(
                        $"Required columns not found. Looking for: {config.TableNameColumn}, " +
                        $"{config.ColumnNameColumn}, {config.ColumnTypeColumn}");
                }

                _logger.LogDebug(
                    "Found columns - Table: {TableCol}, Column: {ColumnCol}, Type: {TypeCol}, Choice: {ChoiceCol}",
                    tableNameColumn,
                    columnNameColumn,
                    columnTypeColumn,
                    choiceOptionsColumn != -1 ? choiceOptionsColumn.ToString() : "Not found");

                if (worksheet.Dimension == null)
                {
                    _logger.LogWarning("Worksheet has no dimension (empty worksheet)");
                    return schemas;
                }

                for (int row = DataverseConstants.Excel.DefaultDataStartRow; row <= worksheet.Dimension.End.Row; row++)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    var tableName = worksheet.Cells[row, tableNameColumn].Text?.Trim();
                    var columnName = worksheet.Cells[row, columnNameColumn].Text?.Trim();
                    var columnType = worksheet.Cells[row, columnTypeColumn].Text?.Trim();
                    var choiceOptions = choiceOptionsColumn != -1
                        ? worksheet.Cells[row, choiceOptionsColumn].Text?.Trim()
                        : null;

                    if (!string.IsNullOrEmpty(tableName) && !string.IsNullOrEmpty(columnName) && !string.IsNullOrEmpty(columnType))
                    {
                        schemas.Add(new SchemaDefinition
                        {
                            TableName = tableName,
                            ColumnName = columnName,
                            ColumnType = columnType,
                            ChoiceOptions = choiceOptions
                        });
                    }
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
    }
}
