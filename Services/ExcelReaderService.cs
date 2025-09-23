using DataverseSchemaManager.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace DataverseSchemaManager.Services
{
    public class ExcelReaderService
    {
        static ExcelReaderService()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public ExcelReaderService()
        {
        }

        public List<SchemaDefinition> ReadSchemaDefinitions(string filePath, AppConfiguration config)
        {
            var schemas = new List<SchemaDefinition>();

            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"Excel file not found: {filePath}");
            }

            string? tempFilePath = null;
            string fileToRead = filePath;

            try
            {
                using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var package = new ExcelPackage(stream);

                return ReadExcelData(package, config);
            }
            catch (IOException ioEx)
            {
                Console.WriteLine($"File is locked, attempting to create temporary copy with retries...");

                int maxRetries = 5;
                int retryDelayMs = 500;
                Exception? lastException = ioEx;

                for (int attempt = 1; attempt <= maxRetries; attempt++)
                {
                    try
                    {
                        tempFilePath = Path.Combine(Path.GetTempPath(), $"schema_temp_{Guid.NewGuid()}.xlsx");

                        Console.WriteLine($"Attempt {attempt} of {maxRetries}...");

                        File.Copy(filePath, tempFilePath, overwrite: true);

                        Console.WriteLine($"Temporary copy created successfully. Reading data...");

                        using var tempStream = new FileStream(tempFilePath, FileMode.Open, FileAccess.Read, FileShare.Read);
                        using var package = new ExcelPackage(tempStream);

                        var result = ReadExcelData(package, config);

                        if (tempFilePath != null && File.Exists(tempFilePath))
                        {
                            try { File.Delete(tempFilePath); Console.WriteLine("Temporary file cleaned up."); } catch { }
                        }

                        return result;
                    }
                    catch (IOException copyEx)
                    {
                        lastException = copyEx;

                        if (attempt < maxRetries)
                        {
                            Console.WriteLine($"  Retry {attempt} failed, waiting {retryDelayMs}ms...");
                            System.Threading.Thread.Sleep(retryDelayMs);
                            retryDelayMs *= 2;
                        }
                    }
                }

                throw new IOException($"Cannot access Excel file after {maxRetries} attempts. The file may be exclusively locked by Excel, OneDrive, or another process. Please close the file and try again. Last error: {lastException?.Message}", lastException);
            }
        }

        private List<SchemaDefinition> ReadExcelData(ExcelPackage package, AppConfiguration config)
        {
            var schemas = new List<SchemaDefinition>();

            var worksheet = package.Workbook.Worksheets.FirstOrDefault();
            if (worksheet == null)
            {
                throw new InvalidOperationException("No worksheets found in the Excel file.");
            }

            var tableNameColumn = FindColumnIndex(worksheet, config.TableNameColumn ?? "Table Name");
            var columnNameColumn = FindColumnIndex(worksheet, config.ColumnNameColumn ?? "Column Name");
            var columnTypeColumn = FindColumnIndex(worksheet, config.ColumnTypeColumn ?? "Column Type");
            var choiceOptionsColumn = FindColumnIndex(worksheet, config.ChoiceOptionsColumn ?? "Choice Options");

            if (tableNameColumn == -1 || columnNameColumn == -1 || columnTypeColumn == -1)
            {
                throw new InvalidOperationException($"Required columns not found. Looking for: {config.TableNameColumn}, {config.ColumnNameColumn}, {config.ColumnTypeColumn}");
            }

            for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
            {
                var tableName = worksheet.Cells[row, tableNameColumn].Text?.Trim();
                var columnName = worksheet.Cells[row, columnNameColumn].Text?.Trim();
                var columnType = worksheet.Cells[row, columnTypeColumn].Text?.Trim();
                var choiceOptions = choiceOptionsColumn != -1 ? worksheet.Cells[row, choiceOptionsColumn].Text?.Trim() : null;

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

            return schemas;
        }

        private int FindColumnIndex(ExcelWorksheet worksheet, string columnName)
        {
            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                if (worksheet.Cells[1, col].Text?.Trim().Equals(columnName, StringComparison.OrdinalIgnoreCase) == true)
                {
                    return col;
                }
            }
            return -1;
        }
    }
}