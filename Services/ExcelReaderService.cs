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

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                if (worksheet == null)
                {
                    throw new InvalidOperationException("No worksheets found in the Excel file.");
                }

                var tableNameColumn = FindColumnIndex(worksheet, config.TableNameColumn ?? "Table Name");
                var columnNameColumn = FindColumnIndex(worksheet, config.ColumnNameColumn ?? "Column Name");
                var columnTypeColumn = FindColumnIndex(worksheet, config.ColumnTypeColumn ?? "Column Type");

                if (tableNameColumn == -1 || columnNameColumn == -1 || columnTypeColumn == -1)
                {
                    throw new InvalidOperationException($"Required columns not found. Looking for: {config.TableNameColumn}, {config.ColumnNameColumn}, {config.ColumnTypeColumn}");
                }

                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    var tableName = worksheet.Cells[row, tableNameColumn].Text?.Trim();
                    var columnName = worksheet.Cells[row, columnNameColumn].Text?.Trim();
                    var columnType = worksheet.Cells[row, columnTypeColumn].Text?.Trim();

                    if (!string.IsNullOrEmpty(tableName) && !string.IsNullOrEmpty(columnName) && !string.IsNullOrEmpty(columnType))
                    {
                        schemas.Add(new SchemaDefinition
                        {
                            TableName = tableName,
                            ColumnName = columnName,
                            ColumnType = columnType
                        });
                    }
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