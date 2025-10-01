using DataverseSchemaManager.Models;
using DataverseSchemaManager.Services;
using Microsoft.Extensions.Logging;
using Moq;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace DataverseSchemaManager.Tests.Services
{
    public class CsvExportServiceTests
    {
        private readonly Mock<ILogger<CsvExportService>> _mockLogger;
        private readonly CsvExportService _service;

        public CsvExportServiceTests()
        {
            _mockLogger = new Mock<ILogger<CsvExportService>>();
            _service = new CsvExportService(_mockLogger.Object);
        }

        [Fact]
        public async Task ExportNewSchemaAsync_ShouldOnlyExportNewSchemas()
        {
            // Arrange
            var outputPath = Path.GetTempFileName();
            var schemas = new List<SchemaDefinition>
            {
                new SchemaDefinition
                {
                    TableName = "account",
                    ColumnName = "existing_column",
                    ColumnType = "text",
                    ColumnExistsInDataverse = true
                },
                new SchemaDefinition
                {
                    TableName = "contact",
                    ColumnName = "new_column",
                    ColumnType = "number",
                    ColumnExistsInDataverse = false
                }
            };

            try
            {
                // Act
                await _service.ExportNewSchemaAsync(schemas, outputPath, CancellationToken.None);

                // Assert
                Assert.True(File.Exists(outputPath));
                var content = await File.ReadAllTextAsync(outputPath);
                Assert.Contains("new_column", content);
                Assert.DoesNotContain("existing_column", content);
            }
            finally
            {
                if (File.Exists(outputPath))
                    File.Delete(outputPath);
            }
        }

        [Fact]
        public async Task ExportAllSchemaAsync_ShouldExportAllSchemas()
        {
            // Arrange
            var outputPath = Path.GetTempFileName();
            var schemas = new List<SchemaDefinition>
            {
                new SchemaDefinition
                {
                    TableName = "account",
                    ColumnName = "column1",
                    ColumnType = "text",
                    ColumnExistsInDataverse = true
                },
                new SchemaDefinition
                {
                    TableName = "contact",
                    ColumnName = "column2",
                    ColumnType = "number",
                    ColumnExistsInDataverse = false
                }
            };

            try
            {
                // Act
                await _service.ExportAllSchemaAsync(schemas, outputPath, CancellationToken.None);

                // Assert
                Assert.True(File.Exists(outputPath));
                var content = await File.ReadAllTextAsync(outputPath);
                Assert.Contains("column1", content);
                Assert.Contains("column2", content);
            }
            finally
            {
                if (File.Exists(outputPath))
                    File.Delete(outputPath);
            }
        }
    }
}
