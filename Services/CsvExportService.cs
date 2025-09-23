using CsvHelper;
using CsvHelper.Configuration;
using DataverseSchemaManager.Models;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace DataverseSchemaManager.Services
{
    public class CsvExportService
    {
        public void ExportNewSchema(List<SchemaDefinition> schemas, string outputPath)
        {
            var newSchemas = schemas.Where(s => !s.ExistsInDataverse).ToList();

            using (var writer = new StreamWriter(outputPath))
            using (var csv = new CsvWriter(writer, new CsvConfiguration(CultureInfo.InvariantCulture)))
            {
                csv.WriteRecords(newSchemas);
            }
        }

        public void ExportAllSchema(List<SchemaDefinition> schemas, string outputPath)
        {
            using (var writer = new StreamWriter(outputPath))
            using (var csv = new CsvWriter(writer, new CsvConfiguration(CultureInfo.InvariantCulture)))
            {
                csv.WriteRecords(schemas);
            }
        }
    }
}