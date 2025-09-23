using OfficeOpenXml;
using System;
using System.IO;

namespace DataverseSchemaManager
{
    public class GenerateSampleExcel
    {
        static GenerateSampleExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public static void CreateSampleFile(string filePath = "sample_schema.xlsx")
        {

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Schema Definitions");

                worksheet.Cells[1, 1].Value = "Table Name";
                worksheet.Cells[1, 2].Value = "Column Name";
                worksheet.Cells[1, 3].Value = "Column Type";

                var sampleData = new[,]
                {
                    { "account", "new_customfield1", "text" },
                    { "account", "new_revenue", "decimal" },
                    { "contact", "new_birthdate", "date" },
                    { "contact", "new_isactive", "boolean" },
                    { "lead", "new_score", "number" },
                    { "opportunity", "new_description", "text" },
                    { "opportunity", "new_closedate", "datetime" }
                };

                for (int i = 0; i < sampleData.GetLength(0); i++)
                {
                    for (int j = 0; j < sampleData.GetLength(1); j++)
                    {
                        worksheet.Cells[i + 2, j + 1].Value = sampleData[i, j];
                    }
                }

                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                worksheet.Row(1).Style.Font.Bold = true;
                worksheet.Row(1).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Row(1).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);

                package.SaveAs(new FileInfo(filePath));
            }

            Console.WriteLine($"Sample Excel file created: {filePath}");
        }
    }
}