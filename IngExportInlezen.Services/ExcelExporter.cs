using IngExportInlezen.Domain;
using OfficeOpenXml;
using System;
using System.IO;

namespace IngExportInlezen.Services
{
    public static class ExcelExporter
    {
        public static void ExportToExcel(ExcelExport excelExport)
        {
            string filePath = @"C:\Users\coenj\Documents\Financieel overzicht\ING export\Financieel Overzicht.xlsx";

            if (!File.Exists(filePath))
            {
                Console.WriteLine("Excel bestand niet gevonden.");
                return;
            }

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            FileInfo existingFile = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                int rowNumber = worksheet.Dimension.Rows + 1;
                int columnNumber = 5;

                worksheet.Cells[rowNumber, columnNumber].Value = "New Data 1";
                package.Save();
            }

                throw new NotImplementedException();
        }
    }
}
