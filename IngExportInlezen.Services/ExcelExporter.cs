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

                worksheet.Cells[rowNumber, 1].Value = excelExport.Maand;
                worksheet.Cells[rowNumber, 2].Value = excelExport.Abonnementen;
                worksheet.Cells[rowNumber, 3].Value = excelExport.VasteLasten;
                worksheet.Cells[rowNumber, 4].Value = excelExport.Boodschappen;
                worksheet.Cells[rowNumber, 5].Value = excelExport.GeldOpnames;
                worksheet.Cells[rowNumber, 6].Value = excelExport.Tanken;
                worksheet.Cells[rowNumber, 7].Value = excelExport.InkomstenSalaris;
                worksheet.Cells[rowNumber, 8].Value = excelExport.OverigeInkomsten;
                worksheet.Cells[rowNumber, 9].Value = excelExport.SpaarOpdrachten;
                worksheet.Cells[rowNumber, 10].Value = excelExport.OverigeKosten;

                package.Save();
            }
        }
    }
}
