using IngExportInlezen.Domain;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

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

                var rowNumber = worksheet.Dimension.Rows + 1;

                worksheet.Cells[rowNumber, 1].Value = excelExport.Maand;
                worksheet.Cells[rowNumber, 2].Value = excelExport.VasteLasten;
                worksheet.Cells[rowNumber, 3].Value = excelExport.Abonnementen;
                worksheet.Cells[rowNumber, 4].Value = excelExport.Boodschappen;
                worksheet.Cells[rowNumber, 5].Value = excelExport.GeldOpnames;
                worksheet.Cells[rowNumber, 6].Value = excelExport.Tanken;
                worksheet.Cells[rowNumber, 7].Value = excelExport.OverigeKosten;
                worksheet.Cells[rowNumber, 8].Value = excelExport.InkomstenSalaris;
                worksheet.Cells[rowNumber, 9].Value = excelExport.OverigeInkomsten;
                worksheet.Cells[rowNumber, 10].Value = excelExport.SpaarOpdrachten;

                //Maken van pie chart.
                var existingChart = worksheet.Drawings["SpreidingKosten"];
                worksheet.Drawings.Remove(existingChart);
                var pieChart = worksheet.Drawings.AddChart("SpreidingKosten", eChartType.Pie);
                var serie = pieChart.Series.Add(worksheet.Cells[rowNumber, 2, rowNumber, 7], worksheet.Cells[1, 2, 1, 7]);
                pieChart.SetPosition(0, 0, 11, 0);
                pieChart.SetSize(600, 400);
                pieChart.Title.Text = "Spreiding kosten laatste maand";
                pieChart.Title.Font.Bold = true;
                pieChart.Legend.Position = eLegendPosition.Left;
                pieChart.Legend.Font.Size = 12;
                var pieSerie = (ExcelPieChartSerie)serie;
                pieSerie.DataLabel.ShowCategory = true;
                pieSerie.DataLabel.ShowPercent = true;

                //Diagram Overzicht kosten
                var diagramKosten = worksheet.Drawings["Chart 13"] as ExcelChart;
                var serieKosten = diagramKosten.Series.Add(worksheet.Cells[rowNumber, 2, rowNumber, 7], worksheet.Cells[rowNumber, 1]);
                serieKosten.HeaderAddress = worksheet.Cells[rowNumber, 1];
                diagramKosten.SetPosition(20, 0, 11, 0);

                //Diagram Overzicht inkomsten
                var diagramInkomsten = worksheet.Drawings["Chart 14"] as ExcelChart;
                var serieInkomsten = diagramInkomsten.Series.Add(worksheet.Cells[rowNumber, 8, rowNumber, 10], worksheet.Cells[rowNumber, 1]);
                serieInkomsten.HeaderAddress = worksheet.Cells[rowNumber, 1];
                diagramInkomsten.SetPosition(44, 0, 11, 0);

                package.Save();
            }
        }
    }
}
