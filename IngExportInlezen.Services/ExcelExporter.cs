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
                ExcelWorksheet worksheetOverzicht = package.Workbook.Worksheets[0];
                ExcelWorksheet worksheetGrafieken = package.Workbook.Worksheets[1];

                var rowNumber = worksheetOverzicht.Dimension.Rows + 2;

                worksheetOverzicht.Cells[rowNumber, 2].Value = excelExport.Maand;
                worksheetOverzicht.Cells[rowNumber, 3].Value = excelExport.VasteLasten;
                worksheetOverzicht.Cells[rowNumber, 4].Value = excelExport.Abonnementen;
                worksheetOverzicht.Cells[rowNumber, 5].Value = excelExport.Boodschappen;
                worksheetOverzicht.Cells[rowNumber, 6].Value = excelExport.GeldOpnames;
                worksheetOverzicht.Cells[rowNumber, 7].Value = excelExport.Tanken;
                worksheetOverzicht.Cells[rowNumber, 8].Value = excelExport.OverigeKosten;
                worksheetOverzicht.Cells[rowNumber, 9].Value = excelExport.InkomstenSalaris;
                worksheetOverzicht.Cells[rowNumber, 10].Value = excelExport.OverigeInkomsten;
                worksheetOverzicht.Cells[rowNumber, 11].Value = excelExport.SpaarOpdrachten;

                //Maken van pie chart.
                var existingPieChart = worksheetOverzicht.Drawings["SpreidingKosten"];
                worksheetOverzicht.Drawings.Remove(existingPieChart);
                var pieChart = worksheetOverzicht.Drawings.AddChart("SpreidingKosten", eChartType.Pie);
                var serie = pieChart.Series.Add(worksheetOverzicht.Cells[rowNumber, 3, rowNumber, 8], worksheetOverzicht.Cells[2, 3, 2, 8]);
                pieChart.SetPosition(1, 0, 12, 0);
                pieChart.SetSize(700, 450);
                pieChart.Title.Text = $"Spreiding kosten {excelExport.Maand}";
                pieChart.Title.Font.Bold = true;
                pieChart.Legend.Position = eLegendPosition.Left;
                pieChart.Legend.Font.Size = 16;
                var pieSerie = (ExcelPieChartSerie)serie;
                pieSerie.DataLabel.ShowCategory = true;
                pieSerie.DataLabel.ShowPercent = true;
                pieSerie.DataLabel.Position = eLabelPosition.InEnd;

                //Maken van doughnut chart.
                var existingDoughnutChart = worksheetOverzicht.Drawings["MaandGoal"];
                worksheetOverzicht.Drawings.Remove(existingDoughnutChart);
                var doughnutChart = worksheetOverzicht.Drawings.AddChart("MaandGoal", eChartType.Doughnut);
                doughnutChart.SetPosition(24, 0, 12, 0);
                doughnutChart.SetSize(700, 450);

                //Nieuwe regel opmaak
                var range = worksheetOverzicht.Cells[$"B{rowNumber}:K{rowNumber}"];
                var rangeMin = worksheetOverzicht.Cells[$"B{rowNumber-1}:K{rowNumber-1}"];
                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 220, 220));
                var maandCell = worksheetOverzicht.Cells[$"B{rowNumber}"];
                var laatsteCell = worksheetOverzicht.Cells[$"K{rowNumber}"];
                range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                maandCell.Style.Font.Bold = true; 
                maandCell.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                laatsteCell.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                rangeMin.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;

                //Diagram Overzicht kosten
                //var diagramKosten = worksheetOverzicht.Drawings["Chart 13"] as ExcelChart;
                //var serieKosten = diagramKosten.Series.Add(worksheetOverzicht.Cells[rowNumber, 2, rowNumber, 7], worksheetOverzicht.Cells[rowNumber, 1]);
                //serieKosten.HeaderAddress = worksheetOverzicht.Cells[rowNumber, 1];
                //diagramKosten.SetPosition(20, 0, 11, 0);

                ////Diagram Overzicht inkomsten
                //var diagramInkomsten = worksheetOverzicht.Drawings["Chart 14"] as ExcelChart;
                //var serieInkomsten = diagramInkomsten.Series.Add(worksheetOverzicht.Cells[rowNumber, 8, rowNumber, 10], worksheetOverzicht.Cells[rowNumber, 1]);
                //serieInkomsten.HeaderAddress = worksheetOverzicht.Cells[rowNumber, 1];
                //diagramInkomsten.SetPosition(44, 0, 11, 0);

                //Diagrammen grafieken worksheet


                //foreach (var drawing in worksheetGrafieken.Drawings)
                //{
                //    if (drawing is ExcelChart)
                //    {
                //        // Handle chart-specific logic
                //        var chart = (ExcelChart)drawing;
                //        Console.WriteLine($"Chart Name: {chart.Name}");
                //        // Add more chart-specific logic as needed
                //    }
                //    else if (drawing is ExcelPicture)
                //    {
                //        // Handle picture-specific logic
                //        var picture = (ExcelPicture)drawing;
                //        Console.WriteLine($"Picture Name: {picture.Name}");
                //        // Add more picture-specific logic as needed
                //    }
                //    // Add more conditions for other types of drawings (e.g., shapes)

                //    // Common properties for all drawings
                //    Console.WriteLine($"Drawing Type: {drawing.GetType().Name}");
                //    Console.WriteLine($"Description: {drawing.Description}");
                //    Console.WriteLine($"Position: {drawing.From.Column}, {drawing.From.Row}");
                //    Console.WriteLine($"Size: {drawing.To.Column - drawing.From.Column}, {drawing.To.Row - drawing.From.Row}");
                //    Console.WriteLine();
                //}

                package.Save();
            }
        }
    }
}
