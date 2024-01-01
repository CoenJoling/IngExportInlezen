using IngExportInlezen.Domain;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;

namespace IngExportInlezen.Services
{
    public static class ExcelExporter
    {
        static ExcelExporter()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public static void ExportToJaaroverzichtExcel(ExcelExport excelExport)
        {
            string filePath = @"C:\Users\coenj\Documents\Financieel overzicht\ING export\Jaarlijks Financieel Overzicht.xlsx";

            if (!File.Exists(filePath))
            {
                Console.WriteLine("Excel bestand niet gevonden.");
                return;
            }

            FileInfo existingFile = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                ExcelWorksheet abonnementen = package.Workbook.Worksheets[0];
                ExcelWorksheet vasteLasten = package.Workbook.Worksheets[1];
                ExcelWorksheet boodschappen = package.Workbook.Worksheets[2];
                ExcelWorksheet geldOpnames = package.Workbook.Worksheets[3];
                ExcelWorksheet tanken = package.Workbook.Worksheets[4];
                ExcelWorksheet inkomstenSalaris = package.Workbook.Worksheets[5];
                ExcelWorksheet overigeInkomsten = package.Workbook.Worksheets[6];
                ExcelWorksheet spaarOpdrachten = package.Workbook.Worksheets[7];
                ExcelWorksheet overigeKosten = package.Workbook.Worksheets[8];

                package.Save();
            }
        }

        public static void ExportToMaandExcel(ExcelExport excelExport)
        {
            string filePath = @"C:\Users\coenj\Documents\Financieel overzicht\ING export\Maandelijks Financieel Overzicht.xlsx";

            if (!File.Exists(filePath))
            {
                Console.WriteLine("Excel bestand niet gevonden.");
                return;
            }

            FileInfo existingFile = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                ExcelWorksheet worksheetOverzicht = package.Workbook.Worksheets[0];
                ExcelWorksheet worksheetGrafieken = package.Workbook.Worksheets[2];

                var rowNumber = worksheetOverzicht.Dimension.Rows + 2;

                worksheetOverzicht.Cells[rowNumber, 2].Value = excelExport.Maand;
                worksheetOverzicht.Cells[rowNumber, 3].Value = excelExport.VasteLasten.Sum(x => x.Bedrag);
                worksheetOverzicht.Cells[rowNumber, 4].Value = excelExport.Abonnementen.Sum(x => x.Bedrag);
                worksheetOverzicht.Cells[rowNumber, 5].Value = excelExport.Boodschappen.Sum(x => x.Bedrag);
                worksheetOverzicht.Cells[rowNumber, 6].Value = excelExport.GeldOpnames.Sum(x => x.Bedrag);
                worksheetOverzicht.Cells[rowNumber, 7].Value = excelExport.Tanken.Sum(x => x.Bedrag);
                worksheetOverzicht.Cells[rowNumber, 8].Value = excelExport.OverigeKosten.Sum(x => x.Bedrag) * -1;
                worksheetOverzicht.Cells[rowNumber, 9].Value = excelExport.InkomstenSalaris.Sum(x => x.Bedrag);
                worksheetOverzicht.Cells[rowNumber, 10].Value = excelExport.OverigeInkomsten.Sum(x => x.Bedrag);
                worksheetOverzicht.Cells[rowNumber, 11].Value = excelExport.SpaarOpdrachtenIngelegd.Sum(x => x.Bedrag) - excelExport.SpaarOpdrachtenOpgenomen.Sum(x => x.Bedrag);

                //Nieuwe regel opmaak
                var range = worksheetOverzicht.Cells[$"B{rowNumber}:K{rowNumber}"];
                var rangeMin = worksheetOverzicht.Cells[$"B{rowNumber - 1}:K{rowNumber - 1}"];
                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 220, 220));
                var maandCell = worksheetOverzicht.Cells[$"B{rowNumber}"];
                var laatsteCell = worksheetOverzicht.Cells[$"K{rowNumber}"];
                maandCell.Style.Font.Bold = true;

                //Maken van pie chart.
                var existingPieChart = worksheetOverzicht.Drawings["SpreidingKosten"];
                worksheetOverzicht.Drawings.Remove(existingPieChart);
                var pieChart = worksheetOverzicht.Drawings.AddChart("SpreidingKosten", eChartType.Pie);
                var serie = pieChart.Series.Add(worksheetOverzicht.Cells[rowNumber, 3, rowNumber, 8], worksheetOverzicht.Cells[2, 3, 2, 8]);
                pieChart.SetPosition(1, 0, 12, 0);
                pieChart.SetSize(700, 450);
                pieChart.Title.Text = $"Spreiding kosten {excelExport.Maand}";
                pieChart.Title.Font.Bold = true;
                pieChart.Title.Font.Size = 16;
                pieChart.Legend.Position = eLegendPosition.Left;
                pieChart.Legend.Font.Size = 16;
                var pieSerie = (ExcelPieChartSerie)serie;
                pieSerie.DataLabel.ShowCategory = true;
                pieSerie.DataLabel.ShowPercent = true;
                pieSerie.DataLabel.ShowLeaderLines = true;
                pieSerie.DataLabel.Position = eLabelPosition.OutEnd;
                pieChart.Fill.Style = eFillStyle.SolidFill;
                pieChart.Fill.Color = System.Drawing.Color.FromArgb(255, 247, 247);
                

                //Maken van doughnut chart.
                var existingDoughnutChart = worksheetOverzicht.Drawings["MaandGoal"];
                worksheetOverzicht.Drawings.Remove(existingDoughnutChart);
                var doughnutChart = worksheetOverzicht.Drawings.AddChart("MaandGoal", eChartType.Doughnut);
                doughnutChart.SetPosition(24, 0, 12, 0);
                doughnutChart.SetSize(700, 450);
                doughnutChart.Series.Add(worksheetGrafieken.Cells["C4:C5"], worksheetGrafieken.Cells["C4:C5"]);
                doughnutChart.Fill.Style = eFillStyle.SolidFill;
                doughnutChart.Fill.Color = System.Drawing.Color.FromArgb(255, 247, 247);



                //Diagrammen grafieken worksheet
                //var diagram1 = worksheetGrafieken.Drawings["Chart 1"] as ExcelChart;
                //var serieKosten = diagram1.Series.Add(worksheetGrafieken.Cells[rowNumber, 4, rowNumber, 4], worksheetGrafieken.Cells[rowNumber, 4]);
                //serieKosten.HeaderAddress = worksheetGrafieken.Cells[rowNumber, 1];

                //var diagram2 = worksheetGrafieken.Drawings["Chart 10"] as ExcelChart;
                //var serieInkomsten = diagram2.Series.Add(worksheetGrafieken.Cells[rowNumber, 8, rowNumber, 10], worksheetGrafieken.Cells[rowNumber, 1]);
                //serieInkomsten.HeaderAddress = worksheetGrafieken.Cells[rowNumber, 1];
                //diagramInkomsten.SetPosition(44, 0, 11, 0);

                foreach (var drawing in worksheetGrafieken.Drawings)
                {
                    if (drawing is ExcelChart)
                    {
                        // Handle chart-specific logic
                        var chart = (ExcelChart)drawing;
                        Console.WriteLine($"Chart Name: {chart.Name}");
                        // Add more chart-specific logic as needed
                    }
                    else if (drawing is ExcelPicture)
                    {
                        // Handle picture-specific logic
                        var picture = (ExcelPicture)drawing;
                        Console.WriteLine($"Picture Name: {picture.Name}");
                        // Add more picture-specific logic as needed
                    }
                    // Add more conditions for other types of drawings (e.g., shapes)

                    // Common properties for all drawings
                    Console.WriteLine($"Drawing Type: {drawing.GetType().Name}");
                    Console.WriteLine($"Description: {drawing.Description}");
                    Console.WriteLine($"Position: {drawing.From.Column}, {drawing.From.Row}");
                    Console.WriteLine($"Size: {drawing.To.Column - drawing.From.Column}, {drawing.To.Row - drawing.From.Row}");
                    Console.WriteLine();
                }

                package.Save();
            }
        }
    }
}
