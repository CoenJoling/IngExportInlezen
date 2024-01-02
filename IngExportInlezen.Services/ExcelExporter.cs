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
                ExcelWorksheet overzicht = package.Workbook.Worksheets[0];
                ExcelWorksheet abonnementen = package.Workbook.Worksheets[1];
                ExcelWorksheet boodschappen = package.Workbook.Worksheets[2];
                ExcelWorksheet geldOpnames = package.Workbook.Worksheets[3];
                ExcelWorksheet inkomstenSalaris = package.Workbook.Worksheets[4];
                ExcelWorksheet overigeInkomsten = package.Workbook.Worksheets[5];
                ExcelWorksheet overigeKosten = package.Workbook.Worksheets[6];
                ExcelWorksheet spaarOpdrachten = package.Workbook.Worksheets[7];
                ExcelWorksheet tanken = package.Workbook.Worksheets[8];
                ExcelWorksheet vasteLasten = package.Workbook.Worksheets[9];

                var rowNumberOverzicht = overzicht.Dimension.Rows + 2;
                var rowNumberAbonnementen = abonnementen.Dimension.Rows + 2;
                var rowNumberBoodschappen = boodschappen.Dimension.Rows + 2;
                var rowNumberGeldOpnames = geldOpnames.Dimension.Rows + 2;
                var rowNumberInkomstenSalaris = inkomstenSalaris.Dimension.Rows + 2;
                var rowNumberOverigeInkomsten = overigeInkomsten.Dimension.Rows + 2;
                var rowNumberOverigeKosten = overigeKosten.Dimension.Rows + 2;
                var rowNumberSpaarOpdrachten = spaarOpdrachten.Dimension.Rows + 2;
                var rowNumberTanken = tanken.Dimension.Rows + 2;
                var rowNumberVasteLasten = vasteLasten.Dimension.Rows + 2;

                //PopulateWorksheet(abonnementen, rowNumberAbonnementen, excelExport.Abonnementen, excelExport);
                //PopulateWorksheet(boodschappen, rowNumberBoodschappen, excelExport.Boodschappen, excelExport);
                //PopulateWorksheet(geldOpnames, rowNumberGeldOpnames, excelExport.GeldOpnames, excelExport);
                //PopulateWorksheet(inkomstenSalaris, rowNumberInkomstenSalaris, excelExport.InkomstenSalaris, excelExport);
                //PopulateWorksheet(overigeInkomsten, rowNumberOverigeInkomsten, excelExport.OverigeInkomsten, excelExport);
                //PopulateWorksheet(tanken, rowNumberTanken, excelExport.Tanken, excelExport);
                //PopulateWorksheet(vasteLasten, rowNumberVasteLasten, excelExport.VasteLasten, excelExport);

                #region Overige kosten
                //overigeKosten.Cells[rowNumberOverigeKosten, 2].Value = excelExport.Maand;
                //var maandCellOk = overigeKosten.Cells[$"B{rowNumberOverigeKosten}:C{rowNumberOverigeKosten}"];
                //maandCellOk.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //maandCellOk.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 220, 220));
                //maandCellOk.Style.Font.Bold = true;

                //for (var i = 0; i < excelExport.OverigeKosten.Count(); i++)
                //{
                //    var entry = excelExport.OverigeKosten[i];
                //    overigeKosten.Cells[rowNumberOverigeKosten + i + 1, 2].Value = entry.Naam;
                //    overigeKosten.Cells[rowNumberOverigeKosten + i + 1, 3].Value = entry.Bedrag;

                //    var range = overigeKosten.Cells[$"B{rowNumberOverigeKosten + i + 1}:C{rowNumberOverigeKosten + i + 1}"];
                //    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 220, 220));
                //}
                #endregion

                #region Spaaropdrachten
                //excelExport.SpaarOpdrachtenIngelegd = excelExport.SpaarOpdrachtenIngelegd.OrderBy(x => x.Datum).ToList();
                //excelExport.SpaarOpdrachtenOpgenomen = excelExport.SpaarOpdrachtenOpgenomen.OrderBy(x => x.Datum).ToList();
                //spaarOpdrachten.Cells[rowNumberSpaarOpdrachten, 2].Value = excelExport.Maand;
                //var maandCellSo = spaarOpdrachten.Cells[$"B{rowNumberSpaarOpdrachten}:F{rowNumberSpaarOpdrachten}"];
                //maandCellSo.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //maandCellSo.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 220, 220));
                //maandCellSo.Style.Font.Bold = true;

                //rowNumberSpaarOpdrachten = spaarOpdrachten.Dimension.Rows + 2;
                //spaarOpdrachten.Cells[rowNumberSpaarOpdrachten, 2].Value = "Spaaropdrachten ingelegd";
                //var inlegCellSo = spaarOpdrachten.Cells[$"B{rowNumberSpaarOpdrachten}:F{rowNumberSpaarOpdrachten}"];
                //inlegCellSo.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //inlegCellSo.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 220, 220));
                //inlegCellSo.Style.Font.Bold = true;

                //for (var i = 0; i < excelExport.SpaarOpdrachtenIngelegd.Count(); i++)
                //{
                //    var entry = excelExport.SpaarOpdrachtenIngelegd[i];
                //    spaarOpdrachten.Cells[rowNumberSpaarOpdrachten + i + 1, 2].Value = entry.Datum.ToString("dd MMMM yyyy");
                //    spaarOpdrachten.Cells[rowNumberSpaarOpdrachten + i + 1, 3].Value = entry.Naam;
                //    spaarOpdrachten.Cells[rowNumberSpaarOpdrachten + i + 1, 4].Value = entry.Code;
                //    spaarOpdrachten.Cells[rowNumberSpaarOpdrachten + i + 1, 5].Value = entry.Bedrag;
                //    spaarOpdrachten.Cells[rowNumberSpaarOpdrachten + i + 1, 6].Value = entry.Mededelingen;

                //    var range = spaarOpdrachten.Cells[$"B{rowNumberSpaarOpdrachten + i + 1}:F{rowNumberSpaarOpdrachten + i + 1}"];
                //    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 220, 220));
                //}

                //rowNumberSpaarOpdrachten = spaarOpdrachten.Dimension.Rows + 2;
                //spaarOpdrachten.Cells[rowNumberSpaarOpdrachten, 2].Value = "Spaaropdrachten opgenomen";
                //var opnameCellSo = spaarOpdrachten.Cells[$"B{rowNumberSpaarOpdrachten}:F{rowNumberSpaarOpdrachten}"];
                //opnameCellSo.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //opnameCellSo.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 220, 220));
                //opnameCellSo.Style.Font.Bold = true;

                //for (var i = 0; i < excelExport.SpaarOpdrachtenOpgenomen.Count(); i++)
                //{
                //    var entry = excelExport.SpaarOpdrachtenOpgenomen[i];
                //    spaarOpdrachten.Cells[rowNumberSpaarOpdrachten + i + 1, 2].Value = entry.Datum.ToString("dd MMMM yyyy");
                //    spaarOpdrachten.Cells[rowNumberSpaarOpdrachten + i + 1, 3].Value = entry.Naam;
                //    spaarOpdrachten.Cells[rowNumberSpaarOpdrachten + i + 1, 4].Value = entry.Code;
                //    spaarOpdrachten.Cells[rowNumberSpaarOpdrachten + i + 1, 5].Value = entry.Bedrag * -1;
                //    spaarOpdrachten.Cells[rowNumberSpaarOpdrachten + i + 1, 6].Value = entry.Mededelingen;

                //    var range = spaarOpdrachten.Cells[$"B{rowNumberSpaarOpdrachten + i + 1}:F{rowNumberSpaarOpdrachten + i + 1}"];
                //    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 220, 220));
                //}
                #endregion

                #region Overzicht
                //for (var i = 2; i < 29; i += 3)
                //{
                //    overzicht.Cells[rowNumberOverzicht, i].Value = excelExport.Maand;
                //}

                //PopulateOverzicht(overzicht, 3, rowNumberOverzicht, excelExport.Abonnementen);
                //PopulateOverzicht(overzicht, 6, rowNumberOverzicht, excelExport.Boodschappen);
                //PopulateOverzicht(overzicht, 9, rowNumberOverzicht, excelExport.GeldOpnames);
                //PopulateOverzicht(overzicht, 12, rowNumberOverzicht, excelExport.InkomstenSalaris);
                //PopulateOverzicht(overzicht, 15, rowNumberOverzicht, excelExport.OverigeInkomsten);
                //PopulateOverzicht(overzicht, 18, rowNumberOverzicht, excelExport.OverigeKosten);
                //var spaarBedrag = excelExport.SpaarOpdrachtenIngelegd.Sum(x => x.Bedrag) - excelExport.SpaarOpdrachtenOpgenomen.Sum(x => x.Bedrag);
                //overzicht.Cells[rowNumberOverzicht, 21].Value = spaarBedrag;
                //var cell1 = overzicht.Cells[rowNumberOverzicht, 20];
                //cell1.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //cell1.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 220, 220));
                //var cell2 = overzicht.Cells[rowNumberOverzicht, 21];
                //cell2.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //cell2.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 220, 220));
                //PopulateOverzicht(overzicht, 24, rowNumberOverzicht, excelExport.Tanken);
                //PopulateOverzicht(overzicht, 27, rowNumberOverzicht, excelExport.VasteLasten);
                #endregion

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
                if (existingPieChart != null)
                {
                    worksheetOverzicht.Drawings.Remove(existingPieChart);
                };
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
                if (existingDoughnutChart != null )
                {
                    worksheetOverzicht.Drawings.Remove(existingDoughnutChart);
                }
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

        private static void PopulateOverzicht(ExcelWorksheet worksheet, int columnNumber, int rowNumberOverzicht, List<IngExport_Internal> data)
        {
            var sumBedrag = data.Sum(x => x.Bedrag);
            worksheet.Cells[rowNumberOverzicht, columnNumber].Value = sumBedrag;

            var cell1 = worksheet.Cells[rowNumberOverzicht, columnNumber-1];
            cell1.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            cell1.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 220, 220));

            var cell2 = worksheet.Cells[rowNumberOverzicht, columnNumber];
            cell2.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            cell2.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 220, 220));
        }
        private static void PopulateWorksheet(ExcelWorksheet worksheet, int rowNumber, List<IngExport_Internal> data, ExcelExport maand)
        {
            data = data.OrderBy(x => x.Datum).ToList();

            worksheet.Cells[rowNumber, 2].Value = maand.Maand;
            var maandCell = worksheet.Cells[$"B{rowNumber}:F{rowNumber}"];
            maandCell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            maandCell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 220, 220));
            maandCell.Style.Font.Bold = true;

            for (var i = 0; i < data.Count(); i++)
            {
                var entry = data[i];
                worksheet.Cells[rowNumber + i + 1, 2].Value = entry.Datum.ToString("dd MMMM yyyy");
                worksheet.Cells[rowNumber + i + 1, 3].Value = entry.Naam;
                worksheet.Cells[rowNumber + i + 1, 4].Value = entry.Code;
                worksheet.Cells[rowNumber + i + 1, 5].Value = entry.Bedrag;
                worksheet.Cells[rowNumber + i + 1, 6].Value = entry.Mededelingen;

                var range = worksheet.Cells[$"B{rowNumber + i + 1}:F{rowNumber + i + 1}"];
                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 220, 220));
            }
        }
    }
}
