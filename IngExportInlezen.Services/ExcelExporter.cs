using IngExportInlezen.Domain;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;

namespace IngExportInlezen.Services
{
    public static class ExcelExporter
    {
        private static readonly Random rnd = new Random();

        static ExcelExporter()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public static void ExportToJaaroverzichtExcel(ExcelExport excelExport, AppSettings appSettings)
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
                ExcelWorksheet vasteLasten = package.Workbook.Worksheets[1];
                ExcelWorksheet abonnementen = package.Workbook.Worksheets[2];
                ExcelWorksheet tanken = package.Workbook.Worksheets[3];
                ExcelWorksheet geldOpnames = package.Workbook.Worksheets[4];
                ExcelWorksheet boodschappen = package.Workbook.Worksheets[5];
                ExcelWorksheet overigeKosten = package.Workbook.Worksheets[6];
                ExcelWorksheet spaarOpdrachten = package.Workbook.Worksheets[7];
                ExcelWorksheet inkomstenSalaris = package.Workbook.Worksheets[8];
                ExcelWorksheet overigeInkomsten = package.Workbook.Worksheets[9];

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

                PopulateJaarWorksheet(abonnementen, rowNumberAbonnementen, excelExport.Abonnementen, excelExport);
                PopulateJaarWorksheet(boodschappen, rowNumberBoodschappen, excelExport.Boodschappen, excelExport);
                PopulateJaarWorksheet(geldOpnames, rowNumberGeldOpnames, excelExport.GeldOpnames, excelExport);
                PopulateJaarWorksheet(inkomstenSalaris, rowNumberInkomstenSalaris, excelExport.InkomstenSalaris, excelExport);
                PopulateJaarWorksheet(overigeInkomsten, rowNumberOverigeInkomsten, excelExport.OverigeInkomsten, excelExport);
                PopulateJaarWorksheet(tanken, rowNumberTanken, excelExport.Tanken, excelExport);
                PopulateJaarWorksheet(vasteLasten, rowNumberVasteLasten, excelExport.VasteLasten, excelExport);

                #region Overige kosten
                overigeKosten.Cells[rowNumberOverigeKosten, 2].Value = excelExport.Maand;
                var maandCellOk = overigeKosten.Cells[$"B{rowNumberOverigeKosten}:C{rowNumberOverigeKosten}"];
                maandCellOk.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                maandCellOk.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 220, 220));
                maandCellOk.Style.Font.Bold = true;

                for (var i = 0; i < excelExport.OverigeKosten.Count(); i++)
                {
                    var entry = excelExport.OverigeKosten[i];
                    overigeKosten.Cells[rowNumberOverigeKosten + i + 1, 2].Value = entry.Naam;
                    overigeKosten.Cells[rowNumberOverigeKosten + i + 1, 3].Value = entry.Bedrag * -1;

                    var range = overigeKosten.Cells[$"B{rowNumberOverigeKosten + i + 1}:C{rowNumberOverigeKosten + i + 1}"];
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 220, 220));
                }
                #endregion

                #region Spaaropdrachten
                excelExport.SpaarOpdrachtenIngelegd = excelExport.SpaarOpdrachtenIngelegd.OrderBy(x => x.Datum).ToList();
                excelExport.SpaarOpdrachtenOpgenomen = excelExport.SpaarOpdrachtenOpgenomen.OrderBy(x => x.Datum).ToList();
                spaarOpdrachten.Cells[rowNumberSpaarOpdrachten, 2].Value = excelExport.Maand;
                var maandCellSo = spaarOpdrachten.Cells[$"B{rowNumberSpaarOpdrachten}:F{rowNumberSpaarOpdrachten}"];
                maandCellSo.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                maandCellSo.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 220, 220));
                maandCellSo.Style.Font.Bold = true;

                rowNumberSpaarOpdrachten = spaarOpdrachten.Dimension.Rows + 2;
                spaarOpdrachten.Cells[rowNumberSpaarOpdrachten, 2].Value = "Spaaropdrachten ingelegd";
                var inlegCellSo = spaarOpdrachten.Cells[$"B{rowNumberSpaarOpdrachten}:F{rowNumberSpaarOpdrachten}"];
                inlegCellSo.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                inlegCellSo.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 220, 220));
                inlegCellSo.Style.Font.Bold = true;

                for (var i = 0; i < excelExport.SpaarOpdrachtenIngelegd.Count(); i++)
                {
                    var entry = excelExport.SpaarOpdrachtenIngelegd[i];
                    spaarOpdrachten.Cells[rowNumberSpaarOpdrachten + i + 1, 2].Value = entry.Datum.ToString("dd MMMM yyyy");
                    spaarOpdrachten.Cells[rowNumberSpaarOpdrachten + i + 1, 3].Value = entry.Naam;
                    spaarOpdrachten.Cells[rowNumberSpaarOpdrachten + i + 1, 4].Value = entry.Code;
                    spaarOpdrachten.Cells[rowNumberSpaarOpdrachten + i + 1, 5].Value = entry.Bedrag;
                    spaarOpdrachten.Cells[rowNumberSpaarOpdrachten + i + 1, 6].Value = entry.Mededelingen;

                    var range = spaarOpdrachten.Cells[$"B{rowNumberSpaarOpdrachten + i + 1}:F{rowNumberSpaarOpdrachten + i + 1}"];
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 220, 220));
                }

                rowNumberSpaarOpdrachten = spaarOpdrachten.Dimension.Rows + 2;
                spaarOpdrachten.Cells[rowNumberSpaarOpdrachten, 2].Value = "Spaaropdrachten opgenomen";
                var opnameCellSo = spaarOpdrachten.Cells[$"B{rowNumberSpaarOpdrachten}:F{rowNumberSpaarOpdrachten}"];
                opnameCellSo.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                opnameCellSo.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 220, 220));
                opnameCellSo.Style.Font.Bold = true;

                for (var i = 0; i < excelExport.SpaarOpdrachtenOpgenomen.Count(); i++)
                {
                    var entry = excelExport.SpaarOpdrachtenOpgenomen[i];
                    spaarOpdrachten.Cells[rowNumberSpaarOpdrachten + i + 1, 2].Value = entry.Datum.ToString("dd MMMM yyyy");
                    spaarOpdrachten.Cells[rowNumberSpaarOpdrachten + i + 1, 3].Value = entry.Naam;
                    spaarOpdrachten.Cells[rowNumberSpaarOpdrachten + i + 1, 4].Value = entry.Code;
                    spaarOpdrachten.Cells[rowNumberSpaarOpdrachten + i + 1, 5].Value = entry.Bedrag * -1;
                    spaarOpdrachten.Cells[rowNumberSpaarOpdrachten + i + 1, 6].Value = entry.Mededelingen;

                    var range = spaarOpdrachten.Cells[$"B{rowNumberSpaarOpdrachten + i + 1}:F{rowNumberSpaarOpdrachten + i + 1}"];
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 220, 220));
                }
                #endregion

                #region Overzicht
                for (var i = 2; i < 29; i += 3)
                {
                    overzicht.Cells[rowNumberOverzicht, i].Value = excelExport.Maand;
                }

                PopulateJaarOverzicht(overzicht, 6, rowNumberOverzicht, excelExport.Abonnementen);
                PopulateJaarOverzicht(overzicht, 15, rowNumberOverzicht, excelExport.Boodschappen);
                PopulateJaarOverzicht(overzicht, 12, rowNumberOverzicht, excelExport.GeldOpnames);
                PopulateJaarOverzicht(overzicht, 24, rowNumberOverzicht, excelExport.InkomstenSalaris, false, true);
                PopulateJaarOverzicht(overzicht, 27, rowNumberOverzicht, excelExport.OverigeInkomsten, false, true);
                PopulateJaarOverzicht(overzicht, 18, rowNumberOverzicht, excelExport.OverigeKosten, true);
                var spaarBedrag = excelExport.SpaarOpdrachtenIngelegd.Sum(x => x.Bedrag) - excelExport.SpaarOpdrachtenOpgenomen.Sum(x => x.Bedrag);
                overzicht.Cells[rowNumberOverzicht, 21].Value = spaarBedrag;
                var cell1 = overzicht.Cells[rowNumberOverzicht, 20];
                cell1.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                cell1.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 220, 220));
                var cell2 = overzicht.Cells[rowNumberOverzicht, 21];
                cell2.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                cell2.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 220, 220));
                if ((double)spaarBedrag < appSettings.Spaardoel)
                {
                    cell2.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(252, 90, 10));
                }
                else
                {
                    cell2.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 220, 220));
                }
                PopulateJaarOverzicht(overzicht, 9, rowNumberOverzicht, excelExport.Tanken);
                PopulateJaarOverzicht(overzicht, 3, rowNumberOverzicht, excelExport.VasteLasten);

                overzicht.Cells["C3"].Formula = $"SUM(C5:C{rowNumberOverzicht})";
                overzicht.Cells["F3"].Formula = $"SUM(F5:F{rowNumberOverzicht})";
                overzicht.Cells["I3"].Formula = $"SUM(I5:I{rowNumberOverzicht})";
                overzicht.Cells["L3"].Formula = $"SUM(L5:L{rowNumberOverzicht})";
                overzicht.Cells["O3"].Formula = $"SUM(O5:O{rowNumberOverzicht})";
                overzicht.Cells["R3"].Formula = $"SUM(R5:R{rowNumberOverzicht})";
                overzicht.Cells["U3"].Formula = $"SUM(U5:U{rowNumberOverzicht})";
                overzicht.Cells["X3"].Formula = $"SUM(X5:X{rowNumberOverzicht})";
                overzicht.Cells["AA3"].Formula = $"SUM(AA5:AA{rowNumberOverzicht})";
                overzicht.Calculate();

                #endregion

                package.Save();
            }
        }

        public static void ExportToMaandExcel(ExcelExport excelExport, AppSettings appSettings)
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
                ExcelWorksheet worksheetPieCharts = package.Workbook.Worksheets[1];
                ExcelWorksheet worksheetBarCharts = package.Workbook.Worksheets[2];

                var rowNumber = worksheetOverzicht.Dimension.Rows + 2;

                worksheetOverzicht.Cells[rowNumber, 2].Value = excelExport.Maand;
                var maandCell = worksheetOverzicht.Cells[$"B{rowNumber}"];
                maandCell.Style.Font.Bold = true;

                var range = worksheetOverzicht.Cells[$"B{rowNumber}:K{rowNumber}"];
                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 220, 220));

                PopulateMaandOverzicht(rowNumber, 3, excelExport.VasteLasten, worksheetOverzicht);
                PopulateMaandOverzicht(rowNumber, 4, excelExport.Abonnementen, worksheetOverzicht);
                PopulateMaandOverzicht(rowNumber, 5, excelExport.Tanken, worksheetOverzicht);
                PopulateMaandOverzicht(rowNumber, 6, excelExport.GeldOpnames, worksheetOverzicht);
                PopulateMaandOverzicht(rowNumber, 7, excelExport.Boodschappen, worksheetOverzicht);
                PopulateMaandOverzicht(rowNumber, 8, excelExport.OverigeKosten, worksheetOverzicht, false, true);
                PopulateMaandOverzicht(rowNumber, 9, excelExport.SpaarOpdrachtenIngelegd, worksheetOverzicht, false, false, true, excelExport.SpaarOpdrachtenOpgenomen, appSettings);
                PopulateMaandOverzicht(rowNumber, 10, excelExport.InkomstenSalaris, worksheetOverzicht, true);
                PopulateMaandOverzicht(rowNumber, 11, excelExport.OverigeInkomsten, worksheetOverzicht, true);

                #region PieChart Overzicht
                //Maken van pie chart.
                var existingPieChart = worksheetOverzicht.Drawings["SpreidingKosten"];
                if (existingPieChart != null)
                {
                    worksheetOverzicht.Drawings.Remove(existingPieChart);
                };
                var pieChart = worksheetOverzicht.Drawings.AddChart("SpreidingKosten", eChartType.Pie);
                var serie = pieChart.Series.Add(worksheetOverzicht.Cells[rowNumber, 3, rowNumber, 8], worksheetOverzicht.Cells[2, 3, 2, 8]);
                pieChart.SetPosition(rowNumber + 1, 0, 1, 0);
                pieChart.SetSize(650, 450);
                pieChart.Title.Text = $"Spreiding kosten {excelExport.Maand}";
                pieChart.Title.Font.Bold = true;
                pieChart.Title.Font.Size = 16;
                pieChart.Legend.Position = eLegendPosition.Left;
                pieChart.Legend.Font.Size = 13;
                var pieSerie = (ExcelPieChartSerie)serie;
                pieSerie.DataLabel.ShowCategory = true;
                pieSerie.DataLabel.ShowPercent = true;
                pieSerie.DataLabel.ShowLeaderLines = true;
                pieSerie.DataLabel.Position = eLabelPosition.OutEnd;
                pieChart.Fill.Style = eFillStyle.SolidFill;
                pieChart.Fill.Color = System.Drawing.Color.FromArgb(255, 247, 247);
                #endregion

                #region BudgetChart
                // Maken van bugdget chart.
                var budgetChart = worksheetOverzicht.Drawings["Budget"];
                if (budgetChart != null)
                {
                    worksheetOverzicht.Drawings.Remove(budgetChart);
                }
                var columnChart = worksheetOverzicht.Drawings.AddChart("Budget", eChartType.ColumnClustered);
                columnChart.SetPosition(rowNumber + 1, 0, 8, 0);
                columnChart.SetSize(650, 450);
                var serie1 = columnChart.Series.Add("G3:I3", "G2:I2");
                var serie2 = columnChart.Series.Add($"G{rowNumber}:I{rowNumber}", "G2:I2");
                columnChart.Legend.Remove();
                columnChart.YAxis.MinorTickMark = eAxisTickMark.None;
                columnChart.PlotArea.Fill.Style = eFillStyle.SolidFill;
                columnChart.PlotArea.Fill.Color = System.Drawing.Color.FromArgb(255, 247, 247);
                columnChart.Fill.Style = eFillStyle.SolidFill;
                columnChart.Fill.Color = System.Drawing.Color.FromArgb(255, 247, 247);
                columnChart.Title.Text = $"Budget vs werkelijkheid {excelExport.Maand}";
                columnChart.Title.Font.Bold = true;
                columnChart.Title.Font.Size = 16;

                var cellG3 = Convert.ToDecimal(worksheetOverzicht.Cells["G3"].Value);
                var cellH3 = Convert.ToDecimal(worksheetOverzicht.Cells["H3"].Value);
                var cellI3 = Convert.ToDecimal(worksheetOverzicht.Cells["I3"].Value);
                var newCellG = Convert.ToDecimal(worksheetOverzicht.Cells[$"G{rowNumber}"].Value);
                var newCellH = Convert.ToDecimal(worksheetOverzicht.Cells[$"H{rowNumber}"].Value);
                var newCellI = Convert.ToDecimal(worksheetOverzicht.Cells[$"I{rowNumber}"].Value);

                if (cellG3 < newCellG || cellH3 < newCellH || cellI3 > newCellI)
                {
                    columnChart.Series[1].Fill.Color = System.Drawing.Color.FromArgb(255, 0, 0);
                }
                else
                {
                    columnChart.Series[1].Fill.Color = System.Drawing.Color.FromArgb(0, 255, 0);
                }
                #endregion

                #region PieCharts
                //Maken en behouden piechart in worksheetPieCharts
                var row = 1;
                var col = 1;
                for (var i = 0; worksheetPieCharts.Drawings.Count >= i; i++)
                {
                    if (!IsChartPresentAtCell(worksheetPieCharts, row, col))
                    {
                        var pieChart2 = worksheetPieCharts.Drawings.AddChart($"SpreidingKosten{rnd.Next(100000)}", eChartType.Pie);
                        var seriee = pieChart2.Series.Add(worksheetOverzicht.Cells[rowNumber, 3, rowNumber, 8], worksheetOverzicht.Cells[2, 3, 2, 8]);
                        pieChart2.SetPosition(row, 0, col, 0);
                        pieChart2.SetSize(630, 400);
                        pieChart2.Title.Text = $"Spreiding kosten {excelExport.Maand}";
                        pieChart2.Title.Font.Bold = true;
                        pieChart2.Title.Font.Size = 16;
                        pieChart2.Legend.Position = eLegendPosition.Left;
                        pieChart2.Legend.Font.Size = 11;
                        var pieSerie2 = (ExcelPieChartSerie)seriee;
                        pieSerie2.DataLabel.ShowCategory = true;
                        pieSerie2.DataLabel.ShowPercent = true;
                        pieSerie2.DataLabel.ShowLeaderLines = true;
                        pieSerie2.DataLabel.Position = eLabelPosition.OutEnd;
                        pieChart2.Fill.Style = eFillStyle.SolidFill;
                        pieChart2.Fill.Color = System.Drawing.Color.FromArgb(255, 247, 247);
                        break;
                    }
                    col +=11;
                    if (!IsChartPresentAtCell(worksheetPieCharts, row, col))
                    {
                        var pieChart2 = worksheetPieCharts.Drawings.AddChart($"SpreidingKosten{rnd.Next(100000)}", eChartType.Pie);
                        var seriee = pieChart2.Series.Add(worksheetOverzicht.Cells[rowNumber, 3, rowNumber, 8], worksheetOverzicht.Cells[2, 3, 2, 8]);
                        pieChart2.SetPosition(row, 0, col, 0);
                        pieChart2.SetSize(630, 400);
                        pieChart2.Title.Text = $"Spreiding kosten {excelExport.Maand}";
                        pieChart2.Title.Font.Bold = true;
                        pieChart2.Title.Font.Size = 16;
                        pieChart2.Legend.Position = eLegendPosition.Left;
                        pieChart2.Legend.Font.Size = 11;
                        var pieSerie2 = (ExcelPieChartSerie)seriee;
                        pieSerie2.DataLabel.ShowCategory = true;
                        pieSerie2.DataLabel.ShowPercent = true;
                        pieSerie2.DataLabel.ShowLeaderLines = true;
                        pieSerie2.DataLabel.Position = eLabelPosition.OutEnd;
                        pieChart2.Fill.Style = eFillStyle.SolidFill;
                        pieChart2.Fill.Color = System.Drawing.Color.FromArgb(255, 247, 247);
                        break;
                    }
                    row += 21;
                    col -= 11;
                }
                #endregion

                #region BarCharts
                //Diagrammen grafieken worksheet
                MaakBarChartMaand(worksheetBarCharts, worksheetOverzicht, "Abonnementen", $"D4:D{rowNumber}", $"B4:B{rowNumber}", 1, 1);
                MaakBarChartMaand(worksheetBarCharts, worksheetOverzicht, "Vaste lasten", $"C4:C{rowNumber}", $"B4:B{rowNumber}", 1, 12);
                MaakBarChartMaand(worksheetBarCharts, worksheetOverzicht, "Tanken", $"E4:E{rowNumber}", $"B4:B{rowNumber}", 22, 1);
                MaakBarChartMaand(worksheetBarCharts, worksheetOverzicht, "Geld opname", $"F4:F{rowNumber}", $"B4:B{rowNumber}", 22, 12);
                MaakBarChartMaand(worksheetBarCharts, worksheetOverzicht, "Boodschappen", $"G4:G{rowNumber}", $"B4:B{rowNumber}", 43, 1);
                MaakBarChartMaand(worksheetBarCharts, worksheetOverzicht, "Spaaropdrachten", $"I4:I{rowNumber}", $"B4:B{rowNumber}", 43, 12);
                MaakBarChartMaand(worksheetBarCharts, worksheetOverzicht, "Inkomsten salaris", $"J4:J{rowNumber}", $"B4:B{rowNumber}", 64, 1);
                MaakBarChartMaand(worksheetBarCharts, worksheetOverzicht, "Overige inkomsten", $"K4:K{rowNumber}", $"B4:B{rowNumber}", 64, 12);
                MaakBarChartMaand(worksheetBarCharts, worksheetOverzicht, "Overige kosten", $"H4:H{rowNumber}", $"B4:B{rowNumber}", 85, 1);
                #endregion

                package.Save();
            }
        }

        private static void MaakBarChartMaand(ExcelWorksheet worksheetBarCharts, ExcelWorksheet worksheetOverzicht, string chartName, string dataRange, string categoryRange, int positionX, int positionY)
        {
            var existingChart = worksheetBarCharts.Drawings[chartName] as ExcelChart;
            if (existingChart != null)
            {
                worksheetBarCharts.Drawings.Remove(existingChart);
            }
            var chart = worksheetBarCharts.Drawings.AddChart(chartName, eChartType.ColumnClustered);
            var series = chart.Series.Add(worksheetOverzicht.Cells[dataRange], worksheetOverzicht.Cells[categoryRange]);
            series.Fill.Color = System.Drawing.Color.FromArgb(255, 217, 102);
            chart.Title.Text = chartName;
            chart.Title.Font.Bold = true;
            chart.Title.Font.Size = 16;
            chart.Legend.Remove();
            chart.SetPosition(positionX, 0, positionY, 0);
            chart.SetSize(650, 390);
            chart.YAxis.MinorTickMark = eAxisTickMark.None;
            chart.XAxis.MajorTickMark = eAxisTickMark.None;
            chart.XAxis.MinorTickMark = eAxisTickMark.None;
            chart.Fill.Style = eFillStyle.SolidFill;
            chart.Fill.Color = System.Drawing.Color.FromArgb(255, 247, 247);
            chart.PlotArea.Fill.Color = System.Drawing.Color.FromArgb(255, 247, 247);
        }

        private static void PopulateMaandOverzicht(int rowNumber, int columnNUmber, List<IngExport_Internal> data, ExcelWorksheet worksheetOverzicht, bool isInkomsten = false, bool isOverigeKosten = false, bool isSpaaropdrachten = false, List<IngExport_Internal> data2 = null, AppSettings appSettings = null)
        {
            decimal bedrag;
            if (isOverigeKosten)
            {
                bedrag = data.Sum(x => x.Bedrag) * -1;
            }
            else if (isSpaaropdrachten)
            {
                bedrag = data.Sum(x => x.Bedrag) - data2.Sum(x => x.Bedrag);
            }
            else
            {
                bedrag = data.Sum(x => x.Bedrag);
            }

            worksheetOverzicht.Cells[rowNumber, columnNUmber].Value = bedrag;

            var cellValues = new List<double>();
            var cells = worksheetOverzicht.Cells[4, columnNUmber, rowNumber, columnNUmber];
            foreach (var cell in cells)
            {
                cellValues.Add(Convert.ToDouble(cell.Value));
            }
            var average = cellValues.Average();
            var stDev = StdDev(cellValues, average);

            var avgBij = average + stDev;
            var avgAf = average - stDev;

            var cell2 = worksheetOverzicht.Cells[rowNumber, columnNUmber];
            cell2.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            if (isInkomsten)
            {
                if (avgBij < (double)bedrag || (double)bedrag < avgAf)
                {
                    cell2.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(252, 90, 10));
                }
            }
            else if (isSpaaropdrachten)
            {
                if (bedrag < appSettings.Spaardoel)
                {
                    cell2.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(252, 90, 10));
                }
            }
            else
            {
                if ((double)bedrag > avgBij)
                {
                    cell2.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(252, 90, 10));
                }
            }
        }

        private static bool IsChartPresentAtCell(ExcelWorksheet worksheet, int targetRow, int targetColumn)
        {
            foreach (var drawing in worksheet.Drawings)
            {
                if (drawing is ExcelChart chart)
                {
                    // Check if the chart position matches the target cell
                    if (chart.From.Row <= targetRow && chart.To.Row >= targetRow &&
                        chart.From.Column <= targetColumn && chart.To.Column >= targetColumn)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private static void PopulateJaarOverzicht(ExcelWorksheet worksheet, int columnNumber, int rowNumberOverzicht, List<IngExport_Internal> data, bool isOverigeKosten = false, bool isInkomsten = false)
        {
            decimal sumBedrag;
            if(isOverigeKosten)
            {
                sumBedrag = data.Sum(x => x.Bedrag) * -1;
            }
            else
            {
                sumBedrag = data.Sum(x => x.Bedrag);
            }
            
            worksheet.Cells[rowNumberOverzicht, columnNumber].Value = sumBedrag;

            var cellValues = new List<double>();
            var cells = worksheet.Cells[5, columnNumber,rowNumberOverzicht,columnNumber];
            foreach (var cell in cells)
            {
                cellValues.Add(Convert.ToDouble(cell.Value));
            }
            var average = cellValues.Average();
            var stDev = StdDev(cellValues, average);

            var avgBij = average + stDev;
            var avgAf = average - stDev;

            var cell1 = worksheet.Cells[rowNumberOverzicht, columnNumber - 1];
            cell1.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            cell1.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 220, 220));

            var cell2 = worksheet.Cells[rowNumberOverzicht, columnNumber];
            cell2.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            if (isInkomsten)
            {
                if (avgBij < (double)sumBedrag || (double)sumBedrag < avgAf)
                {
                    cell2.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(252, 90, 10));
                }
                else
                {
                    cell2.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 220, 220));
                }
            }
            else
            {
                if ((double)sumBedrag > avgBij)
                {
                    cell2.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(252, 90, 10));
                }
                else
                {
                    cell2.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 220, 220));
                }
            }
        }

        private static double StdDev(List<double> values, double average)
        {
            var sumOfSquares = values.Select(val => Math.Pow(val - average, 2)).Sum();
            var variance = sumOfSquares / values.Count;
            var standardDeviation = Math.Sqrt(variance);
            return standardDeviation;
        }

        private static void PopulateJaarWorksheet(ExcelWorksheet worksheet, int rowNumber, List<IngExport_Internal> data, ExcelExport maand)
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
