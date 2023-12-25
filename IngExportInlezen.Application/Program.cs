using CsvHelper;
using CsvHelper.Configuration;
using IngExportInlezen.Domain;
using IngExportInlezen.Services;
using Microsoft.Extensions.Configuration;
using System.Globalization;

namespace IngExportInlezen.Application
{
    class Program
    {
        static void Main()
        {
            Console.WriteLine("\nZorg ervoor dat 1 ING export bestand in de map staat!\n");
            Task.Delay(2000).Wait();
            Console.WriteLine("\nControleren...\n");
            Task.Delay(2000).Wait();

            string folderPath = @"C:\Users\coenj\Documents\Financieel overzicht\ING export\";

            string[] files = Directory.GetFiles(folderPath);

            string targetFileExtension = ".csv";
            string targetFileNameSubstring = "NL81INGB0008739620";

            string[] matchingFiles = files
                .Where(file =>
                    Path.GetExtension(file).Equals(targetFileExtension, StringComparison.OrdinalIgnoreCase) &&
                    Path.GetFileNameWithoutExtension(file).Contains(targetFileNameSubstring, StringComparison.OrdinalIgnoreCase))
                .ToArray();

            if (matchingFiles.Length == 0)
            {
                Console.Clear();
                Console.WriteLine("Geen geschikt bestand gevonden in de map");
                Console.ReadLine();
                return;
            }
            else if (matchingFiles.Length > 1)
            {
                Console.Clear();
                Console.WriteLine("Er staan meerdere bestanden met de opgegeven criteria in de map. App gestopt!");
                Console.ReadLine();
                return;
            }
            else
            {
                string csvInput = matchingFiles[0];

                Console.WriteLine("\nInlezen ING Export...\n");

                Task.Delay(1000).Wait();

                Console.WriteLine("...");

                Task.Delay(1000).Wait();
                Console.Clear();

                var config = new CsvConfiguration(CultureInfo.InvariantCulture)
                {
                    Delimiter = ";",
                };

                IConfiguration configuration = new ConfigurationBuilder()
               .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
               .Build();

                var appSettings = new AppSettings();
                configuration.GetSection("AppSettings").Bind(appSettings);

                using (var reader = new StreamReader(csvInput))
                using (var csv = new CsvReader(reader, config))
                {
                    var records = csv.GetRecords<IngExport_File>();
                    var completeCsvList = new List<IngExport_Internal>();

                    foreach (var record in records)
                    {
                        var singleCsvRow = IngExportMappings.MapFiletoInternal(record);
                        completeCsvList.Add(singleCsvRow);
                    }

                    for (int i = 0; i < completeCsvList.Count; i++)
                    {
                        completeCsvList[i].Id = i + 1;
                    }

                    var laatsteDatum = completeCsvList.First().Datum.ToShortDateString();
                    var eersteDatum = completeCsvList.Last().Datum.ToShortDateString();
                    Console.WriteLine($"\nBestand ingelezen voor de periode van: {eersteDatum} tm {laatsteDatum}\n");
                    var resultList = new List<string>();
                    var assignedLineList = new List<IngExport_Internal>();
                    var unassignedEntries = new List<IngExport_Internal>();

                    var excelExport = new ExcelExport();
                    excelExport.Maand = completeCsvList.FirstOrDefault().Datum.ToString("MMMM yyyy");
                    excelExport.Abonnementen = ConsoleServices.Abonnementen(appSettings, completeCsvList, laatsteDatum, eersteDatum, resultList, assignedLineList);
                    excelExport.VasteLasten = ConsoleServices.VasteLasten(appSettings, completeCsvList, laatsteDatum, eersteDatum, resultList, assignedLineList);
                    excelExport.Boodschappen = ConsoleServices.Boodschappen(appSettings, completeCsvList, laatsteDatum, eersteDatum, resultList, assignedLineList);
                    excelExport.GeldOpnames = ConsoleServices.GeldOpname(completeCsvList, laatsteDatum, eersteDatum, resultList, assignedLineList, unassignedEntries);
                    excelExport.Tanken = ConsoleServices.Tanken(appSettings, completeCsvList, laatsteDatum, eersteDatum, resultList, assignedLineList);
                    excelExport.InkomstenSalaris = ConsoleServices.InkomstenSalaris(appSettings, completeCsvList, laatsteDatum, eersteDatum, resultList, assignedLineList);
                    excelExport.OverigeInkomsten = ConsoleServices.OverigeInkomsten(completeCsvList, laatsteDatum, eersteDatum, resultList, assignedLineList);
                    excelExport.SpaarOpdrachten = ConsoleServices.Spaaropdrachten(appSettings, completeCsvList, resultList, assignedLineList);
                    excelExport.OverigeKosten = (ConsoleServices.ResultatenEnOverigeKosten(completeCsvList, resultList, assignedLineList, unassignedEntries)) * -1;

                    Console.WriteLine("\nMaak een keuze \n" +
                        "Druk op 1 om de resultaten in Excel te importen\n" +
                        "Druk op een knop om af te sluiten\n");
                    var input = Console.ReadKey().KeyChar.ToString();

                    switch(input)
                    {
                        case "1":
                            try
                            {
                                ExcelExporter.ExportToExcel(excelExport);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                                Console.ReadKey();
                                throw;
                            }
                            Console.WriteLine("\nDe gegevens zijn succesvol in de Excel sheet geschreven!\n");
                            break;
                        default:
                            break;
                    }
                    Console.ReadKey();
                }
            }
        }
    }        
}
