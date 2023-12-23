using CsvHelper;
using CsvHelper.Configuration;
using IngExportInlezen.Domain;
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
                return;
            }
            else if (matchingFiles.Length > 1)
            {
                Console.Clear();
                Console.WriteLine("Er staan meerdere bestanden met de opgegeven criteria in de map. App gestopt!");
                return;
            }
            else
            {
                string csvInput = matchingFiles[0];

                Console.WriteLine("\nInlezen ING Export...\n");

                Task.Delay(1000).Wait();

                Console.WriteLine("...");

                Task.Delay(1000).Wait();
                Console.Clear() ;

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

                    Abonnementen(appSettings, completeCsvList, laatsteDatum, eersteDatum, resultList, assignedLineList);
                    VasteLasten(appSettings, completeCsvList, laatsteDatum, eersteDatum, resultList, assignedLineList);
                    Boodschappen(appSettings, completeCsvList, laatsteDatum, eersteDatum, resultList, assignedLineList);
                    GeldOpname(completeCsvList, laatsteDatum, eersteDatum, resultList, assignedLineList, unassignedEntries);
                    Tanken(appSettings, completeCsvList, laatsteDatum, eersteDatum, resultList, assignedLineList);
                    InkomstenSalaris(appSettings, completeCsvList, laatsteDatum, eersteDatum, resultList, assignedLineList);
                    OverigeInkomsten(completeCsvList, laatsteDatum, eersteDatum, resultList, assignedLineList);
                    Spaaropdrachten(appSettings, completeCsvList, resultList, assignedLineList);
                    ResultatenEnOverigeKosten(completeCsvList, resultList, assignedLineList, unassignedEntries);

                    Console.WriteLine("\nDruk op een toets om af te sluiten.");
                    Console.ReadLine();
                }
            }
        }

        private static void ResultatenEnOverigeKosten(List<IngExport_Internal> completeCsvList, List<string> resultList, List<IngExport_Internal> assignedLineList, List<IngExport_Internal> unassignedEntries)
        {
            foreach (string entry in resultList)
            {
                Console.WriteLine(entry);
            }

            unassignedEntries.AddRange(completeCsvList.Except(assignedLineList));

            decimal bedragRest = 0;
            foreach (var rest in unassignedEntries)
            {
                if (rest.AfBij == "Af".ToLower())
                {
                    bedragRest -= rest.Bedrag;
                }
                else if (rest.AfBij == "Bij".ToLower())
                {
                    bedragRest += rest.Bedrag;
                }
            }
            Console.WriteLine($"\nBedrag van overige uitgaves is: {bedragRest}\n");

            var naamSumList = unassignedEntries
                .GroupBy(entry => entry.Naam)
                .Select(group => new IngExport_Internal
                {
                    Naam = group.Key,
                    Bedrag = group.Sum(entry => entry.AfBij == "Af".ToLower() ? -entry.Bedrag : entry.Bedrag)
                }).OrderBy(entry => entry.Naam)
                .ToList();

            foreach (var uniqueNaam in naamSumList)
            {
                Console.WriteLine($"Overige kosten post: {uniqueNaam.Naam} || {uniqueNaam.Bedrag}");
            }
        }

        private static void Spaaropdrachten(AppSettings appSettings, List<IngExport_Internal> completeCsvList, List<string> resultList, List<IngExport_Internal> assignedLineList)
        {
            var inleggenSpaaropdrachtenList = new List<IngExport_Internal>();
            var opgenomenSpaaropdrachtenList = new List<IngExport_Internal>();
            foreach (var inleggenSpaaropdracht in appSettings.Spaaropdrachten)
            {
                var filteredItems = completeCsvList.Where(x => x.Mededelingen.Contains(inleggenSpaaropdracht.ToLower()) && x.AfBij.Equals("Af".ToLower())).ToList();
                inleggenSpaaropdrachtenList.AddRange(filteredItems);
                assignedLineList.AddRange(filteredItems);
            }

            var bitvavo = completeCsvList.Where(x => x.Code.Contains("Bitvavo".ToLower()) && x.AfBij.Equals("Af".ToLower())).ToList();
            inleggenSpaaropdrachtenList.AddRange(bitvavo);
            assignedLineList.AddRange(bitvavo);

            foreach (var inleggenSpaaropdracht in appSettings.Spaaropdrachten)
            {
                var filteredItems = completeCsvList.Where(x => x.Mededelingen.Contains(inleggenSpaaropdracht.ToLower()) && x.AfBij.Equals("Bij".ToLower())).ToList();
                opgenomenSpaaropdrachtenList.AddRange(filteredItems);
                assignedLineList.AddRange(filteredItems);
            }

            var inleggenSpaaropdrachtenBedrag = inleggenSpaaropdrachtenList.Sum(x => x.Bedrag);
            var opgenomenpaaropdrachtenBedrag = opgenomenSpaaropdrachtenList.Sum(x => x.Bedrag);
            resultList.Add($"Het totale bedrag voor spaaropdrachten is: {inleggenSpaaropdrachtenBedrag}");
            resultList.Add($"De totale hoeveelheid opgenomen van spaarrekeningen is: {opgenomenpaaropdrachtenBedrag}");
            resultList.Add($"Netto gespaard: {inleggenSpaaropdrachtenBedrag - opgenomenpaaropdrachtenBedrag}");
        }

        private static void OverigeInkomsten(List<IngExport_Internal> completeCsvList, string laatsteDatum, string eersteDatum, List<string> resultList, List<IngExport_Internal> assignedLineList)
        {
            var overigeInkomstenList = new List<IngExport_Internal>();

            var filteredInkomstenItems = completeCsvList.
                Where(x => !x.Naam.Contains("Working Talent".ToLower()) &&
                !x.Naam.Contains("Organon".ToLower()) &&
                !x.Naam.Contains("ABBOTT BIOLOGICALS".ToLower()) &&
                !x.Naam.Contains("c joling-quist,hr c joling".ToLower()) &&
                !x.Naam.Contains("mw ing c joling-quist".ToLower()) &&
                !x.Naam.Contains("Oranje Spaarrekening".ToLower()) &&
                x.AfBij.Equals("Bij".ToLower())).ToList();
            overigeInkomstenList.AddRange(filteredInkomstenItems);
            assignedLineList.AddRange(filteredInkomstenItems);

            var inkomstenBedrag = overigeInkomstenList.Sum(x => x.Bedrag);
            resultList.Add($"Overige inkomsten: {inkomstenBedrag}");
        }

        private static void InkomstenSalaris(AppSettings appSettings, List<IngExport_Internal> completeCsvList, string laatsteDatum, string eersteDatum, List<string> resultList, List<IngExport_Internal> assignedLineList)
        {
            var salarisList = new List<IngExport_Internal>();
            foreach (var salaris in appSettings.InkomstenSalaris)
            {
                var filteredItems = completeCsvList.Where(x => x.Naam.Contains(salaris.ToLower()) && x.AfBij.Equals("Bij".ToLower())).ToList();
                salarisList.AddRange(filteredItems);
                assignedLineList.AddRange(filteredItems);
            }
            var salarisBedrag = salarisList.Sum(x => x.Bedrag);
            resultList.Add($"Inkomsten salaris: {salarisBedrag}");
        }

        private static void Tanken(AppSettings appSettings, List<IngExport_Internal> completeCsvList, string laatsteDatum, string eersteDatum, List<string> resultList, List<IngExport_Internal> assignedLineList)
        {
            var tankenList = new List<IngExport_Internal>();
            foreach (var tank in appSettings.Tanken)
            {
                var filteredItems = completeCsvList.Where(x => x.Naam.Contains(tank.ToLower()) && x.AfBij.Equals("Af".ToLower())).ToList();
                tankenList.AddRange(filteredItems);
                assignedLineList.AddRange(filteredItems);
            }

            var tankenBedrag = tankenList.Sum(x => x.Bedrag);
            resultList.Add($"Kosten tanken: {tankenBedrag}");
        }

        private static void GeldOpname(List<IngExport_Internal> completeCsvList, string laatsteDatum, string eersteDatum, List<string> resultList, List<IngExport_Internal> assignedLineList, List<IngExport_Internal> unassignedEntries)
        {
            var geldOpnameList = new List<IngExport_Internal>();

            var filteredGMItems = completeCsvList.Where(x => x.Code.Contains("GM".ToLower()) && x.AfBij.Equals("Af".ToLower())).ToList();
            geldOpnameList.AddRange(filteredGMItems);
            assignedLineList.AddRange(filteredGMItems);
            unassignedEntries.AddRange(filteredGMItems);

            var geldOpnameBedrag = geldOpnameList.Sum(x => x.Bedrag);
            resultList.Add($"Geld opnames: {geldOpnameBedrag}");
        }

        private static void Boodschappen(AppSettings appSettings, List<IngExport_Internal> completeCsvList, string laatsteDatum, string eersteDatum, List<string> resultList, List<IngExport_Internal> assignedLineList)
        {
            var boodschappenList = new List<IngExport_Internal>();
            foreach (var winkel in appSettings.Boodschappen)
            {
                var filteredItems = completeCsvList.Where(x => x.Naam.Contains(winkel.ToLower()) && x.AfBij.Equals("Af".ToLower())).ToList();
                boodschappenList.AddRange(filteredItems);
                assignedLineList.AddRange(filteredItems);
            }

            var ahCount = boodschappenList.Count(x => x.Naam.Contains("Albert Heijn".ToLower()));
            var plusCount = boodschappenList.Count(x => x.Naam.Contains("Plus".ToLower()));

            var gemiddeldeAlbertHeijn = Math.Round(boodschappenList.Where(entry => entry.Naam.Contains("Albert Heijn".ToLower()))
                .Average(entry => entry.Bedrag), 2);

            var gemiddeldePlus = Math.Round(boodschappenList.Where(entry => entry.Naam.Contains("Plus".ToLower()))
                .Average(entry => entry.Bedrag), 2);

            var boodschappenBedrag = boodschappenList.Sum(x => x.Bedrag);
            resultList.Add($"Kosten boodschappen: {boodschappenBedrag}");
            resultList.Add($"Gemiddeld bij AH(n={ahCount}): {gemiddeldeAlbertHeijn}");
            resultList.Add($"Gemiddeld bij Plus(n={plusCount}): {gemiddeldePlus}");
        }

        private static void VasteLasten(AppSettings appSettings, List<IngExport_Internal> completeCsvList, string laatsteDatum, string eersteDatum, List<string> resultList, List<IngExport_Internal> assignedLineList)
        {
            var vasteLastenList = new List<IngExport_Internal>();
            foreach (var last in appSettings.VasteLasten)
            {
                var filteredItems = completeCsvList.Where(x => x.Naam.Contains(last.ToLower()) && x.AfBij.Equals("Af".ToLower())).ToList();
                vasteLastenList.AddRange(filteredItems);
                assignedLineList.AddRange(filteredItems);
            }

            var vastenLastenBedrag = vasteLastenList.Sum(x => x.Bedrag);
            resultList.Add($"Kosten vaste lasten: {vastenLastenBedrag}");
        }

        private static void Abonnementen(AppSettings appSettings, List<IngExport_Internal> completeCsvList, string laatsteDatum, string eersteDatum, List<string> resultList, List<IngExport_Internal> assignedLineList)
        {
            var abonnementenList = new List<IngExport_Internal>();
            foreach (var abonnement in appSettings.Abonnementen)
            {
                var filteredItems = completeCsvList.Where(x => x.Naam.Contains(abonnement.ToLower()) && x.AfBij.Equals("Af".ToLower())).ToList();
                abonnementenList.AddRange(filteredItems);
                assignedLineList.AddRange(filteredItems);
            }

            var netflix = completeCsvList.Where(x => x.Mededelingen.Contains("Netflix".ToLower())).ToList();
            abonnementenList.AddRange(netflix);
            assignedLineList.AddRange(netflix);

            var dvo = completeCsvList.Where(x => x.Naam.Contains("D.V.O.".ToLower()) && x.Mededelingen.Contains("contr".ToLower())).ToList();
            abonnementenList.AddRange(dvo);
            assignedLineList.AddRange(dvo);

            var abonnementenBedrag = abonnementenList.Sum(x => x.Bedrag);
            resultList.Add($"Kosten abonnementen: {abonnementenBedrag}");
        }
    }
}
