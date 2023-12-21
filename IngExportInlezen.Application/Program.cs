using CsvHelper;
using CsvHelper.Configuration;
using IngExportInlezen.Domain;
using Microsoft.Extensions.Configuration;
using System.Globalization;
using System.Threading.Tasks;

namespace IngExportInlezen.Application
{
    class Program
    {
        static void Main()
        {
            //Console.WriteLine("Vul nu de locatie in waar het ING Export bestand is opgeslagen.");
            //var inputLocatie = Console.ReadLine();
            //Console.WriteLine(@"Vul nu de naam van het ING Export bestand in. Bijvoorbeeld: ""voorbeeld.csv"".");
            //var inputFile = Console.ReadLine();
            //var csvFile = inputLocatie + @"\" + inputFile;

            //Jaar
            //var csvInput = @"C:\Users\coenj\Documents\Financieel overzicht\ING input\NL81INGB0008739620_01-01-2023_17-12-2023.csv";
            //Maand
            var csvInput = @"C:\Users\coenj\Documents\Financieel overzicht\ING input\NL81INGB0008739620_01-11-2023_30-11-2023.csv";

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
                var resultList = new List<string>();
                var assignedLineList = new List<IngExport_Internal>();
                var unassignedEntries = new List<IngExport_Internal>();

                var state = true;
                while (state)
                {
                    Console.WriteLine("\nWat moet er gebeuren?\n" +
                        "1) abo\n" +
                        "2) vaste lasten\n" +
                        "3) boodschappen\n" +
                        "4) Geld opname\n" +
                        "5) Tanken\n" +
                        "6) Inkomsten salaris\n" +
                        "7) Inkomsten overige\n" +
                        "8) Spaaropdrachten\n" +
                        "9) Lijst resultaten\n" +
                        "Klik op een knop om af te sluiten");
                    
                    var inputVraag = Console.ReadLine();

                    switch (inputVraag)
                    {
                        case "1": //Abonnementen
                            Console.Clear();
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
                            resultList.Add($"Abonnement: {abonnementenBedrag}");
                            Console.WriteLine($"Het totale bedrag betreft abonnementen voor de periode {eersteDatum} - {laatsteDatum} is: {abonnementenBedrag}");
                            break;

                        case "2": //Vaste lasten
                            Console.Clear();
                            var vasteLastenList = new List<IngExport_Internal>();
                            foreach (var last in appSettings.VasteLasten)
                            {
                                var filteredItems = completeCsvList.Where(x => x.Naam.Contains(last.ToLower()) && x.AfBij.Equals("Af".ToLower())).ToList();
                                vasteLastenList.AddRange(filteredItems);
                                assignedLineList.AddRange(filteredItems);
                            }

                            var vastenLastenBedrag = vasteLastenList.Sum(x => x.Bedrag);
                            resultList.Add($"Vaste lasten: {vastenLastenBedrag}");
                            Console.WriteLine($"Het totale bedrag betreft vaste lasten voor de periode {eersteDatum} - {laatsteDatum} is: {vastenLastenBedrag}");
                            break;

                        case "3": //Boodschappen
                            Console.Clear();
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
                                .Average(entry => entry.Bedrag),2);

                            var gemiddeldePlus = Math.Round(boodschappenList.Where(entry => entry.Naam.Contains("Plus".ToLower()))
                                .Average(entry => entry.Bedrag),2);

                            var boodschappenBedrag = boodschappenList.Sum(x => x.Bedrag);
                            resultList.Add($"Boodschappen: {boodschappenBedrag}");
                            resultList.Add($"Gemiddeld bij AH(n={ahCount}): {gemiddeldeAlbertHeijn}");
                            resultList.Add($"Gemiddeld bij Plus(n={plusCount}): {gemiddeldePlus}");
                            Console.WriteLine($"Het totale bedrag betreft boodschappen voor de periode {eersteDatum} - {laatsteDatum} is: {boodschappenBedrag}\n");
                            Console.WriteLine($"Het gemiddelde bedrag betreft boodschappen bij de Albert Heijn(n={ahCount}) voor de periode {eersteDatum} - {laatsteDatum} is: {gemiddeldeAlbertHeijn}\n");
                            Console.WriteLine($"Het gemiddelde bedrag betreft boodschappen bij de Plus(n={plusCount}) voor de periode {eersteDatum} - {laatsteDatum} is: {gemiddeldePlus}\n");
                            break;

                        case "4": //Geld opname
                            Console.Clear();
                            var geldOpnameList = new List<IngExport_Internal>();

                            var filteredGMItems = completeCsvList.Where(x => x.Code.Contains("GM".ToLower()) && x.AfBij.Equals("Af".ToLower())).ToList();
                            geldOpnameList.AddRange(filteredGMItems);
                            assignedLineList.AddRange(filteredGMItems);
                            unassignedEntries.AddRange(filteredGMItems);

                            var geldOpnameBedrag = geldOpnameList.Sum(x => x.Bedrag);
                            resultList.Add($"Geld opname: {geldOpnameBedrag}");
                            Console.WriteLine($"Het totale bedrag betreft geldopnames voor de periode {eersteDatum} - {laatsteDatum} is: {geldOpnameBedrag}");
                            break;

                        case "5": //Tanken
                            Console.Clear();
                            var tankenList = new List<IngExport_Internal>();
                            foreach (var tank in appSettings.Tanken)
                            {
                                var filteredItems = completeCsvList.Where(x => x.Naam.Contains(tank.ToLower()) && x.AfBij.Equals("Af".ToLower())).ToList();
                                tankenList.AddRange(filteredItems);
                                assignedLineList.AddRange(filteredItems);
                            }

                            var tankenBedrag = tankenList.Sum(x => x.Bedrag);
                            resultList.Add($"Tanken: {tankenBedrag}");
                            Console.WriteLine($"Het totale bedrag betreft tanken voor de periode {eersteDatum} - {laatsteDatum} is: {tankenBedrag}");
                            break;

                        case "6": //Inkomsten salaris
                            Console.Clear();
                            var salarisList = new List<IngExport_Internal>();
                            foreach (var salaris in appSettings.InkomstenSalaris)
                            {
                                var filteredItems = completeCsvList.Where(x => x.Naam.Contains(salaris.ToLower()) && x.AfBij.Equals("Bij".ToLower())).ToList();
                                salarisList.AddRange(filteredItems);
                                assignedLineList.AddRange(filteredItems);
                            }
                            var salarisBedrag = salarisList.Sum(x => x.Bedrag);
                            resultList.Add($"Inkomsten salaris: {salarisBedrag}");
                            Console.WriteLine($"Het totale bedrag betreft inkomsten voor de periode {eersteDatum} - {laatsteDatum} is: {salarisBedrag}");
                            break;

                        case "7": //Overige inkomsten
                            Console.Clear();
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
                            Console.WriteLine($"Het totale bedrag betreft overige inkomsten voor de periode {eersteDatum} - {laatsteDatum} is: {inkomstenBedrag}");
                            break;

                        case "8": //Spaaropdrachten
                            Console.Clear();
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
                            resultList.Add($"Netto gespaard: {inleggenSpaaropdrachtenBedrag - opgenomenpaaropdrachtenBedrag}");
                            Console.WriteLine($"Het totale maandbedrag voor spaaropdrachten is: {inleggenSpaaropdrachtenBedrag}\n" +
                                $"Het totale bedrag van opgenomen van de spaarrekeningen is: {opgenomenpaaropdrachtenBedrag}\n" +
                                $"Netto gespaard: {inleggenSpaaropdrachtenBedrag - opgenomenpaaropdrachtenBedrag}");
                            break;

                        case "9":
                            Console.Clear();

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
                            Console.WriteLine($"\nHet bedrag van overige uitgave is: {bedragRest}\n");

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
                                Console.WriteLine($"Overige kosten posten: {uniqueNaam.Naam} || {uniqueNaam.Bedrag}");
                            }

                            break;

                        default:
                            Console.Clear();
                            state = false;
                            break;
                    }
                }
            }
        }
    }
}
