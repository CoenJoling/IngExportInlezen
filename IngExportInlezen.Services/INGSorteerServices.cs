using IngExportInlezen.Domain;

namespace IngExportInlezen.Services
{
    public static class INGSorteerServices
    {
        public static List<IngExport_Internal> ResultatenEnOverigeKosten(List<IngExport_Internal> completeCsvList, List<string> resultList, List<IngExport_Internal> assignedLineList, List<IngExport_Internal> unassignedEntries)
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
            return naamSumList;
        }

        public static (List<IngExport_Internal> inleggenList, List<IngExport_Internal> opgenomenList) Spaaropdrachten(AppSettings appSettings, List<IngExport_Internal> completeCsvList, List<string> resultList, List<IngExport_Internal> assignedLineList)
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
            var netto = inleggenSpaaropdrachtenBedrag - opgenomenpaaropdrachtenBedrag;
            resultList.Add($"Het totale bedrag voor spaaropdrachten is: {inleggenSpaaropdrachtenBedrag}");
            resultList.Add($"De totale hoeveelheid opgenomen van spaarrekeningen is: {opgenomenpaaropdrachtenBedrag}");
            resultList.Add($"Netto gespaard: {netto}\n");
            return (inleggenSpaaropdrachtenList, opgenomenSpaaropdrachtenList);
        }

        public static List<IngExport_Internal> OverigeInkomsten(List<IngExport_Internal> completeCsvList, string laatsteDatum, string eersteDatum, List<string> resultList, List<IngExport_Internal> assignedLineList)
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
            return overigeInkomstenList;
        }

        public static List<IngExport_Internal> InkomstenSalaris(AppSettings appSettings, List<IngExport_Internal> completeCsvList, string laatsteDatum, string eersteDatum, List<string> resultList, List<IngExport_Internal> assignedLineList)
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
            return salarisList;
        }

        public static List<IngExport_Internal> Tanken(AppSettings appSettings, List<IngExport_Internal> completeCsvList, string laatsteDatum, string eersteDatum, List<string> resultList, List<IngExport_Internal> assignedLineList)
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
            return tankenList;
        }

        public static List<IngExport_Internal> GeldOpname(List<IngExport_Internal> completeCsvList, string laatsteDatum, string eersteDatum, List<string> resultList, List<IngExport_Internal> assignedLineList, List<IngExport_Internal> unassignedEntries)
        {
            var geldOpnameList = new List<IngExport_Internal>();

            var filteredGMItems = completeCsvList.Where(x => x.Code.Contains("GM".ToLower()) && x.AfBij.Equals("Af".ToLower())).ToList();
            geldOpnameList.AddRange(filteredGMItems);
            assignedLineList.AddRange(filteredGMItems);
            unassignedEntries.AddRange(filteredGMItems);

            var geldOpnameBedrag = geldOpnameList.Sum(x => x.Bedrag);
            resultList.Add($"Geld opnames: {geldOpnameBedrag}");
            return geldOpnameList;
        }

        public static List<IngExport_Internal> Boodschappen(AppSettings appSettings, List<IngExport_Internal> completeCsvList, string laatsteDatum, string eersteDatum, List<string> resultList, List<IngExport_Internal> assignedLineList)
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

            if (boodschappenList != null && boodschappenList.Any(entry => entry.Naam.Contains("Albert Heijn".ToLower())))
            {
                var gemiddeldeAlbertHeijn = Math.Round(boodschappenList
                    .Where(entry => entry.Naam.Contains("Albert Heijn".ToLower()))
                    .Average(entry => entry.Bedrag), 2);

                resultList.Add($"Gemiddeld bij AH(n={ahCount}): {gemiddeldeAlbertHeijn}");
            }

            if (boodschappenList != null && boodschappenList.Any(entry => entry.Naam.Contains("Plus".ToLower())))
            {
                var gemiddeldePlus = Math.Round(boodschappenList
                    .Where(entry => entry.Naam.Contains("Plus".ToLower()))
                    .Average(entry => entry.Bedrag), 2);

                resultList.Add($"Gemiddeld bij Plus(n={plusCount}): {gemiddeldePlus}");
            }

            var boodschappenBedrag = boodschappenList.Sum(x => x.Bedrag);
            resultList.Add($"Kosten boodschappen: {boodschappenBedrag}\n");
            return boodschappenList;
        }

        public static List<IngExport_Internal> VasteLasten(AppSettings appSettings, List<IngExport_Internal> completeCsvList, string laatsteDatum, string eersteDatum, List<string> resultList, List<IngExport_Internal> assignedLineList)
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
            return vasteLastenList;
        }

        public static List<IngExport_Internal> Abonnementen(AppSettings appSettings, List<IngExport_Internal> completeCsvList, string laatsteDatum, string eersteDatum, List<string> resultList, List<IngExport_Internal> assignedLineList)
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
            return abonnementenList;
        }
    }
}