using CsvHelper.Configuration.Attributes;

namespace IngExportInlezen.Domain
{
    public class IngExport_File
    {
        [Name("Datum")]
        public string Datum { get; set; }

        [Name("Naam / Omschrijving")]
        public string Naam { get; set; }

        [Name("Rekening")]
        public string Rekening { get; set; }

        [Name("Tegenrekening")]
        public string Tegenrekening { get; set; }

        [Name("Code")]
        public string Code { get; set; }

        [Name("Af Bij")]
        public string AfBij {  get; set; }

        [Name("Bedrag (EUR)")]
        public string Bedrag { get; set; }

        [Name("Mutatiesoort")]
        public string MutatieSoort { get; set; }

        [Name("Mededelingen")]
        public string Mededelingen { get; set; }

        [Name("Saldo na mutatie")]
        public string SaldoNaMutatie { get; set; }

        [Name("Tag")]
        public string Tag { get; set; }
    }
}