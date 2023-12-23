namespace IngExportInlezen.Domain
{
    public static class IngExportMappings
    {
        public static IngExport_Internal MapFiletoInternal(IngExport_File input)
        {
            return new IngExport_Internal
            {
                Datum = DateTime.ParseExact(input.Datum, "yyyyMMdd", null),
                Naam = input.Naam.ToLower(),
                Rekening = input.Rekening.ToLower(),
                Tegenrekening = input.Tegenrekening.ToLower(),
                Code = input.Code.ToLower(),
                AfBij = input.AfBij.ToLower(),
                Bedrag = Convert.ToDecimal(input.Bedrag),
                MutatieSoort = input.MutatieSoort.ToLower(),
                Mededelingen = input.Mededelingen.ToLower(),
                SaldoNaMutatie = Convert.ToDecimal(input.SaldoNaMutatie),
                Tag = input.Tag.ToLower(),
            };
        }
    }
}
