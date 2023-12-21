using CsvHelper.Configuration.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IngExportInlezen.Domain
{
    public class IngExport_Internal
    {
        public int Id { get; set; }

        public DateTime Datum { get; set; }

        public string Naam { get; set; }

        public string AfBij { get; set; }

        public string Code { get; set; }

        public decimal Bedrag { get; set; }

        public string Mededelingen { get; set; }

        public string Rekening { get; set; }

        public string Tegenrekening { get; set; }

        public string MutatieSoort { get; set; }

        public decimal SaldoNaMutatie { get; set; }

        public string Tag { get; set; }
    }
}
