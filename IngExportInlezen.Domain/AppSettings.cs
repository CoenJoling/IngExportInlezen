using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IngExportInlezen.Domain
{
    public class AppSettings
    {
        public List<string> Boodschappen { get; set; }
        public List<string> Tanken { get; set; }
        public List<string> Abonnementen { get; set; }
        public List<string> VasteLasten { get; set; }
        public List<string> InkomstenSalaris { get; set; }
        public List<string> Spaaropdrachten { get; set; }
    }
}
