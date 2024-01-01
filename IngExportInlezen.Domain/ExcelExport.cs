namespace IngExportInlezen.Domain
{
    public class ExcelExport
    {
        public string Maand {  get; set; }
        public List<IngExport_Internal> Abonnementen { get; set; }
        public List<IngExport_Internal> VasteLasten { get; set; }
        public List<IngExport_Internal> Boodschappen { get; set; }
        public List<IngExport_Internal> GeldOpnames { get; set; }
        public List<IngExport_Internal> Tanken {  get; set; }
        public List<IngExport_Internal> OverigeKosten { get; set; }
        public List<IngExport_Internal> InkomstenSalaris { get; set; }
        public List<IngExport_Internal> OverigeInkomsten { get; set; }
        public List<IngExport_Internal> SpaarOpdrachtenIngelegd { get; set; }
        public List<IngExport_Internal> SpaarOpdrachtenOpgenomen { get; set; }
    }
}
