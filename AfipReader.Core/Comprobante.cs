namespace AfipReader.Core
{
    public class Comprobante
    {
        public string Nombretipocomp { get; set; } = "";
        public double NetoNogravado { get; set; } = 0;
        public double OpExentas { get; set; } = 0;
        public double Total { get; set; } = 0;
        public Importes Alicuota21 { get; set; } = new Importes();
        public Importes Alicuota105 { get; set; } = new Importes();
        public Importes Alicuota27 { get; set; } = new Importes();
        public Importes AlicuotaVarias { get; set; } = new Importes();
    }
}