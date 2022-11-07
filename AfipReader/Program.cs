using Aspose.Cells;

Workbook excel = new("C:\\Users\\diiee\\Desktop\\DDC.xlsx");

Worksheet page = excel.Worksheets[0];

List<Comprobante> DetalleComprobantes = new List<Comprobante>();

Comprobante comprobante = new Comprobante();
int colTipoComp = 11;
int colNetoGravado = 11;
int colNoGravado = 12;
int colOpExentas = 13;
int colIVA = 14;
int colTotalComp = 15;


for (int row = 0; row <= page.Cells.MaxDataRow; row++)
{


	for (int col = 0; col <= page.Cells.MaxDataColumn; col++)
	{

        if (col==colTipoComp && row > 1)
        {
            if (!DetalleComprobantes.Exists(x => x.Nombretipocomp == page.Cells[row, col].StringValue))
            {
                DetalleComprobantes.Add(new Comprobante() { Nombretipocomp = page.Cells[row, col].StringValue });
            }
            

        }






		if (col==colNetoGravado&&row>1)
		{
            
            switch (page.Cells[row, colIVA].DoubleValue/page.Cells[row, colNetoGravado].DoubleValue)
            {
                case 0.21:
                    comprobante.Alicuota21.Netogravado += page.Cells[row, col].DoubleValue;
                    comprobante.Alicuota21.Iva+= page.Cells[row, colIVA].DoubleValue;
                    comprobante.Alicuota21.Total = comprobante.Alicuota21.Netogravado + comprobante.Alicuota21.Iva;

                    break;
                case 0.27:
                    comprobante.Alicuota27.Netogravado += page.Cells[row, col].DoubleValue;
                    comprobante.Alicuota27.Iva += page.Cells[row, colIVA].DoubleValue;
                    comprobante.Alicuota27.Total = comprobante.Alicuota27.Netogravado + comprobante.Alicuota27.Iva;

                    break;
                case 0.105:
                    comprobante.Alicuota105.Netogravado += page.Cells[row, col].DoubleValue;
                    comprobante.Alicuota105.Iva += page.Cells[row, colIVA].DoubleValue;
                    comprobante.Alicuota105.Total = comprobante.Alicuota105.Netogravado + comprobante.Alicuota105.Iva;
                    break;

                default:
                    comprobante.AlicuotaVarias.Netogravado += page.Cells[row, col].DoubleValue;
                    comprobante.AlicuotaVarias.Iva += page.Cells[row, colIVA].DoubleValue;
                    comprobante.AlicuotaVarias.Total = comprobante.AlicuotaVarias.Netogravado + comprobante.AlicuotaVarias.Iva;
                    break;
            }
            comprobante.NetoNogravado += page.Cells[row, colNoGravado].DoubleValue;
            comprobante.OpExentas += page.Cells[row, colOpExentas].DoubleValue;
            comprobante.Total += page.Cells[row, colTotalComp].DoubleValue;
        }
       //Console.Write(page.Cells[row, col].Value + " | ");

    }

    
    Console.WriteLine();

}



Console.Write("neto21: ");
Console.WriteLine(comprobante.Alicuota21.Netogravado + " | " + comprobante.Alicuota21.Iva);

Console.Write("neto105: ");
Console.WriteLine(comprobante.Alicuota105.Netogravado + " | " + comprobante.Alicuota105.Iva);

Console.Write("neto27: ");
Console.WriteLine(comprobante.Alicuota27.Netogravado + " | " + comprobante.Alicuota27.Iva);

Console.Write("Ex: ");
Console.WriteLine(comprobante.OpExentas + " | " );

Console.Write("No gravado: ");
Console.WriteLine(comprobante.NetoNogravado + " | " );

Console.Write("total: ");
Console.WriteLine(comprobante.Total + " | " );

Console.WriteLine(); Console.WriteLine(); Console.WriteLine(); Console.WriteLine();



Console.ReadKey();



class Comprobante   
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

class Importes
{
    public double Netogravado { get; set; } = 0;
    public double Iva { get; set; } = 0;
    public double Total { get; set; } = 0;
}