using Aspose.Cells;
using AfipReader.Core;

Workbook excel = new("C:\\Users\\diiee\\Desktop\\DDC.xlsx");

Worksheet page = excel.Worksheets[0];

List<Comprobante> DetalleComprobantes = new List<Comprobante>();

Comprobante comprobante = new Comprobante();
TODOChange add = new TODOChange();
int colTipoComp = 11;
int colNetoGravado = 11;
int colNoGravado = 12;
int colOpExentas = 13;
int colIVA = 14;
int colTotalComp = 15;


for (int row = 0; row <= page.Cells.MaxDataRow; row++)
{
    // TODO SEGRAGAR EN METODOS

	for (int col = 0; col <= page.Cells.MaxDataColumn; col++)
	{

        if (col==colTipoComp && row > 1)

        {

            if (!DetalleComprobantes.Exists(x => x.Nombretipocomp == page.Cells[row, col].StringValue))
            {
                DetalleComprobantes.Add(new Comprobante() { Nombretipocomp = page.Cells[row, col].StringValue });
            }
            int index = DetalleComprobantes.IndexOf(DetalleComprobantes.FirstOrDefault(x => x.Nombretipocomp == page.Cells[row, col].StringValue));
            
        }






		if (col==colNetoGravado&&row>1)
		{
            
            switch (page.Cells[row, colIVA].DoubleValue/page.Cells[row, colNetoGravado].DoubleValue)
            {
                case 0.21:
                    
                    add.ValuesToList(comprobante,
                                               comprobante.Alicuota21,
                                               page,
                                               row,
                                               colNetoGravado,
                                               colIVA,
                                               colNoGravado,
                                               colOpExentas,
                                               colTotalComp);

                    break;
                case 0.27:
                    add.ValuesToList(comprobante,
                                               comprobante.Alicuota27,
                                               page,
                                               row,
                                               colNetoGravado,
                                               colIVA,
                                               colNoGravado,
                                               colOpExentas,
                                               colTotalComp);

                    break;
                case 0.105:
                    add.ValuesToList(comprobante,
                                                comprobante.Alicuota105,
                                                page,
                                                row,
                                                colNetoGravado,
                                                colIVA,
                                                colNoGravado,
                                                colOpExentas,
                                                colTotalComp);
                    break;

                default:
                    add.ValuesToList(comprobante,
                                               comprobante.AlicuotaVarias,
                                               page,
                                               row,
                                               colNetoGravado,
                                               colIVA,
                                               colNoGravado,
                                               colOpExentas,
                                               colTotalComp);
                    break;
            }
            
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

