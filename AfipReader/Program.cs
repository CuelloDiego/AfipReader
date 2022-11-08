using Aspose.Cells;
using AfipReader.Core;

Workbook excel = new("C:\\Users\\diiee\\Desktop\\DDC.xlsx");

Worksheet page = excel.Worksheets[0];


AfipWorksheet afip = new AfipWorksheet();
var resultado = afip.GetDetails().Item1;
var rowsnotreaded = afip.GetDetails().Item2;
foreach (var comprobante in resultado)
{
    Console.WriteLine("Comprobante: " + comprobante.Nombretipocomp);
    Console.WriteLine("Neto Gravado 21%: "+comprobante.Alicuota21.Netogravado);
    Console.WriteLine("Monto IVA 21%: " + comprobante.Alicuota21.Iva);
    Console.WriteLine("Neto Gravado 10.5%: " + comprobante.Alicuota105.Netogravado);
    Console.WriteLine("Monto IVA 10.5%: " + comprobante.Alicuota105.Iva);
    Console.WriteLine("Neto Gravado 27%: " + comprobante.Alicuota27.Netogravado);
    Console.WriteLine("Monto IVA 27%:: " + comprobante.Alicuota27.Iva);
    Console.WriteLine("Neto Gravado Otras Alicuotas: " + comprobante.AlicuotaVarias.Netogravado);
    Console.WriteLine("Monto IVA Otras alicuotas : " + comprobante.AlicuotaVarias.Iva);
    Console.WriteLine("Neto NO Gravado: " + comprobante.NetoNogravado);
    Console.WriteLine("Operciones exentas: " + comprobante.OpExentas);
    Console.WriteLine("Total: " + comprobante.Total);
    Console.WriteLine("------------------------------------------------------ ");
}

Console.WriteLine("\nFilas no leidas ");
foreach (var row in rowsnotreaded)
{
    
    Console.WriteLine(row);
}


Console.WriteLine(); Console.WriteLine(); Console.WriteLine(); Console.WriteLine();



Console.ReadKey();

