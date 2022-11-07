using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;


namespace AfipReader.Core
{
    public class TODOChange
    {
        public void ValuesToList(Comprobante x,Importes alicuota,Worksheet page, int row, int colNetoGravado, int colIVA,int colNoGravado,int colOpExentas, int colTotalComp)
        {
            alicuota.Netogravado += page.Cells[row, colNetoGravado].DoubleValue;
            alicuota.Iva += page.Cells[row, colIVA].DoubleValue;
            alicuota.Total = x.Alicuota21.Netogravado + x.Alicuota21.Iva;
            x.NetoNogravado += page.Cells[row, colNoGravado].DoubleValue;
            x.OpExentas += page.Cells[row, colOpExentas].DoubleValue;
            x.Total += page.Cells[row, colTotalComp].DoubleValue;
            
        }





    }











}
