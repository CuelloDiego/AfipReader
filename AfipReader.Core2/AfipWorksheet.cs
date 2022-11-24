using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;


namespace AfipReader.Core
{
    public class AfipWorksheet
    {

        public (IEnumerable<Comprobante>, IEnumerable<int>) GetDetails(Workbook excel)
        {


            Worksheet page = excel.Worksheets[0];

            int colTipoComp = 1;
            int colNetoGravado = 11;
            int colNoGravado = 12;
            int colOpExentas = 13;
            int colIVA = 14;
            int colTotalComp = 15;
            int index;

            List<int> RowsNotReaded = new List<int>();

            List<Comprobante> DetalleComprobantes = new List<Comprobante>();

            for (int row = 0; row <= page.Cells.MaxDataRow; row++)
            {


                //Verificar variables o saltar ejecucion
                if (!VerifyCellsValues(row, colNetoGravado, colNoGravado, colOpExentas, colIVA, colTotalComp, RowsNotReaded, page))
                {
                    continue;
                }



                //agregar tipo comprobante y devolver index
                index = GetIndex(DetalleComprobantes, page, row, colTipoComp);



                //Falta agregar redondeo
                switch (page.Cells[row, colIVA].DoubleValue / page.Cells[row, colNetoGravado].DoubleValue)
                {
                    case 0.21:

                        AddValuesToList(DetalleComprobantes[index],
                                                   DetalleComprobantes[index].Alicuota21,
                                                   page,
                                                   row,
                                                   colNetoGravado,
                                                   colIVA,
                                                   colNoGravado,
                                                   colOpExentas,
                                                   colTotalComp);

                        break;
                    case 0.27:
                        AddValuesToList(DetalleComprobantes[index],
                                                   DetalleComprobantes[index].Alicuota27,
                                                   page,
                                                   row,
                                                   colNetoGravado,
                                                   colIVA,
                                                   colNoGravado,
                                                   colOpExentas,
                                                   colTotalComp);

                        break;
                    case 0.105:
                        AddValuesToList(DetalleComprobantes[index],
                                                    DetalleComprobantes[index].Alicuota105,
                                                    page,
                                                    row,
                                                    colNetoGravado,
                                                    colIVA,
                                                    colNoGravado,
                                                    colOpExentas,
                                                    colTotalComp);
                        break;

                    default:
                        AddValuesToList(DetalleComprobantes[index],
                                                   DetalleComprobantes[index].AlicuotaVarias,
                                                   page,
                                                   row,
                                                   colNetoGravado,
                                                   colIVA,
                                                   colNoGravado,
                                                   colOpExentas,
                                                   colTotalComp);
                        break;
                }




                //Console.Write(page.Cells[row, col].Value + " | ");






            }



            return (DetalleComprobantes, RowsNotReaded);



        }





        public void AddValuesToList(Comprobante x, Importes alicuota, Worksheet page, int row, int colNetoGravado, int colIVA, int colNoGravado, int colOpExentas, int colTotalComp)
        {
            alicuota.Netogravado += page.Cells[row, colNetoGravado].DoubleValue;
            alicuota.Iva += page.Cells[row, colIVA].DoubleValue;
            alicuota.Total = x.Alicuota21.Netogravado + x.Alicuota21.Iva;
            x.NetoNogravado += page.Cells[row, colNoGravado].DoubleValue;
            x.OpExentas += page.Cells[row, colOpExentas].DoubleValue;
            x.Total += page.Cells[row, colTotalComp].DoubleValue;

        }



        public bool VerifyCellsValues(int row, int colNetoGravado, int colNoGravado, int colOpExentas, int colIVA, int colTotalComp, List<int> RowNotReaded, Worksheet page)
        {
            double aux;
            if (double.TryParse(page.Cells[row, colNetoGravado].StringValue, out aux) &
                double.TryParse(page.Cells[row, colNoGravado].StringValue, out aux) &
                double.TryParse(page.Cells[row, colOpExentas].StringValue, out aux) &
                double.TryParse(page.Cells[row, colIVA].StringValue, out aux) &
                double.TryParse(page.Cells[row, colTotalComp].StringValue, out aux))
            {
                return true;
            }




            RowNotReaded.Add(row);
            return false;

        }

        public int GetIndex(List<Comprobante> DetalleComprobantes, Worksheet page, int row, int colTipoComp)
        {
            //agregar comprobante y devolver index

            if (!DetalleComprobantes.Exists(x => x.Nombretipocomp == page.Cells[row, colTipoComp].StringValue))
            {
                DetalleComprobantes.Add(new Comprobante() { Nombretipocomp = page.Cells[row, colTipoComp].StringValue });
            }

            return DetalleComprobantes.IndexOf(DetalleComprobantes.First(x => x.Nombretipocomp == page.Cells[row, colTipoComp].StringValue));




        }







    }











}
