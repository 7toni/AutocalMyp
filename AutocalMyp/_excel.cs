using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace AutocalMyp
{
    public class _excel
    {
        _Application excel = new Application();
        Workbook wb;
        Worksheet ws;
        Range range;

        public _excel() { }

        public void Open(string nombre)
        {
            string path = @"C:\datamyp\" + nombre + ".xlsx";
            wb = excel.Workbooks.Open(path);
            ws = (Worksheet)wb.Worksheets.get_Item(1);
            range = ws.UsedRange;
            //set columns format to text format
            //ws = (Worksheet)wb.Worksheets.Add();
            ws.Columns.NumberFormat = "@";
        }

        public void CreateExel(string nombre)
        {
            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook worKbooK;
            Microsoft.Office.Interop.Excel.Worksheet worKsheeT;


            if (!System.IO.File.Exists(@"C:\datamyp\" + nombre + ".xlsx"))
            {
                xlApp = new Microsoft.Office.Interop.Excel.Application();
                worKbooK = xlApp.Workbooks.Add(Type.Missing);
                worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;

                worKbooK.SaveAs(@"C:\datamyp\" + nombre + ".xlsx");
                worKbooK.Close();
                xlApp.Quit();

                Marshal.ReleaseComObject(worKbooK);
                Marshal.ReleaseComObject(worKsheeT);
                Marshal.ReleaseComObject(xlApp);
            }

        }

        public string ReadCell()
        {
            string valor = "";
            string str;
            int rCnt;
            int cCnt;
            int rw = 0;
            int cl = 0;

            rw = range.Rows.Count;
            cl = range.Columns.Count;

            for (rCnt = 1; rCnt <= rw; rCnt++)
            {
                for (cCnt = 1; cCnt <= cl; cCnt++)
                {
                    str = (string)(range.Cells[rCnt, cCnt] as Range).Value2;
                    valor = str;
                }
            }

            return valor;
        }

        public int ScanCell()
        {
            int valor = 0;
            valor = range.Rows.Count;
            return valor;
        }

        public void WriteFuncion(int i, string funcion)
        {
            // [ Fila , Columna ]                      
            ws.Cells[i, 1] = funcion;

        }

        public void WriteRango(int i, string rango)
        {
            // [ Fila , Columna ] 
            ws.Cells[i, 1] = rango;
        }

        public void WriteValorLectura(int i, string valorpatron, string lecturas)
        {
            // [ Fila , Columna ]                                           
            ws.Cells[i, 1] = valorpatron;
            ws.Cells[i, 2] = lecturas;
        }

        public void Save()
        {
            wb.Save();
        }

        public void SaveAs(string path)
        {
            wb.SaveAs(path);
        }

        public void Close()
        {
            wb.Close();
        }

        public void Quit()
        {
            excel.Quit();
            Marshal.ReleaseComObject(ws);
            Marshal.ReleaseComObject(wb);
            Marshal.ReleaseComObject(excel);
        }
    }
}
