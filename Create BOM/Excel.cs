using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace ConsoleApplication1
{
    class Excel
    {
        Application xlApp = new Application();
        Workbook wb;
        Worksheet ws;


        int numLine = 1;
        private void _setCell(String cell, String val)
        {
            Range aRange = ws.get_Range(cell, cell);
            aRange.Select();
            String value = aRange.Value;
            if ((value != null )&&(value.Length > 0) )
                value += ", ";
            aRange.Value = value + val;

        }

        public void create()
        {
            //xlApp.Visible = true;

            wb = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            ws = (Worksheet)wb.Worksheets[1];

            numLine = 1;
            _setCell("A" + numLine.ToString(), "Référence Client");
            _setCell("B" + numLine.ToString(), "Valeur");
            _setCell("C" + numLine.ToString(), "Référence Digi-Key");
            //_setCell("D" + numLine.ToString(), "Référence Fabriquant");
            _setCell("D" + numLine.ToString(), "Quantité 1");
        }

        public void newLine()
        {
            numLine++;
        }

        public void setRefClient(String Val)
        {
            _setCell("A" + numLine.ToString(), Val);
        }

        public void setValue(String Val)
        {
            _setCell("B" + numLine.ToString(), Val);
        }

        public void setRefDigikey(String Val)
        {
            _setCell("C" + numLine.ToString(), Val);
        }

        public void setRefFabriquant(String Val)
        {
            //_setCell("D" + numLine.ToString(), Val);
        }

        public void setQuantity(String Val)
        {
            _setCell("D" + numLine.ToString(), Val);
        }

        public void close()
        {
            xlApp.Quit();
        }
    }
}
