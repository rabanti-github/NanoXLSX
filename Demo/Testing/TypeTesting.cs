/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2019
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NanoXLSX;

namespace Demo.Testing
{
    /// <summary>
    /// Class for testing of several data types
    /// </summary>
    public class TypeTesting
    {

        public static void NumericTypeTesting(string fileName)
        {
            Workbook wb = new Workbook(fileName, "NumericTest");
            Cell c;
            byte bVal = 11;
            sbyte sbVal = -22;
            decimal dcVal = 1.223m;
            double dVal = -999987.559712345689;
            float fVal = 4.000012f;
            int iVal = -187;
            uint uiVal = UInt32.MaxValue;
            long lVal = -99987;
            ulong ulVal = ulong.MaxValue;
            short sVal = -33;
            ushort usVal = 127;

            int i  = 0;
            c = new Cell(bVal, Cell.CellType.NUMBER);
            wb.CurrentWorksheet.AddCell(c, 0,i);
            i++;
            c = new Cell(sbVal, Cell.CellType.NUMBER);
            wb.CurrentWorksheet.AddCell(c, 0, i);
            i++;
            c = new Cell(dcVal, Cell.CellType.NUMBER);
            wb.CurrentWorksheet.AddCell(c, 0, i);
            i++;
            c = new Cell(dVal, Cell.CellType.NUMBER);
            wb.CurrentWorksheet.AddCell(c, 0, i);
            i++;
            c = new Cell(fVal, Cell.CellType.NUMBER);
            wb.CurrentWorksheet.AddCell(c, 0, i);
            i++;
            c = new Cell(iVal, Cell.CellType.NUMBER);
            wb.CurrentWorksheet.AddCell(c, 0, i);
            i++;
            c = new Cell(uiVal, Cell.CellType.NUMBER);
            wb.CurrentWorksheet.AddCell(c, 0, i);
            i++;
            c = new Cell(lVal, Cell.CellType.NUMBER);
            wb.CurrentWorksheet.AddCell(c, 0, i);
            i++;
            c = new Cell(ulVal, Cell.CellType.NUMBER);
            wb.CurrentWorksheet.AddCell(c, 0, i);
            i++;
            c = new Cell(sVal, Cell.CellType.NUMBER);
            wb.CurrentWorksheet.AddCell(c, 0, i);
            i++;
            c = new Cell(usVal, Cell.CellType.NUMBER);
            wb.CurrentWorksheet.AddCell(c, 0, i);
           
            wb.Save();
        }

    }
}
