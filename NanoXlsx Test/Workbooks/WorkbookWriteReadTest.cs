using NanoXLSX;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace NanoXLSX_Test.Workbooks
{
    public class WorkbookWriteReadTest
    {

        [Fact(DisplayName = "Test of the (virtual) 'MruColors' property when writing and reading a workbook")]
        public void ReadMruColorsTest()
        {
            Workbook workbook = new Workbook();
            string color1 = "AACC00";
            string color2 = "FFDD22";
            workbook.AddMruColor(color1);
            workbook.AddMruColor(color2);
            Workbook givenWorkbook = WriteAndReadWorkbook(workbook);
            List<string> mruColors = ((List<string>)givenWorkbook.GetMruColors());
            mruColors.Sort();
            Assert.Equal(2, mruColors.Count);
            Assert.Equal("FF" + color1, mruColors[0]);
            Assert.Equal("FF" + color2, mruColors[1]);
        }

        private static Workbook WriteAndReadWorkbook(Workbook workbook)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                workbook.SaveAsStream(stream, true);
                stream.Position = 0;
                Workbook readWorkbook = Workbook.Load(stream);
                return readWorkbook;
            }
        }

    }
}
