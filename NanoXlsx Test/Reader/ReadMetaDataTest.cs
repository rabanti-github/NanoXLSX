using NanoXLSX;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace NanoXLSX_Test.Reader
{
    public class ReadMetaDataTest
    {

        [Fact(DisplayName = "Test of name property of worksheets when loading a workbook")]
        public void WorksheetNameTest()
        {
            Workbook workbook = new Workbook();
            workbook.AddWorksheet("test1");
            workbook.AddWorksheet("test2");
            workbook.AddWorksheet("test3");

            MemoryStream stream = new MemoryStream();
            workbook.SaveAsStream(stream, true);
            stream.Position = 0;
            Workbook givenWorkbook = Workbook.Load(stream);

            Assert.Equal(3, givenWorkbook.Worksheets.Count);
            Assert.Equal("test1", givenWorkbook.Worksheets[0].SheetName);
            Assert.Equal("test2", givenWorkbook.Worksheets[1].SheetName);
            Assert.Equal("test3", givenWorkbook.Worksheets[2].SheetName);
        }

        [Fact(DisplayName = "Test of hidden property of worksheets when loading a workbook")]
        public void WorksheetHiddenTest()
        {
            Workbook workbook = new Workbook();
            workbook.AddWorksheet("test1");
            workbook.AddWorksheet("test2");
            workbook.AddWorksheet("test3");
            workbook.SetSelectedWorksheet(1);
            workbook.Worksheets[0].Hidden = true;
            workbook.Worksheets[2].Hidden = true;

            MemoryStream stream = new MemoryStream();
            workbook.SaveAsStream(stream, true);
            stream.Position = 0;
            Workbook givenWorkbook = Workbook.Load(stream);

            Assert.Equal(3, givenWorkbook.Worksheets.Count);
            Assert.True(givenWorkbook.Worksheets[0].Hidden);
            Assert.False(givenWorkbook.Worksheets[1].Hidden);
            Assert.True(givenWorkbook.Worksheets[2].Hidden);
        }

    }


}
