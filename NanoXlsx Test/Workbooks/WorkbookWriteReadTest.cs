using NanoXLSX;
using NanoXLSX.Styles;
using System.Collections.Generic;
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
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            List<string> mruColors = ((List<string>)givenWorkbook.GetMruColors());
            mruColors.Sort();
            Assert.Equal(2, mruColors.Count);
            Assert.Equal("FF" + color1, mruColors[0]);
            Assert.Equal("FF" + color2, mruColors[1]);
        }

        [Fact(DisplayName = "Test of the (virtual) 'MruColors' property when writing and reading a workbook, neglecting the default color")]
        public void ReadMruColorsTest2()
        {
            Workbook workbook = new Workbook();
            string color1 = "AACC00";
            string color2 = Fill.DEFAULT_COLOR; // Should not be added (black)
            workbook.AddMruColor(color1);
            workbook.AddMruColor(color2);
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            List<string> mruColors = ((List<string>)givenWorkbook.GetMruColors());
            mruColors.Sort();
            Assert.Single(mruColors);
            Assert.Equal("FF" + color1, mruColors[0]);
        }

        [Theory(DisplayName = "Test of the 'Hidden' property when writing and reading a workbook")]
        [InlineData(true)]
        [InlineData(false)]
        public void ReadWorkbookHiddenTest(bool hidden)
        {
            Workbook workbook = new Workbook();
            workbook.Hidden = hidden;
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.Equal(hidden, givenWorkbook.Hidden);
        }

        [Theory(DisplayName = "Test of the 'SelectedWorksheet' property when writing and reading a workbook")]
        [InlineData(0)]
        [InlineData(1)]
        [InlineData(2)]
        public void ReadWorkbookSelectedWorksheetTest(int index)
        {
            Workbook workbook = new Workbook("sheet1");
            workbook.AddWorksheet("sheet2");
            workbook.AddWorksheet("sheet3");
            workbook.AddWorksheet("sheet4");
            workbook.SetSelectedWorksheet(index);
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.Equal(index, givenWorkbook.SelectedWorksheet);
        }

        [Theory(DisplayName = "Test of the 'LockWindowsIfProtected' property when writing and reading a workbook")]
        [InlineData(true)]
        [InlineData(false)]
        public void ReadWorkbookLockWindowsTest(bool locked)
        {
            Workbook workbook = new Workbook("sheet1");
            workbook.SetWorkbookProtection(true, locked, false, null);
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.Equal(locked, givenWorkbook.LockWindowsIfProtected);
        }

        [Theory(DisplayName = "Test of the 'LockStructureIfProtected' property when writing and reading a workbook")]
        [InlineData(true)]
        [InlineData(false)]
        public void ReadWorkbookLockStructureTest(bool locked)
        {
            Workbook workbook = new Workbook("sheet1");
            workbook.SetWorkbookProtection(true, false, locked, null);
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.Equal(locked, givenWorkbook.LockStructureIfProtected);
        }

        [Theory(DisplayName = "Test of the 'WorkbookProtectionPasswordHash' property when writing and reading a workbook")]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("A")]
        [InlineData("123")]
        [InlineData("test")]
        public void ReadWorkbookPasswordHashTest(string plainText)
        {
            Workbook workbook = new Workbook("sheet1");
            workbook.SetWorkbookProtection(true, false, true, plainText);
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            string hash = Utils.GeneratePasswordHash(plainText);
            if (hash == "")
            {
                hash = null;
            }
            Assert.Equal(hash, givenWorkbook.WorkbookProtectionPasswordHash);
        }




    }
}
