﻿using System.Collections.Generic;
using NanoXLSX;
using NanoXLSX.Styles;
using NanoXLSX.Test.Writer_Reader.Utils;
using Xunit;

namespace NanoXLSX.Test.Writer_Reader.WorkbookTest
{
    public class WorkbookWriteReadTest
    {
        [Fact(DisplayName = "Test of the correct processing of 3 strings when saving and loading a workbook")]
        public void StringProcessingTest()
        {
            Workbook workbook = new Workbook("sheet1");
            workbook.CurrentWorksheet.AddCell("Text1", "A1");
            workbook.CurrentWorksheet.AddCell("Text2", "A2");
            workbook.CurrentWorksheet.AddCell("", "A3");
            workbook.CurrentWorksheet.AddCell(null, "A4");
            workbook.CurrentWorksheet.AddCell("Text1", "A5");
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.Equal(Cell.CellType.STRING, givenWorkbook.CurrentWorksheet.Cells["A1"].DataType);
            Assert.Equal("Text1", givenWorkbook.CurrentWorksheet.Cells["A1"].Value.ToString());
            Assert.Equal(Cell.CellType.STRING, givenWorkbook.CurrentWorksheet.Cells["A2"].DataType);
            Assert.Equal("Text2", givenWorkbook.CurrentWorksheet.Cells["A2"].Value.ToString());
            Assert.Equal(Cell.CellType.STRING, givenWorkbook.CurrentWorksheet.Cells["A3"].DataType);
            Assert.Equal("", givenWorkbook.CurrentWorksheet.Cells["A3"].Value.ToString());
            Assert.Equal(Cell.CellType.EMPTY, givenWorkbook.CurrentWorksheet.Cells["A4"].DataType);
            Assert.Null(givenWorkbook.CurrentWorksheet.Cells["A4"].Value);
            Assert.Equal(Cell.CellType.STRING, givenWorkbook.CurrentWorksheet.Cells["A5"].DataType);
            Assert.Equal("Text1", givenWorkbook.CurrentWorksheet.Cells["A5"].Value.ToString());
        }

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

        [Theory(DisplayName = "Test of the 'WorkbookProtectionPasswordHash' property when writing and reading a workbook, using legacy password")]
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
            string hash = LegacyPassword.GenerateLegacyPasswordHash(plainText);
            if (hash == "")
            {
                hash = null;
            }
            Assert.Equal(hash, givenWorkbook.WorkbookProtectionPassword.PasswordHash);
        }




    }
}
