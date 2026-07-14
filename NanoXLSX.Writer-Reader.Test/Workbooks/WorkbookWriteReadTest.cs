using System.Collections.Generic;
using NanoXLSX.Colors;
using NanoXLSX.Extensions;
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
            Assert.Equal(Cell.CellType.String, givenWorkbook.CurrentWorksheet.Cells["A1"].DataType);
            Assert.Equal("Text1", givenWorkbook.CurrentWorksheet.Cells["A1"].Value.ToString());
            Assert.Equal(Cell.CellType.String, givenWorkbook.CurrentWorksheet.Cells["A2"].DataType);
            Assert.Equal("Text2", givenWorkbook.CurrentWorksheet.Cells["A2"].Value.ToString());
            Assert.Equal(Cell.CellType.String, givenWorkbook.CurrentWorksheet.Cells["A3"].DataType);
            Assert.Equal("", givenWorkbook.CurrentWorksheet.Cells["A3"].Value.ToString());
            Assert.Equal(Cell.CellType.Empty, givenWorkbook.CurrentWorksheet.Cells["A4"].DataType);
            Assert.Null(givenWorkbook.CurrentWorksheet.Cells["A4"].Value);
            Assert.Equal(Cell.CellType.String, givenWorkbook.CurrentWorksheet.Cells["A5"].DataType);
            Assert.Equal("Text1", givenWorkbook.CurrentWorksheet.Cells["A5"].Value.ToString());
        }

        [Fact(DisplayName = "Test of the (virtual) 'MruColors' property on a ARGB value, when writing and reading a workbook")]
        public void ReadMruColorsTest()
        {
            Workbook workbook = new Workbook();
            string color1 = "AACC00";
            string color2 = "FFDD22";
            workbook.AddMruColor(color1);
            workbook.AddMruColor(color2);
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            List<Color> mruColors = ((List<Color>)givenWorkbook.GetMruColors());
            mruColors.Sort();
            Assert.Equal(2, mruColors.Count);
            Assert.Equal("FF" + color1, mruColors[0].GetArgbValue());
            Assert.Equal("FF" + color2, mruColors[1].GetArgbValue());
        }


        [Fact(DisplayName = "Test of the (virtual) 'MruColors' property on a indexed color, when writing and reading a workbook")]
        public void ReadMruColorsTest2()
        {
            Workbook workbook = new Workbook();
            workbook.AddMruColor(IndexedColor.Value.Blue4);
            workbook.AddMruColor(IndexedColor.Value.StrongYellow);
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            List<Color> mruColors = ((List<Color>)givenWorkbook.GetMruColors());
            mruColors.Sort();
            Assert.Equal(2, mruColors.Count);
            Assert.Equal(IndexedColor.GetArgbValue(IndexedColor.Value.Blue4), mruColors[0].GetArgbValue());
            Assert.Equal(IndexedColor.GetArgbValue(IndexedColor.Value.StrongYellow), mruColors[1].GetArgbValue());
        }


        [Fact(DisplayName = "Test of the (virtual) 'MruColors' property when writing and reading a workbook, neglecting the default color")]
        public void ReadMruColorsTest3()
        {
            Workbook workbook = new Workbook();
            string color1 = "AACC00";
            string color2 = Fill.DefaultColor.RgbColor.ColorValue; // Should not be added (black / default color)
            workbook.AddMruColor(color1);
            workbook.AddMruColor(color2);
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            List<Color> mruColors = ((List<Color>)givenWorkbook.GetMruColors());
            mruColors.Sort();
            Assert.Single(mruColors);
            Assert.Equal("FF" + color1, mruColors[0].GetArgbValue());
        }

        [Fact(DisplayName = "Test of the (virtual) 'MruColors' property when writing and reading a workbook, neglecting an undefined color")]
        public void ReadMruColorsTest4()
        {
            Workbook workbook = new Workbook();
            Color color = Color.CreateNone();
            workbook.AddMruColor(color);
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            List<Color> mruColors = ((List<Color>)givenWorkbook.GetMruColors());
            Assert.Empty(mruColors);
        }

        [Theory(DisplayName = "Test of the 'Hidden' property when writing and reading a workbook")]
        [InlineData(true)]
        [InlineData(false)]
        public void ReadWorkbookHiddenTest(bool hidden)
        {
            Workbook workbook = new Workbook
            {
                Hidden = hidden
            };
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

        [Fact(DisplayName = "Test that a workbook without defined names produces no defined names after round-trip")]
        public void DefinedNames_EmptyRoundTrip()
        {
            Workbook workbook = new Workbook("sheet1");
            workbook.CurrentWorksheet.AddCell(1, "A1");
            Workbook given = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.Empty(given.GetDefinedNames());
        }

        [Fact(DisplayName = "Test of a workbook-scoped defined name with a single cell reference")]
        public void DefinedNames_WorkbookScope_CellReference()
        {
            Workbook workbook = new Workbook("sheet1");
            workbook.CurrentWorksheet.AddCell(42, "A1");
            workbook.AddDefinedName("MyCell", "sheet1!$A$1");
            Workbook given = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.Single(given.GetDefinedNames());
            DefinedName dn = given.GetDefinedName("MyCell");
            Assert.NotNull(dn);
            Assert.Equal("sheet1!$A$1", dn.Reference);
            Assert.Null(dn.LocalSheet);
            Assert.Null(dn.Comment);
        }

        [Fact(DisplayName = "Test of a workbook-scoped defined name with a range reference")]
        public void DefinedNames_WorkbookScope_RangeReference()
        {
            Workbook workbook = new Workbook("sheet1");
            workbook.CurrentWorksheet.AddCell(1, "A1");
            workbook.CurrentWorksheet.AddCell(2, "A2");
            workbook.CurrentWorksheet.AddCell(3, "A3");
            workbook.AddDefinedName("MyRange", "sheet1!$A$1:$A$3");
            Workbook given = TestUtils.WriteAndReadWorkbook(workbook);
            DefinedName dn = given.GetDefinedName("MyRange");
            Assert.NotNull(dn);
            Assert.Equal("sheet1!$A$1:$A$3", dn.Reference);
        }

        [Fact(DisplayName = "Test of a workbook-scoped defined name holding a formula expression")]
        public void DefinedNames_WorkbookScope_FormulaReference()
        {
            Workbook workbook = new Workbook("sheet1");
            workbook.AddDefinedName("MySum", "SUM(sheet1!$A$1:$A$3)");
            Workbook given = TestUtils.WriteAndReadWorkbook(workbook);
            DefinedName dn = given.GetDefinedName("MySum");
            Assert.NotNull(dn);
            Assert.Equal("SUM(sheet1!$A$1:$A$3)", dn.Reference);
        }

        [Fact(DisplayName = "Test of a worksheet-scoped defined name (localSheetId) round-trip")]
        public void DefinedNames_WorksheetScope()
        {
            Workbook workbook = new Workbook("sheet1");
            workbook.AddWorksheet("sheet2");
            workbook.AddDefinedName("LocalName", "sheet2!$B$2", workbook.Worksheets[1]);
            Workbook given = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.Single(given.GetDefinedNames());
            Worksheet sheet2 = given.GetWorksheet("sheet2");
            DefinedName dn = given.GetDefinedName("LocalName", sheet2);
            Assert.NotNull(dn);
            Assert.Equal("sheet2!$B$2", dn.Reference);
            Assert.Same(sheet2, dn.LocalSheet);
        }

        [Fact(DisplayName = "Test of multiple defined names preserving insertion order on round-trip")]
        public void DefinedNames_OrderPreserved()
        {
            Workbook workbook = new Workbook("sheet1");
            workbook.AddDefinedName("Beta", "sheet1!$A$1");
            workbook.AddDefinedName("Alpha", "sheet1!$A$2");
            workbook.AddDefinedName("Gamma", "sheet1!$A$3");
            Workbook given = TestUtils.WriteAndReadWorkbook(workbook);
            System.Collections.Generic.IReadOnlyList<DefinedName> names = given.GetDefinedNames();
            Assert.Equal(3, names.Count);
            Assert.Equal("Beta", names[0].Name);
            Assert.Equal("Alpha", names[1].Name);
            Assert.Equal("Gamma", names[2].Name);
        }

        [Fact(DisplayName = "Test of the comment attribute round-trip on a defined name")]
        public void DefinedNames_CommentRoundTrip()
        {
            Workbook workbook = new Workbook("sheet1");
            workbook.AddDefinedName("MyName", "sheet1!$A$1", null, "this is a comment");
            Workbook given = TestUtils.WriteAndReadWorkbook(workbook);
            DefinedName dn = given.GetDefinedName("MyName");
            Assert.NotNull(dn);
            Assert.Equal("this is a comment", dn.Comment);
        }

        [Fact(DisplayName = "Test that ReadDefinedNameReference returns empty string for a self-closing definedName element")]
        public void DefinedNames_ReadReference_SelfClosing()
        {
            const string xml = "<definedName name=\"X\"/>";
            using (System.IO.StringReader sr = new System.IO.StringReader(xml))
            using (System.Xml.XmlReader reader = System.Xml.XmlReader.Create(sr))
            {
                reader.MoveToContent();
                string result = NanoXLSX.Internal.Readers.WorkbookReader.ReadDefinedNameReference(reader);
                Assert.Equal(string.Empty, result);
            }
        }

        [Fact(DisplayName = "Test that ReadDefinedNameReference returns the text content for a non-empty definedName element")]
        public void DefinedNames_ReadReference_TextContent()
        {
            const string xml = "<definedName name=\"X\">sheet1!$A$1</definedName>";
            using (System.IO.StringReader sr = new System.IO.StringReader(xml))
            using (System.Xml.XmlReader reader = System.Xml.XmlReader.Create(sr))
            {
                reader.MoveToContent();
                string result = NanoXLSX.Internal.Readers.WorkbookReader.ReadDefinedNameReference(reader);
                Assert.Equal("sheet1!$A$1", result);
            }
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
