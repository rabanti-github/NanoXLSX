using NanoXLSX.Exceptions;
using NanoXLSX.Styles;
using Xunit;
using FormatException = NanoXLSX.Exceptions.FormatException;

namespace NanoXLSX.Test.Core.WorkbookTest
{
    public class DefinedNameTest
    {
        #region constructor tests

        [Fact(DisplayName = "Test of the DefinedName constructor with all parameters")]
        public void ConstructorTest_AllParameters()
        {
            Workbook wb = new Workbook("Sheet1");
            DefinedName dn = new DefinedName("MyName", "Sheet1!$A$1", wb.CurrentWorksheet, "a comment");
            Assert.Equal("MyName", dn.Name);
            Assert.Equal("Sheet1!$A$1", dn.Reference);
            Assert.Same(wb.CurrentWorksheet, dn.LocalSheet);
            Assert.Equal("a comment", dn.Comment);
        }

        [Fact(DisplayName = "Test of the DefinedName constructor with default optional parameters")]
        public void ConstructorTest_Defaults()
        {
            DefinedName dn = new DefinedName("MyName", "Sheet1!$A$1");
            Assert.Equal("MyName", dn.Name);
            Assert.Equal("Sheet1!$A$1", dn.Reference);
            Assert.Null(dn.LocalSheet);
            Assert.Null(dn.Comment);
        }

        [Theory(DisplayName = "Test of the DefinedName constructor for invalid name (FormatException)")]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("1Foo")]
        [InlineData("9")]
        [InlineData("A1")] // Cell address is not allowed
        [InlineData("a1")] // Cell address is not allowed
        [InlineData("XFD1048576")] // Cell address is not allowed
        [InlineData("Z42")]
        public void ConstructorTest_InvalidName(string name)
        {
            Assert.Throws<FormatException>(() => new DefinedName(name, "Sheet1!$A$1"));
        }

        [Theory(DisplayName = "Test of the DefinedName constructor for valid (non cell-reference) names")]
        [InlineData("MyName")]
        [InlineData("LOGO1000")]
        [InlineData("Test")]
        [InlineData("_Foo")]
        [InlineData("XFE1")] // Out of range cell address - so, allowed as a name
        [InlineData("ABCD1")] // Out of range cell address - so, allowed as a name
        [InlineData("A1048577")] // Out of range cell address - so, allowed as a name
        [InlineData("A99999999999")] // Digit suffix exceeds int.MaxValue (overflow), so int.TryParse fails - allowed as a name
        [InlineData("XFD2147483648")] // Row number is one past int.MaxValue - allowed as a name
        public void ConstructorTest_ValidName(string name)
        {
            DefinedName dn = new DefinedName(name, "Sheet1!$A$1");
            Assert.Equal(name, dn.Name);
        }

        [Theory(DisplayName = "Test of the DefinedName constructor for invalid reference")]
        [InlineData(null)]
        [InlineData("")]
        public void ConstructorTest_InvalidReference(string reference)
        {
            Assert.Throws<FormatException>(() => new DefinedName("MyName", reference));
        }

        #endregion

        #region equality tests

        [Fact(DisplayName = "Test that DefinedName equals itself (reflexive)")]
        public void Equals_Reflexive()
        {
            DefinedName dn = new DefinedName("A", "Sheet1!$A$1");
            Assert.True(dn.Equals(dn));
            Assert.True(dn.Equals((object)dn));
        }

        [Fact(DisplayName = "Test that two equal DefinedName instances are Equals")]
        public void Equals_TwoEqualInstances()
        {
            Workbook wb = new Workbook("Sheet1");
            DefinedName a = new DefinedName("X", "Sheet1!$A$1", wb.CurrentWorksheet, "c");
            DefinedName b = new DefinedName("X", "Sheet1!$A$1", wb.CurrentWorksheet, "c");
            Assert.True(a.Equals(b));
            Assert.True(b.Equals(a));
            Assert.True(a.Equals((object)b));
            Assert.Equal(a.GetHashCode(), b.GetHashCode());
        }

        [Fact(DisplayName = "Test inequality on differing Name")]
        public void Equals_DiffersByName()
        {
            DefinedName a = new DefinedName("X", "Sheet1!$A$1");
            DefinedName b = new DefinedName("Y", "Sheet1!$A$1");
            Assert.False(a.Equals(b));
        }

        [Fact(DisplayName = "Test inequality on differing Reference")]
        public void Equals_DiffersByReference()
        {
            DefinedName a = new DefinedName("X", "Sheet1!$A$1");
            DefinedName b = new DefinedName("X", "Sheet1!$A$2");
            Assert.False(a.Equals(b));
        }

        [Fact(DisplayName = "Test inequality on differing Comment")]
        public void Equals_DiffersByComment()
        {
            DefinedName a = new DefinedName("X", "Sheet1!$A$1", null, "c1");
            DefinedName b = new DefinedName("X", "Sheet1!$A$1", null, "c2");
            Assert.False(a.Equals(b));
        }

        [Fact(DisplayName = "Test inequality on differing LocalSheet (by reference)")]
        public void Equals_DiffersByLocalSheet()
        {
            Workbook wb = new Workbook("Sheet1");
            wb.AddWorksheet("Sheet2");
            DefinedName a = new DefinedName("X", "Sheet1!$A$1", wb.Worksheets[0]);
            DefinedName b = new DefinedName("X", "Sheet1!$A$1", wb.Worksheets[1]);
            DefinedName c = new DefinedName("X", "Sheet1!$A$1", null);
            Assert.False(a.Equals(b));
            Assert.False(a.Equals(c));
        }

        [Fact(DisplayName = "Test that Equals returns false against null and different types")]
        public void Equals_NullAndOtherType()
        {
            DefinedName a = new DefinedName("X", "Sheet1!$A$1");
            Assert.False(a.Equals((DefinedName)null));
            Assert.False(a.Equals((object)null));
            Assert.False(a.Equals("not a defined name"));
        }

        #endregion

        #region CompareTo tests

        [Fact(DisplayName = "Test of CompareTo: by Name (ordinal)")]
        public void CompareTo_ByName()
        {
            DefinedName a = new DefinedName("Apple", "x");
            DefinedName b = new DefinedName("Banana", "x");
            Assert.True(a.CompareTo(b) < 0);
            Assert.True(b.CompareTo(a) > 0);
            Assert.Equal(0, a.CompareTo(new DefinedName("Apple", "x")));
        }

        [Fact(DisplayName = "Test of CompareTo: workbook scope sorts before worksheet scope")]
        public void CompareTo_ByScope()
        {
            Workbook wb = new Workbook("Sheet1");
            DefinedName workbookScope = new DefinedName("X", "y");
            DefinedName sheetScope = new DefinedName("X", "y", wb.CurrentWorksheet);
            Assert.True(workbookScope.CompareTo(sheetScope) < 0);
            Assert.True(sheetScope.CompareTo(workbookScope) > 0);
        }

        [Fact(DisplayName = "Test of CompareTo: worksheet scopes ordered by SheetID")]
        public void CompareTo_BySheetId()
        {
            Workbook wb = new Workbook("Sheet1");
            wb.AddWorksheet("Sheet2");
            DefinedName onSheet1 = new DefinedName("X", "y", wb.Worksheets[0]);
            DefinedName onSheet2 = new DefinedName("X", "y", wb.Worksheets[1]);
            Assert.True(onSheet1.CompareTo(onSheet2) < 0);
            Assert.True(onSheet2.CompareTo(onSheet1) > 0);
        }

        [Fact(DisplayName = "Test of CompareTo: by Reference when name and scope match")]
        public void CompareTo_ByReference()
        {
            DefinedName a = new DefinedName("X", "AA");
            DefinedName b = new DefinedName("X", "BB");
            Assert.True(a.CompareTo(b) < 0);
        }

        [Fact(DisplayName = "Test of CompareTo: by Comment as last tiebreaker")]
        public void CompareTo_ByComment()
        {
            DefinedName a = new DefinedName("X", "y", null, "alpha");
            DefinedName b = new DefinedName("X", "y", null, "beta");
            Assert.True(a.CompareTo(b) < 0);
        }

        [Fact(DisplayName = "Test of CompareTo: null comparand returns positive")]
        public void CompareTo_Null()
        {
            DefinedName a = new DefinedName("X", "y");
            Assert.True(a.CompareTo(null) > 0);
        }

        #endregion

        #region ToString tests

        [Fact(DisplayName = "Test of ToString includes name, scope and reference")]
        public void ToString_ContainsAllInfo()
        {
            Workbook wb = new Workbook("MySheet");
            DefinedName dn = new DefinedName("MyName", "MySheet!$A$1", wb.CurrentWorksheet);
            string s = dn.ToString();
            Assert.Contains("MyName", s);
            Assert.Contains("MySheet", s);
            Assert.Contains("MySheet!$A$1", s);
            DefinedName workbookScope = new DefinedName("Other", "ref");
            Assert.Contains("workbook", workbookScope.ToString());
        }

        #endregion

        #region Workbook API tests

        [Fact(DisplayName = "Test of Workbook.AddDefinedName / GetDefinedNames / GetDefinedName")]
        public void Workbook_AddAndGet()
        {
            Workbook wb = new Workbook("Sheet1");
            wb.AddDefinedName("MyName", "Sheet1!$A$1");
            Assert.Single(wb.GetDefinedNames());
            DefinedName dn = wb.GetDefinedName("MyName");
            Assert.NotNull(dn);
            Assert.Equal("Sheet1!$A$1", dn.Reference);
            Assert.Null(dn.LocalSheet);
        }

        [Fact(DisplayName = "Test of Workbook.AddDefinedName(DefinedName) overload")]
        public void Workbook_AddInstance()
        {
            Workbook wb = new Workbook("Sheet1");
            wb.AddDefinedName(new DefinedName("MyName", "Sheet1!$A$1"));
            Assert.Single(wb.GetDefinedNames());
        }

        [Fact(DisplayName = "Test that AddDefinedName(null) throws")]
        public void Workbook_AddNullThrows()
        {
            Workbook wb = new Workbook("Sheet1");
            Assert.Throws<WorksheetException>(() => wb.AddDefinedName((DefinedName)null));
        }

        [Fact(DisplayName = "Test that AddDefinedName with duplicate name and scope throws")]
        public void Workbook_AddDuplicateThrows()
        {
            Workbook wb = new Workbook("Sheet1");
            wb.AddDefinedName("MyName", "Sheet1!$A$1");
            Assert.Throws<WorksheetException>(() => wb.AddDefinedName("MyName", "Sheet1!$A$2"));
        }

        [Fact(DisplayName = "Test that AddDefinedName with same name but different scopes is allowed")]
        public void Workbook_AddSameNameDifferentScope()
        {
            Workbook wb = new Workbook("Sheet1");
            wb.AddWorksheet("Sheet2");
            wb.AddDefinedName("MyName", "Sheet1!$A$1");
            wb.AddDefinedName("MyName", "Sheet1!$A$2", wb.Worksheets[0]);
            wb.AddDefinedName("MyName", "Sheet1!$A$3", wb.Worksheets[1]);
            Assert.Equal(3, wb.GetDefinedNames().Count);
            Assert.NotNull(wb.GetDefinedName("MyName"));
            Assert.Null(wb.GetDefinedName("MyName").LocalSheet);
            Assert.Same(wb.Worksheets[0], wb.GetDefinedName("MyName", wb.Worksheets[0]).LocalSheet);
            Assert.Same(wb.Worksheets[1], wb.GetDefinedName("MyName", wb.Worksheets[1]).LocalSheet);
        }

        [Fact(DisplayName = "Test of GetDefinedName retrieving by scope")]
        public void Workbook_GetByScope()
        {
            Workbook wb = new Workbook("Sheet1");
            wb.AddWorksheet("Sheet2");
            wb.AddDefinedName("MyName", "wb-ref");
            wb.AddDefinedName("MyName", "sheet1-ref", wb.Worksheets[0]);
            wb.AddDefinedName("MyName", "sheet2-ref", wb.Worksheets[1]);
            Assert.Equal("wb-ref", wb.GetDefinedName("MyName").Reference);
            Assert.Equal("sheet1-ref", wb.GetDefinedName("MyName", wb.Worksheets[0]).Reference);
            Assert.Equal("sheet2-ref", wb.GetDefinedName("MyName", wb.Worksheets[1]).Reference);
        }

        [Fact(DisplayName = "Test that GetDefinedName returns null when not found")]
        public void Workbook_GetMissingReturnsNull()
        {
            Workbook wb = new Workbook("Sheet1");
            Assert.Null(wb.GetDefinedName("Unknown"));
        }

        [Fact(DisplayName = "Test of RemoveDefinedName: removes only matching scope")]
        public void Workbook_RemoveByScope()
        {
            Workbook wb = new Workbook("Sheet1");
            wb.AddDefinedName("MyName", "wb-ref");
            wb.AddDefinedName("MyName", "sheet-ref", wb.CurrentWorksheet);
            Assert.True(wb.RemoveDefinedName("MyName"));
            Assert.Single(wb.GetDefinedNames());
            Assert.NotNull(wb.GetDefinedName("MyName", wb.CurrentWorksheet));
            Assert.Null(wb.GetDefinedName("MyName"));
        }

        [Fact(DisplayName = "Test that RemoveDefinedName returns false for missing entry")]
        public void Workbook_RemoveMissingReturnsFalse()
        {
            Workbook wb = new Workbook("Sheet1");
            Assert.False(wb.RemoveDefinedName("Unknown"));
        }

        #endregion

        #region Worksheet API tests

        [Fact(DisplayName = "Test of Worksheet.AddDefinedName creates a worksheet-scoped name")]
        public void Worksheet_AddDefinedName()
        {
            Workbook wb = new Workbook("Sheet1");
            wb.CurrentWorksheet.AddDefinedName("MyName", "Sheet1!$A$1");
            DefinedName dn = wb.GetDefinedName("MyName", wb.CurrentWorksheet);
            Assert.NotNull(dn);
            Assert.Same(wb.CurrentWorksheet, dn.LocalSheet);
        }

        [Fact(DisplayName = "Test that Worksheet.AddDefinedName on a detached worksheet throws")]
        public void Worksheet_AddDefinedName_Detached()
        {
            Worksheet ws = new Worksheet("orphan");
            Assert.Throws<WorksheetException>(() => ws.AddDefinedName("MyName", "Sheet1!$A$1"));
        }

        [Fact(DisplayName = "Test of Worksheet.RemoveDefinedName removes only worksheet scope")]
        public void Worksheet_RemoveDefinedName()
        {
            Workbook wb = new Workbook("Sheet1");
            wb.AddDefinedName("MyName", "wb-ref");
            wb.CurrentWorksheet.AddDefinedName("MyName", "sheet-ref");
            Assert.True(wb.CurrentWorksheet.RemoveDefinedName("MyName"));
            Assert.NotNull(wb.GetDefinedName("MyName"));
            Assert.Null(wb.GetDefinedName("MyName", wb.CurrentWorksheet));
        }

        [Fact(DisplayName = "Test that Worksheet.RemoveDefinedName on detached worksheet throws")]
        public void Worksheet_RemoveDefinedName_Detached()
        {
            Worksheet ws = new Worksheet("orphan");
            Assert.Throws<WorksheetException>(() => ws.RemoveDefinedName("X"));
        }

        [Fact(DisplayName = "Test that Worksheet.GetDefinedName returns only worksheet scope")]
        public void Worksheet_GetDefinedName()
        {
            Workbook wb = new Workbook("Sheet1");
            wb.AddDefinedName("MyName", "wb-ref");
            Assert.Null(wb.CurrentWorksheet.GetDefinedName("MyName"));
            wb.CurrentWorksheet.AddDefinedName("MyName", "sheet-ref");
            Assert.Equal("sheet-ref", wb.CurrentWorksheet.GetDefinedName("MyName").Reference);
        }

        [Fact(DisplayName = "Test that Worksheet.GetDefinedName on detached worksheet throws")]
        public void Worksheet_GetDefinedName_Detached()
        {
            Worksheet ws = new Worksheet("orphan");
            Assert.Throws<WorksheetException>(() => ws.GetDefinedName("X"));
        }

        #endregion

        #region AddCellReference tests

        [Fact(DisplayName = "Test of AddCellReference(DefinedName, address) creates a Reference cell")]
        public void Worksheet_AddCellReference_StringAddress()
        {
            Workbook wb = new Workbook("Sheet1");
            DefinedName dn = new DefinedName("MyName", "Sheet1!$A$1");
            wb.AddDefinedName(dn);
            wb.CurrentWorksheet.AddCellReference(dn, "B2");
            Cell c = wb.CurrentWorksheet.Cells["B2"];
            Assert.Equal(Cell.CellType.Reference, c.DataType);
            Assert.Equal("MyName", c.Value);
        }

        [Fact(DisplayName = "Test of AddCellReference(DefinedName, address) creates a Reference cell with a style")]
        public void Worksheet_AddCellReference_StringAddress_WithStyle()
        {
            Workbook wb = new Workbook("Sheet1");
            DefinedName dn = new DefinedName("MyName", "Sheet1!$A$1");
            wb.AddDefinedName(dn);
            wb.CurrentWorksheet.AddCellReference(dn, "B2", (Style)BasicStyles.Bold.Copy());
            Cell c = wb.CurrentWorksheet.Cells["B2"];
            Assert.Equal(Cell.CellType.Reference, c.DataType);
            Assert.Equal("MyName", c.Value);
            Assert.Equal(BasicStyles.Bold.GetHashCode(), c.CellStyle.GetHashCode());
        }

        [Fact(DisplayName = "Test of AddCellReference(DefinedName, col, row) creates a Reference cell")]
        public void Worksheet_AddCellReference_ColRow()
        {
            Workbook wb = new Workbook("Sheet1");
            DefinedName dn = new DefinedName("MyName", "Sheet1!$A$1");
            wb.AddDefinedName(dn);
            wb.CurrentWorksheet.AddCellReference(dn, 1, 1);
            Cell c = wb.CurrentWorksheet.Cells["B2"];
            Assert.Equal(Cell.CellType.Reference, c.DataType);
            Assert.Equal("MyName", c.Value);
        }
        [Fact(DisplayName = "Test of AddCellReference(DefinedName, col, row) creates a Reference cell with a style")]
        public void Worksheet_AddCellReference_ColRow_WithStyle()
        {
            Workbook wb = new Workbook("Sheet1");
            DefinedName dn = new DefinedName("MyName", "Sheet1!$A$1");
            wb.AddDefinedName(dn);
            wb.CurrentWorksheet.AddCellReference(dn, 1, 1, (Style)BasicStyles.Italic.Copy());
            Cell c = wb.CurrentWorksheet.Cells["B2"];
            Assert.Equal(Cell.CellType.Reference, c.DataType);
            Assert.Equal("MyName", c.Value);
            Assert.Equal(BasicStyles.Italic.GetHashCode(), c.CellStyle.GetHashCode());
        }

        [Fact(DisplayName = "Test that AddCellReference(null) throws WorksheetException")]
        public void Worksheet_AddCellReference_NullThrows()
        {
            Workbook wb = new Workbook("Sheet1");
            Assert.Throws<WorksheetException>(() => wb.CurrentWorksheet.AddCellReference(null, "A1"));
        }

        #endregion

        #region misc

        [Fact(DisplayName = "Test that GetDefinedNames returns insertion order")]
        public void Workbook_InsertionOrderPreserved()
        {
            Workbook wb = new Workbook("Sheet1");
            wb.AddDefinedName("Beta", "x");
            wb.AddDefinedName("Alpha", "y");
            wb.AddDefinedName("Gamma", "z");
            Assert.Equal(new[] { "Beta", "Alpha", "Gamma" },
                new[] { wb.GetDefinedNames()[0].Name, wb.GetDefinedNames()[1].Name, wb.GetDefinedNames()[2].Name });
        }

        #endregion
    }
}
