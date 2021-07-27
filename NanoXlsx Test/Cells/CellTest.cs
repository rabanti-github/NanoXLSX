using NanoXLSX;
using NanoXLSX.Exceptions;
using NanoXLSX.Styles;
using NanoXLSX_Test.Cells.Types;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;
using static NanoXLSX.Cell;

namespace NanoXLSX_Test.Cells
{
    // Ensure that these tests are executed sequentially, since static repository methods are called 
    [Collection(nameof(SequentialCollection))]
    public class CellTest
    {
        private Cell cell;
        private Type cellType;
        private object cellValue;
        private Address cellAddress;
        private CellTypeUtils utils;

        public CellTest()
        {

            this.utils = new CellTypeUtils();
            this.cellType = typeof(string);
            this.cellValue = "value";
            this.cellAddress = this.utils.CellAddress;
            this.cell = this.utils.CreateVariantCell(this.cellValue as string, this.cellAddress, true, BasicStyles.BoldItalic);
        }

        [Theory(DisplayName = "Test of the set function of the CellAdress property")]
        [InlineData("A1", 0, 0, AddressType.Default)]
        [InlineData("J100", 9, 99, AddressType.Default)]
        [InlineData("$B2", 1, 1, AddressType.FixedColumn)]
        [InlineData("B$2", 1, 1, AddressType.FixedRow)]
        [InlineData("$B$2", 1, 1, AddressType.FixedRowAndColumn)]
        public void SetAdressStringPropertyTest(string givenAddress, int expectedColumn, int expectedRow, Cell.AddressType expectedType)
        {
            this.cell.CellAddress = givenAddress;
            Assert.Equal(this.cell.CellAddress2.Column, expectedColumn);
            Assert.Equal(this.cell.CellAddress2.Row, expectedRow);
            Assert.Equal(this.cell.CellAddress2.Type, expectedType);
        }

        [Theory(DisplayName = "Test of the get function of the CellAdressString property")]
        [InlineData(0, 0, AddressType.Default, "A1")]
        [InlineData(9, 99, AddressType.Default, "J100")]
        [InlineData(1, 1, AddressType.FixedColumn, "$B2")]
        [InlineData(1, 1, AddressType.FixedRow, "B$2")]
        [InlineData(1, 1, AddressType.FixedRowAndColumn, "$B$2")]
        public void GetAddressStringPropertyTest(int givendColumn, int givenRow, Cell.AddressType givenTyp, string expectedAddress)
        {
            this.cell.CellAddressType = givenTyp;
            this.cell.ColumnNumber = givendColumn;
            this.cell.RowNumber = givenRow;
            Assert.Equal(this.cell.CellAddress, expectedAddress);
        }

        [Theory(DisplayName = "Test of the set function of the CellAdress property, as well as get functions of columnNumber, RowNumber and AddressType")]
        [InlineData("A1", 0, 0, AddressType.Default)]
        [InlineData("XFD1048576", 16383, 1048575, AddressType.Default)]
        [InlineData("$B2", 1, 1, AddressType.FixedColumn)]
        [InlineData("B$2", 1, 1, AddressType.FixedRow)]
        [InlineData("$B$2", 1, 1, AddressType.FixedRowAndColumn)]
        public void SetAdressPropertyTest(string givenAddress, int expectedColumn, int expectedRow, Cell.AddressType expectedType)
        {
            Address given = new Address(givenAddress);
            this.cell.CellAddress2 = given;
            Assert.Equal(this.cell.ColumnNumber, expectedColumn);
            Assert.Equal(this.cell.RowNumber, expectedRow);
            Assert.Equal(this.cell.CellAddressType, expectedType);
        }

        [Theory(DisplayName = "Test of the get function of the CellAdress property, as well as set functions of columnNumber, RowNumber and AddressType")]
        [InlineData(0, 0, AddressType.Default, "A1")]
        [InlineData(16383, 1048575, AddressType.Default, "XFD1048576")]
        [InlineData(1, 1, AddressType.FixedColumn, "$B2")]
        [InlineData(1, 1, AddressType.FixedRow, "B$2")]
        [InlineData(1, 1, AddressType.FixedRowAndColumn, "$B$2")]
        public void GetAdressPropertyTest(int givendColumn, int givenRow, Cell.AddressType givenTyp, string expectedAddress)
        {
            this.cell.ColumnNumber = givendColumn;
            this.cell.RowNumber = givenRow;
            this.cell.CellAddressType = givenTyp;
            Address expected = new Address(expectedAddress);
            Assert.Equal(this.cell.CellAddress2, expected);
        }

        [Theory(DisplayName = "Test of the address scope check function")]
        [InlineData("A1", AddressScope.SingleAddress)]
        [InlineData("$A$1", AddressScope.SingleAddress)]
        [InlineData("A1:B2", AddressScope.Range)]
        [InlineData("$A$1:$C5", AddressScope.Range)]
        [InlineData("A0", AddressScope.Invalid)]
        [InlineData("ZZZZZZZZZZZZZZZZZ0", AddressScope.Invalid)]
        [InlineData("A1:C0", AddressScope.Invalid)]
        [InlineData("A1:ZZZZZZZZZZZZ0", AddressScope.Invalid)]
        [InlineData("ZZZZZZZZZZ1:C1", AddressScope.Invalid)]
        [InlineData("ZZZZZZZZZZZZZZ:ZZZZZZZZZZZZZ", AddressScope.Invalid)]
        [InlineData(":Z5", AddressScope.Invalid)]
        [InlineData("A2:", AddressScope.Invalid)]
        [InlineData(":", AddressScope.Invalid)]
        public void AddressScopeTest(String addressString, AddressScope expectedScope)
        {
            AddressScope scope = Cell.GetAddressScope(addressString);
            Assert.Equal(expectedScope, scope);
        }

        [Fact(DisplayName = "Test of the WorksheetReference property when cell is not assigned to a worksheet")]
        public void WorksheetReferenceTest()
        {
           Cell cell = new Cell(this.cellValue, CellType.DEFAULT, this.cellAddress);
           Assert.Null(cell.WorksheetReference);
        }


        [Fact(DisplayName = "Test of the WorksheetReference property when cell is added to a worksheet")]
        public void WorksheetReferenceTest2()
        {
            Worksheet dummyWorksheet = new Worksheet("dummy", 1, null);
            Worksheet worksheet = new Worksheet("worksheet1", 2, null);
            int expectedWorksheetId = worksheet.SheetID;
            worksheet.AddCell(this.cellValue, 1,0);
            Assert.Equal(this.cellValue, worksheet.GetCell(1,0).Value);
            Assert.Equal(expectedWorksheetId, worksheet.GetCell(1, 0).WorksheetReference.SheetID);
        }

        [Fact(DisplayName = "Test of the get function of the Style property")]
        public void CellStyleTest()
        {
            DateTime givenDate = new DateTime(2020, 11, 1, 12, 30, 22);
            Cell cell = utils.CreateVariantCell<DateTime>(givenDate, this.cellAddress, true);
            //Cell dateCell = new Cell(givenDate, CellType.DATE, this.cellAddress);
            Style expectedStyle = BasicStyles.DateFormat;
            Assert.Equal(expectedStyle, cell.CellStyle);
        }

        [Fact(DisplayName = "Test of the get function of the Style property, when no style was assigned")]
        public void CellStyleTest2()
        {
            Cell cell = new Cell(42, CellType.NUMBER, this.cellAddress);
            Assert.Null(cell.CellStyle);
        }

        [Fact(DisplayName = "Test of the set function of the Style property")]
        public void CellStyleTest3()
        {
            Cell cell = utils.CreateVariantCell<int>(42, this.cellAddress, true);
            Style style = BasicStyles.BoldItalic;
            Style returnedStyle = cell.SetStyle(style);
            Assert.NotNull(cell.CellStyle);
            Assert.Equal(style, cell.CellStyle);
            Assert.Equal(style, returnedStyle);
        }

        [Fact(DisplayName = "Test of the failing set function of the Style property, when the style is null")]
        public void CellStyleFailTest()
        {
            Cell intCell = new Cell(42, CellType.NUMBER, this.cellAddress);
            Style style = null;
            Exception ex = Assert.Throws<StyleException>(() => cell.SetStyle(style));
            Assert.Equal(typeof(StyleException), ex.GetType());
        }

        [Fact(DisplayName = "Test of the RemoveStyle method")]
        public void RemoveStyleTest()
        {
            Style style = BasicStyles.Bold;
            Cell floatCell = utils.CreateVariantCell<float>(11.11f, this.cellAddress, true, style);
            Assert.Equal(style, floatCell.CellStyle);
            floatCell.RemoveStyle();
            Assert.Null(floatCell.CellStyle);

        }

        [Fact(DisplayName = "Test of the Equals method (simplified use cases)")]
        public void EqualsTest()
        {
            AssertEquals<object>(null, null, "Data");
            AssertEquals<int>(27, 27, 28);
            AssertEquals<float>(0.27778f, 0.27778f, 0.27777f);
            AssertEquals<string>("ABC", "ABC", "abc");
            AssertEquals<string>("", "", " ");
            AssertEquals<bool>(true, true, false);
            AssertEquals<bool>(false, false, true);
            AssertEquals<DateTime>(new DateTime(11, 10, 9, 8, 7, 6), new DateTime(11, 10, 9, 8, 7, 6), new DateTime(11, 10, 9, 8, 7, 5));
        }

        [Fact(DisplayName = "Test failing of the Equals method, when other cell is null (simplified use cases)")]
        public void EqualsFailTest()
        {
            Cell cell1 = utils.CreateVariantCell<String>("test", this.cellAddress, true, BasicStyles.BoldItalic);
            Cell cell2 = null;
            Assert.False(cell1.Equals(cell2));
        }

        [Fact(DisplayName = "Test failing of the Equals method, when the address of the other cell is different (simplified use cases)")]
        public void EqualsFailTest2()
        {
            Cell cell1 = utils.CreateVariantCell<String>("test", this.cellAddress, true, BasicStyles.BoldItalic);
            Cell cell2 = utils.CreateVariantCell<String>("test", new Address(99, 99), true, BasicStyles.BoldItalic);
            Assert.False(cell1.Equals(cell2));
        }

        [Fact(DisplayName = "Test failing of the Equals method, when the style of the other cell is different (simplified use cases)")]
        public void EqualsFailTest4()
        {
            Cell cell1 = utils.CreateVariantCell<String>("test", this.cellAddress, true, BasicStyles.BoldItalic);
            Cell cell2 = utils.CreateVariantCell<String>("test", this.cellAddress, true, BasicStyles.Italic);
            Assert.False(cell1.Equals(cell2));
        }

        [Fact(DisplayName = "Test failing of the Equals method, when the workbook of the other cell is different (simplified use cases)")]
        public void EqualsFailTest5()
        {
            Workbook workbook1 = new Workbook(true);
            Workbook workbook2 = new Workbook(true);
            Cell cell1 = utils.CreateVariantCell<String>("test", this.cellAddress, true, BasicStyles.BoldItalic);
            cell1.WorksheetReference = workbook1.Worksheets[0];
            Cell cell2 = utils.CreateVariantCell<String>("test", this.cellAddress, true, BasicStyles.BoldItalic);
            cell2.WorksheetReference = workbook2.Worksheets[0];
            Assert.False(cell1.Equals(cell2));
        }

        [Fact(DisplayName = "Test failing of the Equals method, when the worksheet of the other cell is different (simplified use cases)")]
        public void EqualsFailTest6()
        {

            Workbook workbook1 = new Workbook(true);
            workbook1.AddWorksheet("worksheet2");
            Cell cell1 = utils.CreateVariantCell<String>("test", this.cellAddress, true, BasicStyles.BoldItalic);
            cell1.WorksheetReference = workbook1.Worksheets[0];
            Cell cell2 = utils.CreateVariantCell<String>("test", this.cellAddress, true, BasicStyles.BoldItalic);
            cell2.WorksheetReference = workbook1.Worksheets[1];
            Assert.False(cell1.Equals(cell2));
        }

        [Theory(DisplayName = "Test of the SetCellLockedState method")]
        [InlineData(true, true)]
        [InlineData(true, false)]
        [InlineData(false, true)]
        [InlineData(false, false)]
        public void SetCellLockedState(bool isLocked, bool isHidden)
        {
            Cell cell = utils.CreateVariantCell<String>("test", this.cellAddress, true);
            cell.SetCellLockedState(isLocked, isHidden);
            Assert.Equal(isLocked, cell.CellStyle.CurrentCellXf.Locked);
            Assert.Equal(isHidden, cell.CellStyle.CurrentCellXf.Hidden);
        }

        [Theory(DisplayName = "Test of the SetCellLockedState method when a cell style already exists")]
        [InlineData(true, true)]
        [InlineData(true, false)]
        [InlineData(false, true)]
        [InlineData(false, false)]
        public void SetCellLockedState2(bool isLocked, bool isHidden)
        {
            Style style = new Style();
            style.CurrentFont.Italic = true;
            style.CurrentCellXf.HorizontalAlign = CellXf.HorizontalAlignValue.justify;
            Cell cell = utils.CreateVariantCell<String>("test", this.cellAddress, true, style);
            cell.SetCellLockedState(isLocked, isHidden);
            Assert.Equal(isLocked, cell.CellStyle.CurrentCellXf.Locked);
            Assert.Equal(isHidden, cell.CellStyle.CurrentCellXf.Hidden);
            Assert.True(cell.CellStyle.CurrentFont.Italic);
            Assert.Equal(CellXf.HorizontalAlignValue.justify, cell.CellStyle.CurrentCellXf.HorizontalAlign);
        }

        private void AssertEquals<T>(T value1, T value2, T inequalValue)
        {
            Cell cell1 = new Cell(value1, CellType.DEFAULT, this.cellAddress);
            Cell cell2 = new Cell(value2, CellType.DEFAULT, this.cellAddress);
            Cell cell3 = new Cell(inequalValue, CellType.DEFAULT, this.cellAddress);
            cell1.Equals(cell2);
            Assert.True(cell1.Equals(cell2));
            Assert.False(cell1.Equals(cell3));
        }

        private static object SequentialCollection()
        {
            throw new NotImplementedException();
        }
    }
}
