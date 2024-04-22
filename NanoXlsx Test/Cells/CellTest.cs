
using NanoXLSX;
using NanoXLSX.Shared.Exceptions;
using NanoXLSX.Styles;
using NanoXLSX_Test.Cells.Types;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;
using static NanoXLSX.Shared.Enums.Styles.CellXfEnums;
using static NanoXLSX.Cell;

namespace NanoXLSX_Test.Cells
{
    // Ensure that these tests are executed sequentially, since static repository methods may be called 
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
            this.cell = this.utils.CreateVariantCell(this.cellValue as string, this.cellAddress, BasicStyles.BoldItalic);
        }


        [Fact(DisplayName = "Test of the default constructor")]
        public void CellConstructorTest()
        {
            Cell cell = new Cell();
            Assert.Equal(CellType.DEFAULT, cell.DataType);
            Assert.Null(cell.Value);
            Assert.Null(cell.CellStyle);
            Assert.Equal("A1", cell.CellAddress); // The address comes from initial row and column of 0
        }

        [Theory(DisplayName = "Test of the constructor with value and type")]
        [InlineData("string", "string", CellType.STRING, CellType.STRING)]
        [InlineData(true, true, CellType.BOOL, CellType.BOOL)]
        [InlineData(false, false, CellType .BOOL, CellType.BOOL)]
        [InlineData(22, 22, CellType.NUMBER, CellType.NUMBER)]
        [InlineData(22.1f, 22.1f, CellType.NUMBER, CellType.NUMBER)]
        [InlineData("=B1", "=B1", CellType.FORMULA, CellType.FORMULA)]
        [InlineData("", "", CellType.DEFAULT, CellType.STRING)]
        [InlineData(null, null, CellType.DEFAULT, CellType.EMPTY)]
        [InlineData("", null, CellType.EMPTY, CellType.EMPTY)]
        [InlineData(11, null, CellType.EMPTY, CellType.EMPTY)]
        public void CellConstructorTest2(object givenValue, object expectedValue, CellType givenType, CellType expectedType)
        {
            Cell cell = new Cell(givenValue, givenType);
            Assert.Equal(expectedType, cell.DataType);
            Assert.Equal(expectedValue, cell.Value);
            Assert.Null(cell.CellStyle);
            Assert.Equal("A1", cell.CellAddress); // The address comes from initial row and column of 0
        }

        [Fact(DisplayName = "Test of the constructor with value and type with a date as value")]
        public void CellConstructorTest2b()
        {
            DateTime date = DateTime.Now;
            Cell cell = new Cell(date, CellType.DEFAULT);
            Assert.Equal(CellType.DATE, cell.DataType);
            Assert.Equal(date, cell.Value);
            Assert.NotNull(cell.CellStyle);
            Assert.True(BasicStyles.DateFormat.Equals(cell.CellStyle));
        }

        [Theory(DisplayName = "Test of the constructor with value, type and address string")]
        [InlineData("string", "string", CellType.STRING, CellType.STRING, "C$7", "C$7")]
        [InlineData(true, true, CellType.BOOL, CellType.BOOL, "D100", "D100")]
        [InlineData(false, false, CellType.BOOL, CellType.BOOL, "$A$2", "$A$2")]
        [InlineData(22, 22, CellType.NUMBER, CellType.NUMBER, "$B5", "$B5")]
        [InlineData(22.1f, 22.1f, CellType.NUMBER, CellType.NUMBER, "AA10", "AA10")]
        [InlineData("=B1", "=B1", CellType.FORMULA, CellType.FORMULA, "$A$15", "$A$15")]
        [InlineData("", "", CellType.DEFAULT, CellType.STRING, "r1", "R1")]
        [InlineData(null, null, CellType.DEFAULT, CellType.EMPTY, "$ab$999", "$AB$999")]
        [InlineData("", null, CellType.EMPTY, CellType.EMPTY, "c$90000", "C$90000")]
        [InlineData(11, null, CellType.EMPTY, CellType.EMPTY, "a17", "A17")]
        public void CellConstructorTest3(object givenValue, object expectedValue, CellType givenType, CellType expectedType, string givenAddress, string expectedAddress)
        {
            Cell cell = new Cell(givenValue, givenType, givenAddress);
            Assert.Equal(expectedType, cell.DataType);
            Assert.Equal(expectedValue, cell.Value);
            Assert.Null(cell.CellStyle);
            Assert.Equal(expectedAddress, cell.CellAddress);
        }

        [Theory(DisplayName = "Test of the constructor with value, type and address object or row and column")]
        [InlineData("string", "string", CellType.STRING, CellType.STRING, 2, 6, "C7")]
        [InlineData(true, true, CellType.BOOL, CellType.BOOL, 3, 99, "D100")]
        [InlineData(false, false, CellType.BOOL, CellType.BOOL, 0, 1, "A2")]
        [InlineData(22, 22, CellType.NUMBER, CellType.NUMBER, 1, 4, "B5")]
        [InlineData(22.1f, 22.1f, CellType.NUMBER, CellType.NUMBER, 26, 9, "AA10")]
        [InlineData("=B1", "=B1", CellType.FORMULA, CellType.FORMULA, 0, 14, "A15")]
        [InlineData("", "", CellType.DEFAULT, CellType.STRING, 17, 0, "R1")]
        [InlineData(null, null, CellType.DEFAULT, CellType.EMPTY, 27, 998, "AB999")]
        [InlineData("", null, CellType.EMPTY, CellType.EMPTY, 2, 89999, "C90000")]
        [InlineData(11, null, CellType.EMPTY, CellType.EMPTY, 0, 16, "A17")]
        public void CellConstructorTest4(object givenValue, object expectedValue, CellType givenType, CellType expectedType, int givenColumn, int givenRow, string expectedAddress)
        {
            Address address = new Address(givenColumn, givenRow);
            Cell cell = new Cell(givenValue, givenType, address);
            Cell cell2 = new Cell(givenValue, givenType, givenColumn, givenRow);
            Assert.Equal(expectedType, cell.DataType);
            Assert.Equal(expectedValue, cell.Value);
            Assert.Null(cell.CellStyle);
            Assert.Equal(expectedAddress, cell.CellAddress);
            Assert.Equal(expectedType, cell2.DataType);
            Assert.Equal(expectedValue, cell2.Value);
            Assert.Null(cell2.CellStyle);
            Assert.Equal(expectedAddress, cell2.CellAddress);
        }

        [Theory(DisplayName = "Test of the set function of the CellAdrdess property")]
        [InlineData("A1", 0, 0, AddressType.Default)]
        [InlineData("J100", 9, 99, AddressType.Default)]
        [InlineData("$B2", 1, 1, AddressType.FixedColumn)]
        [InlineData("B$2", 1, 1, AddressType.FixedRow)]
        [InlineData("$B$2", 1, 1, AddressType.FixedRowAndColumn)]
        public void SetAddressStringPropertyTest(string givenAddress, int expectedColumn, int expectedRow, Cell.AddressType expectedType)
        {
            this.cell.CellAddress = givenAddress;
            Assert.Equal(this.cell.CellAddress2.Column, expectedColumn);
            Assert.Equal(this.cell.CellAddress2.Row, expectedRow);
            Assert.Equal(this.cell.CellAddress2.Type, expectedType);
        }

        [Theory(DisplayName = "Test of the get function of the CellAddressString property")]
        [InlineData(0, 0, AddressType.Default, "A1")]
        [InlineData(9, 99, AddressType.Default, "J100")]
        [InlineData(1, 1, AddressType.FixedColumn, "$B2")]
        [InlineData(1, 1, AddressType.FixedRow, "B$2")]
        [InlineData(1, 1, AddressType.FixedRowAndColumn, "$B$2")]
        public void GetAddressStringPropertyTest(int givenColumn, int givenRow, Cell.AddressType givenTyp, string expectedAddress)
        {
            this.cell.CellAddressType = givenTyp;
            this.cell.ColumnNumber = givenColumn;
            this.cell.RowNumber = givenRow;
            Assert.Equal(this.cell.CellAddress, expectedAddress);
        }

        [Theory(DisplayName = "Test of the set function of the CellAddress property, as well as get functions of columnNumber, RowNumber and AddressType")]
        [InlineData("A1", 0, 0, AddressType.Default)]
        [InlineData("XFD1048576", 16383, 1048575, AddressType.Default)]
        [InlineData("$B2", 1, 1, AddressType.FixedColumn)]
        [InlineData("B$2", 1, 1, AddressType.FixedRow)]
        [InlineData("$B$2", 1, 1, AddressType.FixedRowAndColumn)]
        public void SetAddressPropertyTest(string givenAddress, int expectedColumn, int expectedRow, Cell.AddressType expectedType)
        {
            Address given = new Address(givenAddress);
            this.cell.CellAddress2 = given;
            Assert.Equal(this.cell.ColumnNumber, expectedColumn);
            Assert.Equal(this.cell.RowNumber, expectedRow);
            Assert.Equal(this.cell.CellAddressType, expectedType);
        }

        [Theory(DisplayName = "Test of the get function of the CellAddress property, as well as set functions of columnNumber, RowNumber and AddressType")]
        [InlineData(0, 0, AddressType.Default, "A1")]
        [InlineData(16383, 1048575, AddressType.Default, "XFD1048576")]
        [InlineData(1, 1, AddressType.FixedColumn, "$B2")]
        [InlineData(1, 1, AddressType.FixedRow, "B$2")]
        [InlineData(1, 1, AddressType.FixedRowAndColumn, "$B$2")]
        public void GetAddressPropertyTest(int givenColumn, int givenRow, Cell.AddressType givenTyp, string expectedAddress)
        {
            this.cell.ColumnNumber = givenColumn;
            this.cell.RowNumber = givenRow;
            this.cell.CellAddressType = givenTyp;
            Address expected = new Address(expectedAddress);
            Assert.Equal(this.cell.CellAddress2, expected);
        }

        [Fact(DisplayName = "Test of the get function of the Style property")]
        public void CellStyleTest()
        {
            DateTime givenDate = new DateTime(2020, 11, 1, 12, 30, 22);
            Cell cell = utils.CreateVariantCell<DateTime>(givenDate, this.cellAddress);
            Style expectedStyle = BasicStyles.DateFormat;
            Assert.True(expectedStyle.Equals(cell.CellStyle)); // Note: Assert.Equals fails here because of object reference comparison 
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
            Cell cell = utils.CreateVariantCell<int>(42, this.cellAddress);
            Assert.Null(cell.CellStyle);
            Style style = BasicStyles.BoldItalic;
            Style returnedStyle = cell.SetStyle(style);
            Assert.NotNull(cell.CellStyle);
            // Note: Assert.Equals fails here because of object reference comparison 
            Assert.True(style.Equals(cell.CellStyle));
            Assert.True(style.Equals(returnedStyle));
            Assert.NotEmpty(StyleRepository.Instance.Styles);
        }

        [Fact(DisplayName = "Test of the set function of the Style property where the style repository is unmanaged")]
        public void CellStyleTest3b()
        {
            Cell cell = utils.CreateVariantCell<int>(42, this.cellAddress);
            Assert.Null(cell.CellStyle);
            StyleRepository.Instance.FlushStyles();
            Assert.Empty(StyleRepository.Instance.Styles);
            Style style = BasicStyles.BoldItalic;
            cell.SetStyle(style, true);
            // Note: Assert.Equals fails here because of object reference comparison 
            Assert.True(style.Equals(cell.CellStyle));
            Assert.Empty(StyleRepository.Instance.Styles);
        }

        [Fact(DisplayName = "Test of the failing set function of the Style property, when the style is null")]
        public void CellStyleFailTest()
        {
            Cell intCell = new Cell(42, CellType.NUMBER, this.cellAddress);
            Style style = null;
            Exception ex = Assert.Throws<StyleException>(() => intCell.SetStyle(style));
            Assert.Equal(typeof(StyleException), ex.GetType());
        }

        [Theory(DisplayName = "Test of the set function of the Value property on changing the data type of the cell")]
        [InlineData(0, Cell.CellType.NUMBER, "test", CellType.STRING)]
        [InlineData(true, Cell.CellType.BOOL, 1, CellType.NUMBER)]
        [InlineData(22.5f, Cell.CellType.NUMBER, 22, CellType.NUMBER)]
        [InlineData("True", Cell.CellType.STRING, true, CellType.BOOL)]
        [InlineData(null, Cell.CellType.EMPTY, 22, CellType.NUMBER)]
        [InlineData("test", Cell.CellType.STRING, null, CellType.EMPTY)]
        public void SetValuePropertyTest(object initialValue, Cell.CellType initialType, object givenValue, Cell.CellType expectedType)
        {
            Cell cell2 = new Cell(initialValue, initialType);
            Assert.Equal(cell2.DataType, initialType);
            cell2.Value = givenValue;
            Assert.Equal(expectedType, cell2.DataType);
        }


        [Fact(DisplayName = "Test of the RemoveStyle method")]
        public void RemoveStyleTest()
        {
            Style style = BasicStyles.Bold;
            Cell floatCell = utils.CreateVariantCell<float>(11.11f, this.cellAddress, style);
            Assert.True(style.Equals(floatCell.CellStyle)); // Note: Assert.Equals fails here because of object reference comparison 
            floatCell.RemoveStyle();
            Assert.Null(floatCell.CellStyle);

        }

        [Theory(DisplayName = "Test of the CompareTo method (simplified use cases)")]
        [InlineData("A1", "A1", 0)]
        [InlineData("A1", "A2", -1)]
        [InlineData("A1", "B1", -1)]
        [InlineData("A2", "A1", 1)]
        [InlineData("B1", "A1", 1)]
        [InlineData("A1", null, -1)]
        public void CompareToTest(string cell1Address, string cell2Address, int expectedResult)
        {
            Cell cell1 = new Cell("Test", CellType.DEFAULT, cell1Address);
            Cell cell2 = null;
            if (cell2Address != null)
            {
                cell2 = new Cell("Test", CellType.DEFAULT, cell2Address);
            }
            Assert.Equal(expectedResult, cell1.CompareTo(cell2));
        }

        [Fact(DisplayName = "Test of the Equals method (simplified use cases)")]
        public void EqualsTest()
        {
            TestUtils.AssertEquals<object>(null, null, "Data", this.cellAddress);
            TestUtils.AssertEquals<int>(27, 27, 28, this.cellAddress);
            TestUtils.AssertEquals<float>(0.27778f, 0.27778f, 0.27777f, this.cellAddress);
            TestUtils.AssertEquals<string>("ABC", "ABC", "abc", this.cellAddress);
            TestUtils.AssertEquals<string>("", "", " ", this.cellAddress);
            TestUtils.AssertEquals<bool>(true, true, false, this.cellAddress);
            TestUtils.AssertEquals<bool>(false, false, true, this.cellAddress);
            TestUtils.AssertEquals<DateTime>(new DateTime(2020, 10, 9, 8, 7, 6), new DateTime(2020, 10, 9, 8, 7, 6), new DateTime(2020, 10, 9, 8, 7, 5), this.cellAddress);
        }

        [Fact(DisplayName = "Test failing of the Equals method, when other cell is null (simplified use cases)")]
        public void EqualsFailTest()
        {
            Cell cell1 = utils.CreateVariantCell<string>("test", this.cellAddress, BasicStyles.BoldItalic);
            Cell cell2 = null;
            Assert.False(cell1.Equals(cell2));
        }

        [Fact(DisplayName = "Test failing of the Equals method, when the address of the other cell is different (simplified use cases)")]
        public void EqualsFailTest2()
        {
            Cell cell1 = utils.CreateVariantCell<string>("test", this.cellAddress, BasicStyles.BoldItalic);
            Cell cell2 = utils.CreateVariantCell<string>("test", new Address(99, 99), BasicStyles.BoldItalic);
            Assert.False(cell1.Equals(cell2));
        }

        [Fact(DisplayName = "Test failing of the Equals method, when the style of the other cell is different (simplified use cases)")]
        public void EqualsFailTest3()
        {
            Cell cell1 = utils.CreateVariantCell<string>("test", this.cellAddress, BasicStyles.BoldItalic);
            Cell cell2 = utils.CreateVariantCell<string>("test", this.cellAddress, BasicStyles.Italic);
            Assert.False(cell1.Equals(cell2));
        }
        
        [Fact(DisplayName = "Test of the Equals method, when two identical cells occur in different workbooks and worksheets (simplified use cases)")]
        public void EqualsFailTest4()
        {
            Workbook workbook1 = new Workbook(true);
            Workbook workbook2 = new Workbook(true);
            Cell cell1 = utils.CreateVariantCell<string>("test", this.cellAddress, BasicStyles.BoldItalic);
            workbook1.CurrentWorksheet.AddCell(cell1, this.cellAddress.GetAddress());
            Cell cell2 = utils.CreateVariantCell<string>("test", this.cellAddress, BasicStyles.BoldItalic);
            workbook2.CurrentWorksheet.AddCell(cell2, this.cellAddress.GetAddress());
            Cell cell3 = utils.CreateVariantCell<string>("test", this.cellAddress, BasicStyles.BoldItalic);
            workbook2.AddWorksheet("2nd");
            workbook2.Worksheets[1].AddCell(cell3, this.cellAddress.GetAddress());
            Assert.True(cell1.Equals(cell2));
            Assert.True(cell2.Equals(cell3));
        }

        [Theory(DisplayName = "Test of the CompareTo method (simplified use cases)")]
        [InlineData("string", CellType.NUMBER, CellType.STRING)]
        [InlineData(12, CellType.STRING, CellType.NUMBER)]
        [InlineData(-12.12d, CellType.STRING, CellType.NUMBER)]
        [InlineData(true, CellType.STRING, CellType.BOOL)]
        [InlineData(false, CellType.STRING, CellType.BOOL)]
        [InlineData("=A2", CellType.FORMULA, CellType.FORMULA)]
        [InlineData(null, CellType.STRING, CellType.EMPTY)]
        [InlineData("Actually not empty", CellType.EMPTY, CellType.EMPTY)]
        [InlineData("string", CellType.DEFAULT, CellType.STRING)]
        [InlineData(0, CellType.DEFAULT, CellType.NUMBER)]
        [InlineData(-12.12f, CellType.DEFAULT, CellType.NUMBER)]
        [InlineData(true, CellType.DEFAULT, CellType.BOOL)]
        [InlineData(false, CellType.DEFAULT, CellType.BOOL)]
        [InlineData("=A2", CellType.DEFAULT, CellType.STRING)]
        [InlineData("", CellType.DEFAULT, CellType.STRING)]
        [InlineData(null, CellType.DEFAULT, CellType.EMPTY)]
        public void ResolveCellTypeTest(object givenValue, CellType givenCellType, CellType expectedCllType)
        {
            Cell cell = new Cell(givenValue, givenCellType, this.cellAddress);
            cell.ResolveCellType();
            Assert.Equal(expectedCllType, cell.DataType);
        }

        [Fact(DisplayName = "Test of the CompareTo method for dates and times")]
        public void ResolveCellTypeTest2()
        {
            Cell dateCell = new Cell(DateTime.Now, CellType.NUMBER, this.cellAddress);
            dateCell.ResolveCellType();
            Assert.Equal(CellType.DATE, dateCell.DataType);
            dateCell = new Cell(DateTime.Now, CellType.DEFAULT, this.cellAddress);
            dateCell.ResolveCellType();
            Assert.Equal(CellType.DATE, dateCell.DataType);
            Cell timeCell = new Cell(TimeSpan.FromMinutes(60), CellType.NUMBER, this.cellAddress);
            timeCell.ResolveCellType();
            Assert.Equal(CellType.TIME, timeCell.DataType);
            timeCell = new Cell(TimeSpan.FromMinutes(60), CellType.DEFAULT, this.cellAddress);
            dateCell.ResolveCellType();
            Assert.Equal(CellType.TIME, timeCell.DataType);
        }

        [Theory(DisplayName = "Test of the SetCellLockedState method")]
        [InlineData(true, true)]
        [InlineData(true, false)]
        [InlineData(false, true)]
        [InlineData(false, false)]
        public void SetCellLockedState(bool isLocked, bool isHidden)
        {
            Cell cell = utils.CreateVariantCell<string>("test", this.cellAddress);
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
            style.CurrentCellXf.HorizontalAlign = HorizontalAlignValue.justify;
            Cell cell = utils.CreateVariantCell<string>("test", this.cellAddress, style);
            cell.SetCellLockedState(isLocked, isHidden);
            Assert.Equal(isLocked, cell.CellStyle.CurrentCellXf.Locked);
            Assert.Equal(isHidden, cell.CellStyle.CurrentCellXf.Hidden);
            Assert.True(cell.CellStyle.CurrentFont.Italic);
            Assert.Equal(HorizontalAlignValue.justify, cell.CellStyle.CurrentCellXf.HorizontalAlign);
        }

        [Theory(DisplayName = "Test of the GetCellRange method with string as range")]
        [InlineData("A1:A1", "A1")]
        [InlineData("A1:A4", "A1,A2,A3,A4")]
        [InlineData("A1:B3", "A1,A2,A3,B1,B2,B3")]
        [InlineData("B3:A2", "A2,A3,B2,B3")]
        public void GetCellRangeTest(string range, string expectedAddresses)
        {
            List<Address> addresses = Cell.GetCellRange(range).ToList();
            TestUtils.AssertCellRange(expectedAddresses, addresses);
        }

        [Theory(DisplayName = "Test of the GetCellRange method with start and end address objects or strings as range")]
        [InlineData("A1", "A1", "A1")]
        [InlineData("A1", "A4", "A1,A2,A3,A4")]
        [InlineData("A1", "B3", "A1,A2,A3,B1,B2,B3")]
        [InlineData("B3", "A2", "A2,A3,B2,B3")]
        public void GetCellRangeTest2(string startAddress, string endAddress, string expectedAddresses)
        {
            Address start = new Address(startAddress);
            Address end = new Address(endAddress);
            List<Address> addresses = Cell.GetCellRange(startAddress, endAddress).ToList();
            TestUtils.AssertCellRange(expectedAddresses, addresses);
            addresses = Cell.GetCellRange(start, end).ToList();
            TestUtils.AssertCellRange(expectedAddresses, addresses);
        }

        [Theory(DisplayName = "Test of the GetCellRange method with start/end row and column numbers as range")]
        [InlineData(0, 0, 0, 0, "A1")]
        [InlineData(0, 0, 0, 3, "A1,A2,A3,A4")]
        [InlineData(0, 0, 1, 2, "A1,A2,A3,B1,B2,B3")]
        [InlineData(1, 2, 0, 1, "A2,A3,B2,B3")]
        public void GetCellRangeTest3(int startColumn, int startRow, int endColumn, int endRow, string expectedAddresses)
        {
            List<Address> addresses = Cell.GetCellRange(startColumn, startRow, endColumn, endRow).ToList();
            TestUtils.AssertCellRange(expectedAddresses, addresses);
        }

        [Theory(DisplayName = "Test of the ResolveCellAddress method")]
        [InlineData(0, 0, AddressType.Default, "A1")]
        [InlineData(0, 0, AddressType.FixedColumn, "$A1")]
        [InlineData(0, 0, AddressType.FixedRow, "A$1")]
        [InlineData(0, 0, AddressType.FixedRowAndColumn, "$A$1")]
        [InlineData(5, 99, AddressType.Default, "F100")]
        [InlineData(5, 99, AddressType.FixedColumn, "$F100")]
        [InlineData(5, 99, AddressType.FixedRow, "F$100")]
        [InlineData(5, 99, AddressType.FixedRowAndColumn, "$F$100")]
        [InlineData(16383, 1048575, AddressType.Default, "XFD1048576")]
        [InlineData(16383, 1048575, AddressType.FixedColumn, "$XFD1048576")]
        [InlineData(16383, 1048575, AddressType.FixedRow, "XFD$1048576")]
        [InlineData(16383, 1048575, AddressType.FixedRowAndColumn, "$XFD$1048576")]
        public void ResolveCellAddressTest(int column, int row, AddressType type, string expectedAddress)
        {
            string address = Cell.ResolveCellAddress(column, row, type);
            Assert.Equal(expectedAddress, address);
        }

        [Theory(DisplayName = "Test of the  ResolveCellCoordinate method with string as parameter")]
        [InlineData("A1", 0, 0, AddressType.Default)]
        [InlineData("$A1", 0, 0, AddressType.FixedColumn)]
        [InlineData("A$1", 0, 0, AddressType.FixedRow)]
        [InlineData("$A$1", 0, 0, AddressType.FixedRowAndColumn)]
        [InlineData("F100", 5, 99, AddressType.Default)]
        [InlineData("$F100", 5, 99, AddressType.FixedColumn)]
        [InlineData("F$100", 5, 99, AddressType.FixedRow)]
        [InlineData("$F$100", 5, 99, AddressType.FixedRowAndColumn)]
        [InlineData("XFD1048576", 16383, 1048575, AddressType.Default)]
        [InlineData("$XFD1048576", 16383, 1048575, AddressType.FixedColumn)]
        [InlineData("XFD$1048576", 16383, 1048575, AddressType.FixedRow)]
        [InlineData("$XFD$1048576", 16383, 1048575, AddressType.FixedRowAndColumn)]
        public void ResolveCellCoordinateTest(string addressString, int expectedColumn, int expectedRow, AddressType expectedType)
        {
            Address address = Cell.ResolveCellCoordinate(addressString);
            Assert.Equal(expectedColumn, address.Column);
            Assert.Equal(expectedRow, address.Row);
            Assert.Equal(expectedType, address.Type);
        }

        [Theory(DisplayName = "Test of the  ResolveCellCoordinate method with string as parameter and out parameters")]
        [InlineData("A1", 0, 0, AddressType.Default)]
        [InlineData("$A1", 0, 0, AddressType.FixedColumn)]
        [InlineData("A$1", 0, 0, AddressType.FixedRow)]
        [InlineData("$A$1", 0, 0, AddressType.FixedRowAndColumn)]
        [InlineData("F100", 5, 99, AddressType.Default)]
        [InlineData("$F100", 5, 99, AddressType.FixedColumn)]
        [InlineData("F$100", 5, 99, AddressType.FixedRow)]
        [InlineData("$F$100", 5, 99, AddressType.FixedRowAndColumn)]
        [InlineData("XFD1048576", 16383, 1048575, AddressType.Default)]
        [InlineData("$XFD1048576", 16383, 1048575, AddressType.FixedColumn)]
        [InlineData("XFD$1048576", 16383, 1048575, AddressType.FixedRow)]
        [InlineData("$XFD$1048576", 16383, 1048575, AddressType.FixedRowAndColumn)]
        public void ResolveCellCoordinateTest2(string addressString, int expectedColumn, int expectedRow, AddressType expectedType)
        {
            int row, column;
            AddressType type;
            Cell.ResolveCellCoordinate(addressString, out column, out row);
            Assert.Equal(expectedColumn, column);
            Assert.Equal(expectedRow, row);
            // Other overloading
            Cell.ResolveCellCoordinate(addressString, out column, out row, out type);
            Assert.Equal(expectedColumn, column);
            Assert.Equal(expectedRow, row);
            Assert.Equal(expectedType, type);
        }

        [Theory(DisplayName = "Test of the ResolveCellRange method")]
        [InlineData("a1:a1", "A1", "A1")]
        [InlineData("C3:C4", "C3", "C4")]
        [InlineData("$a1:Z$10", "$A1", "Z$10")]
        [InlineData("$R$9:a2", "A2", "$R$9")]
        [InlineData("A1", "A1", "A1")]
        public void ResolveCellRangeTest(string rangeString, string expectedStartAddress, string expectedEndAddress) 
        {
            NanoXLSX.Range range = Cell.ResolveCellRange(rangeString);
            Address start = new Address(expectedStartAddress);
            Address end = new Address(expectedEndAddress);
            Assert.Equal(start, range.StartAddress);
            Assert.Equal(end, range.EndAddress);
        }

        [Fact(DisplayName = "Test of the failing ResolveCellRange method")]
        public void ResolveCellRangeTest2()
        {
            Exception ex = Assert.Throws<NanoXLSX.Shared.Exceptions.FormatException>(() => Cell.ResolveCellRange(null));
            Assert.Equal(typeof(NanoXLSX.Shared.Exceptions.FormatException), ex.GetType());
            ex = Assert.Throws<NanoXLSX.Shared.Exceptions.FormatException>(() => Cell.ResolveCellRange(""));
            Assert.Equal(typeof(NanoXLSX.Shared.Exceptions.FormatException), ex.GetType());
        }

        [Theory(DisplayName = "Test of the ResolveColumn method")]
        [InlineData("A", 0)]
        [InlineData("c", 2)]
        [InlineData("XFD", 16383)]
        public void ResolveColumnTest(string address, int expectedColumn)
        {
            int column = Cell.ResolveColumn(address);
            Assert.Equal(expectedColumn, column);
        }

        [Fact(DisplayName = "Test of the failing ResolveColumn method")]
        public void ResolveColumnTest2()
        {
            Exception ex = Assert.Throws<RangeException>(() => Cell.ResolveColumn(null));
            Assert.Equal(typeof(RangeException), ex.GetType());
            ex = Assert.Throws<RangeException>(() => Cell.ResolveColumn(""));
            Assert.Equal(typeof(RangeException), ex.GetType());
            ex = Assert.Throws<RangeException>(() => Cell.ResolveColumn("XFE"));
            Assert.Equal(typeof(RangeException), ex.GetType());
        }

        [Theory(DisplayName = "Test of the ResolveColumnAddress method")]
        [InlineData(0, "A")]
        [InlineData(2, "C")]
        [InlineData(16383, "XFD")]
        public void ResolveColumnAddressTest(int columnNumber, string expectedAddress)
        {
            string address = Cell.ResolveColumnAddress(columnNumber);
            Assert.Equal(expectedAddress, address);
        }

        [Fact(DisplayName = "Test of the failing ResolveColumnAddress method")]
        public void ResolveColumnAddressTest2()
        {
            Exception ex = Assert.Throws<RangeException>(() => Cell.ResolveColumnAddress(-1));
            Assert.Equal(typeof(RangeException), ex.GetType());
            ex = Assert.Throws<RangeException>(() => Cell.ResolveColumnAddress(16384));
            Assert.Equal(typeof(RangeException), ex.GetType());
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
        public void AddressScopeTest(string addressString, AddressScope expectedScope)
        {
            AddressScope scope = Cell.GetAddressScope(addressString);
            Assert.Equal(expectedScope, scope);
        }


    }
}
