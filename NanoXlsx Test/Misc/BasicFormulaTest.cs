using NanoXLSX;
using System;
using Xunit;
using FormatException = NanoXLSX.Exceptions.FormatException;
using Range = NanoXLSX.Range;

namespace NanoXLSX_Test.Misc
{
    public class BasicFormulaTest
    {
        [Theory(DisplayName = "Test of the Average function on a Range object")]
        [InlineData("A1:A1", "AVERAGE(A1)")]
        [InlineData("A1:C2", "AVERAGE(A1:C2)")]
        [InlineData("$A1:C2", "AVERAGE($A1:C2)")]
        [InlineData("$A$1:C2", "AVERAGE($A$1:C2)")]
        [InlineData("$A$1:$C2", "AVERAGE($A$1:$C2)")]
        [InlineData("$A$1:$C$2", "AVERAGE($A$1:$C$2)")]
        public void AverageTest(string rangeExpression, string expectedFormula)
        {
            Range range = new Range(rangeExpression);
            Cell formula = BasicFormulas.Average(range);
            Assert.Equal(expectedFormula, formula.Value.ToString());
            Assert.Equal(Cell.CellType.FORMULA, formula.DataType);
        }

        [Theory(DisplayName = "Test of the Average function on a Range object and a target worksheet")]
        [InlineData("worksheet1", "A1:A1", "AVERAGE(worksheet1!A1)")]
        [InlineData("worksheet1", "A1:C2", "AVERAGE(worksheet1!A1:C2)")]
        [InlineData("worksheet1", "$A1:C2", "AVERAGE(worksheet1!$A1:C2)")]
        [InlineData("worksheet1", "$A$1:C2", "AVERAGE(worksheet1!$A$1:C2)")]
        [InlineData("worksheet1", "$A$1:$C2", "AVERAGE(worksheet1!$A$1:$C2)")]
        [InlineData("worksheet1", "$A$1:$C$2", "AVERAGE(worksheet1!$A$1:$C$2)")]
        public void AverageTest2(string worksheetName, string rangeExpression, string expectedFormula)
        {
            Worksheet worksheet = new Worksheet(worksheetName);
            Range range = new Range(rangeExpression);
            Cell formula = BasicFormulas.Average(worksheet, range);
            Assert.Equal(expectedFormula, formula.Value.ToString());
            Assert.Equal(Cell.CellType.FORMULA, formula.DataType);
        }

        [Theory(DisplayName = "Test of the Ceil function on a value and a number of decimals")]
        [InlineData("A1", 1, "ROUNDUP(A1,1)")]
        [InlineData("C4", 0, "ROUNDUP(C4,0)")]
        [InlineData("$A1", 10, "ROUNDUP($A1,10)")]
        [InlineData("$A$1", 5, "ROUNDUP($A$1,5)")]
        [InlineData("A1", -2, "ROUNDUP(A1,-2)")] // This seems to be valid
        public void CeilTest(string addressExpression, int numberOfDecimals, string expectedFormula)
        {
            Address address = new Address(addressExpression);
            Cell formula = BasicFormulas.Ceil(address, numberOfDecimals);
            Assert.Equal(expectedFormula, formula.Value.ToString());
            Assert.Equal(Cell.CellType.FORMULA, formula.DataType);
        }

        [Theory(DisplayName = "Test of the Ceil function on a value and a number of decimals")]
        [InlineData("worksheet3", "A1", 1, "ROUNDUP(worksheet3!A1,1)")]
        [InlineData("worksheet3", "C4", 0, "ROUNDUP(worksheet3!C4,0)")]
        [InlineData("worksheet3", "$A1", 10, "ROUNDUP(worksheet3!$A1,10)")]
        [InlineData("worksheet3", "$A$1", 5, "ROUNDUP(worksheet3!$A$1,5)")]
        [InlineData("worksheet3", "A1", -2, "ROUNDUP(worksheet3!A1,-2)")] // This seems to be valid
        public void CeilTest2(string worksheetName, string addressExpression, int numberOfDecimals, string expectedFormula)
        {
            Worksheet worksheet = new Worksheet(worksheetName);
            Address address = new Address(addressExpression);
            Cell formula = BasicFormulas.Ceil(worksheet, address, numberOfDecimals);
            Assert.Equal(expectedFormula, formula.Value.ToString());
            Assert.Equal(Cell.CellType.FORMULA, formula.DataType);
        }

        [Theory(DisplayName = "Test of the Floor function on a value and a number of decimals")]
        [InlineData("A1", 1, "ROUNDDOWN(A1,1)")]
        [InlineData("C4", 0, "ROUNDDOWN(C4,0)")]
        [InlineData("$A1", 10, "ROUNDDOWN($A1,10)")]
        [InlineData("$A$1", 5, "ROUNDDOWN($A$1,5)")]
        [InlineData("A1", -2, "ROUNDDOWN(A1,-2)")] // This seems to be valid
        public void FloorTest(string addressExpression, int numberOfDecimals, string expectedFormula)
        {
            Address address = new Address(addressExpression);
            Cell formula = BasicFormulas.Floor(address, numberOfDecimals);
            Assert.Equal(expectedFormula, formula.Value.ToString());
            Assert.Equal(Cell.CellType.FORMULA, formula.DataType);
        }

        [Theory(DisplayName = "Test of the Floor function on a value and a number of decimals")]
        [InlineData("worksheet3", "A1", 1, "ROUNDDOWN(worksheet3!A1,1)")]
        [InlineData("worksheet3", "C4", 0, "ROUNDDOWN(worksheet3!C4,0)")]
        [InlineData("worksheet3", "$A1", 10, "ROUNDDOWN(worksheet3!$A1,10)")]
        [InlineData("worksheet3", "$A$1", 5, "ROUNDDOWN(worksheet3!$A$1,5)")]
        [InlineData("worksheet3", "A1", -2, "ROUNDDOWN(worksheet3!A1,-2)")] // This seems to be valid
        public void FloorTest2(string worksheetName, string addressExpression, int numberOfDecimals, string expectedFormula)
        {
            Worksheet worksheet = new Worksheet(worksheetName);
            Address address = new Address(addressExpression);
            Cell formula = BasicFormulas.Floor(worksheet, address, numberOfDecimals);
            Assert.Equal(expectedFormula, formula.Value.ToString());
            Assert.Equal(Cell.CellType.FORMULA, formula.DataType);
        }

        [Theory(DisplayName = "Test of the Max function on a Range object")]
        [InlineData("A1:A1", "MAX(A1)")]
        [InlineData("A1:C2", "MAX(A1:C2)")]
        [InlineData("$A1:C2", "MAX($A1:C2)")]
        [InlineData("$A$1:C2", "MAX($A$1:C2)")]
        [InlineData("$A$1:$C2", "MAX($A$1:$C2)")]
        [InlineData("$A$1:$C$2", "MAX($A$1:$C$2)")]
        public void MaxTest(string rangeExpression, string expectedFormula)
        {
            Range range = new Range(rangeExpression);
            Cell formula = BasicFormulas.Max(range);
            Assert.Equal(expectedFormula, formula.Value.ToString());
            Assert.Equal(Cell.CellType.FORMULA, formula.DataType);
        }

        [Theory(DisplayName = "Test of the Max function on a Range object and a target worksheet")]
        [InlineData("worksheet1", "A1:A1", "MAX(worksheet1!A1)")]
        [InlineData("worksheet1", "A1:C2", "MAX(worksheet1!A1:C2)")]
        [InlineData("worksheet1", "$A1:C2", "MAX(worksheet1!$A1:C2)")]
        [InlineData("worksheet1", "$A$1:C2", "MAX(worksheet1!$A$1:C2)")]
        [InlineData("worksheet1", "$A$1:$C2", "MAX(worksheet1!$A$1:$C2)")]
        [InlineData("worksheet1", "$A$1:$C$2", "MAX(worksheet1!$A$1:$C$2)")]
        public void MaxTest2(string worksheetName, string rangeExpression, string expectedFormula)
        {
            Worksheet worksheet = new Worksheet(worksheetName);
            Range range = new Range(rangeExpression);
            Cell formula = BasicFormulas.Max(worksheet, range);
            Assert.Equal(expectedFormula, formula.Value.ToString());
            Assert.Equal(Cell.CellType.FORMULA, formula.DataType);
        }

        [Theory(DisplayName = "Test of the Min function on a Range object")]
        [InlineData("A1:A1", "MIN(A1)")]
        [InlineData("A1:C2", "MIN(A1:C2)")]
        [InlineData("$A1:C2", "MIN($A1:C2)")]
        [InlineData("$A$1:C2", "MIN($A$1:C2)")]
        [InlineData("$A$1:$C2", "MIN($A$1:$C2)")]
        [InlineData("$A$1:$C$2", "MIN($A$1:$C$2)")]
        public void MinTest(string rangeExpression, string expectedFormula)
        {
            Range range = new Range(rangeExpression);
            Cell formula = BasicFormulas.Min(range);
            Assert.Equal(expectedFormula, formula.Value.ToString());
            Assert.Equal(Cell.CellType.FORMULA, formula.DataType);
        }

        [Theory(DisplayName = "Test of the Min function on a Range object and a target worksheet")]
        [InlineData("worksheet1", "A1:A1", "MIN(worksheet1!A1)")]
        [InlineData("worksheet1", "A1:C2", "MIN(worksheet1!A1:C2)")]
        [InlineData("worksheet1", "$A1:C2", "MIN(worksheet1!$A1:C2)")]
        [InlineData("worksheet1", "$A$1:C2", "MIN(worksheet1!$A$1:C2)")]
        [InlineData("worksheet1", "$A$1:$C2", "MIN(worksheet1!$A$1:$C2)")]
        [InlineData("worksheet1", "$A$1:$C$2", "MIN(worksheet1!$A$1:$C$2)")]
        public void MinTest2(string worksheetName, string rangeExpression, string expectedFormula)
        {
            Worksheet worksheet = new Worksheet(worksheetName);
            Range range = new Range(rangeExpression);
            Cell formula = BasicFormulas.Min(worksheet, range);
            Assert.Equal(expectedFormula, formula.Value.ToString());
            Assert.Equal(Cell.CellType.FORMULA, formula.DataType);
        }

        [Theory(DisplayName = "Test of the Median function on a Range object")]
        [InlineData("A1:A1", "MEDIAN(A1)")]
        [InlineData("A1:C2", "MEDIAN(A1:C2)")]
        [InlineData("$A1:C2", "MEDIAN($A1:C2)")]
        [InlineData("$A$1:C2", "MEDIAN($A$1:C2)")]
        [InlineData("$A$1:$C2", "MEDIAN($A$1:$C2)")]
        [InlineData("$A$1:$C$2", "MEDIAN($A$1:$C$2)")]
        public void MedianTest(string rangeExpression, string expectedFormula)
        {
            Range range = new Range(rangeExpression);
            Cell formula = BasicFormulas.Median(range);
            Assert.Equal(expectedFormula, formula.Value.ToString());
            Assert.Equal(Cell.CellType.FORMULA, formula.DataType);
        }

        [Theory(DisplayName = "Test of the Median function on a Range object and a target worksheet")]
        [InlineData("worksheet1", "A1:A1", "MEDIAN(worksheet1!A1)")]
        [InlineData("worksheet1", "A1:C2", "MEDIAN(worksheet1!A1:C2)")]
        [InlineData("worksheet1", "$A1:C2", "MEDIAN(worksheet1!$A1:C2)")]
        [InlineData("worksheet1", "$A$1:C2", "MEDIAN(worksheet1!$A$1:C2)")]
        [InlineData("worksheet1", "$A$1:$C2", "MEDIAN(worksheet1!$A$1:$C2)")]
        [InlineData("worksheet1", "$A$1:$C$2", "MEDIAN(worksheet1!$A$1:$C$2)")]
        public void MedianTest2(string worksheetName, string rangeExpression, string expectedFormula)
        {
            Worksheet worksheet = new Worksheet(worksheetName);
            Range range = new Range(rangeExpression);
            Cell formula = BasicFormulas.Median(worksheet, range);
            Assert.Equal(expectedFormula, formula.Value.ToString());
            Assert.Equal(Cell.CellType.FORMULA, formula.DataType);
        }

        [Theory(DisplayName = "Test of the Round function on a value and a number of decimals")]
        [InlineData("A1", 1, "ROUND(A1,1)")]
        [InlineData("C4", 0, "ROUND(C4,0)")]
        [InlineData("$A1", 10, "ROUND($A1,10)")]
        [InlineData("$A$1", 5, "ROUND($A$1,5)")]
        [InlineData("A1", -2, "ROUND(A1,-2)")] // This seems to be valid
        public void RoundTest(string addressExpression, int numberOfDecimals, string expectedFormula)
        {
            Address address = new Address(addressExpression);
            Cell formula = BasicFormulas.Round(address, numberOfDecimals);
            Assert.Equal(expectedFormula, formula.Value.ToString());
            Assert.Equal(Cell.CellType.FORMULA, formula.DataType);
        }

        [Theory(DisplayName = "Test of the Round function on a value and a number of decimals")]
        [InlineData("worksheet3", "A1", 1, "ROUND(worksheet3!A1,1)")]
        [InlineData("worksheet3", "C4", 0, "ROUND(worksheet3!C4,0)")]
        [InlineData("worksheet3", "$A1", 10, "ROUND(worksheet3!$A1,10)")]
        [InlineData("worksheet3", "$A$1", 5, "ROUND(worksheet3!$A$1,5)")]
        [InlineData("worksheet3", "A1", -2, "ROUND(worksheet3!A1,-2)")] // This seems to be valid
        public void RoundTest2(string worksheetName, string addressExpression, int numberOfDecimals, string expectedFormula)
        {
            Worksheet worksheet = new Worksheet(worksheetName);
            Address address = new Address(addressExpression);
            Cell formula = BasicFormulas.Round(worksheet, address, numberOfDecimals);
            Assert.Equal(expectedFormula, formula.Value.ToString());
            Assert.Equal(Cell.CellType.FORMULA, formula.DataType);
        }

        [Theory(DisplayName = "Test of the Sum function on a Range object")]
        [InlineData("A1:A1", "SUM(A1)")]
        [InlineData("A1:C2", "SUM(A1:C2)")]
        [InlineData("$A1:C2", "SUM($A1:C2)")]
        [InlineData("$A$1:C2", "SUM($A$1:C2)")]
        [InlineData("$A$1:$C2", "SUM($A$1:$C2)")]
        [InlineData("$A$1:$C$2", "SUM($A$1:$C$2)")]
        public void SumTest(string rangeExpression, string expectedFormula)
        {
            Range range = new Range(rangeExpression);
            Cell formula = BasicFormulas.Sum(range);
            Assert.Equal(expectedFormula, formula.Value.ToString());
            Assert.Equal(Cell.CellType.FORMULA, formula.DataType);
        }

        [Theory(DisplayName = "Test of the Sum function on a Range object and a target worksheet")]
        [InlineData("worksheet1", "A1:A1", "SUM(worksheet1!A1)")]
        [InlineData("worksheet1", "A1:C2", "SUM(worksheet1!A1:C2)")]
        [InlineData("worksheet1", "$A1:C2", "SUM(worksheet1!$A1:C2)")]
        [InlineData("worksheet1", "$A$1:C2", "SUM(worksheet1!$A$1:C2)")]
        [InlineData("worksheet1", "$A$1:$C2", "SUM(worksheet1!$A$1:$C2)")]
        [InlineData("worksheet1", "$A$1:$C$2", "SUM(worksheet1!$A$1:$C$2)")]
        public void SumTest2(string worksheetName, string rangeExpression, string expectedFormula)
        {
            Worksheet worksheet = new Worksheet(worksheetName);
            Range range = new Range(rangeExpression);
            Cell formula = BasicFormulas.Sum(worksheet, range);
            Assert.Equal(expectedFormula, formula.Value.ToString());
            Assert.Equal(Cell.CellType.FORMULA, formula.DataType);
        }

        [Theory(DisplayName = "Test of the VLookup function on a Range object with an arbitrary number, the column index and the option of an exact match")]
        [InlineData(11, "A1:A1", 1, false, "VLOOKUP(11,A1:A1,1,FALSE)")]
        [InlineData(0.5f, "A1:C4", 3, false, "VLOOKUP(0.5,A1:C4,3,FALSE)")]
        [InlineData(-800L, "A10:XFD999999", 200, true, "VLOOKUP(-800,A10:XFD999999,200,TRUE)")]
        [InlineData(0, "X100:A1", 5, true, "VLOOKUP(0,A1:X100,5,TRUE)")]
        public void VLookupTest(object number, string rangeExpression, int columnIndex, bool exactMatch, string expectedFormula)
        {
            Range range = new Range(rangeExpression);
            Cell formula = BasicFormulas.VLookup(number, range, columnIndex, exactMatch);
            Assert.Equal(expectedFormula, formula.Value.ToString());
            Assert.Equal(Cell.CellType.FORMULA, formula.DataType);
        }

        [Theory(DisplayName = "Test of the VLookup function on a Range object with an arbitrary number, the column index, the option of an exact match and a target worksheet")]
        [InlineData("worksheet1", 11u, "$A$1:A1", 1, false, "VLOOKUP(11,worksheet1!$A$1:A1,1,FALSE)")]
        [InlineData("worksheet1", 0.5d, "A1:$C4", 3, false, "VLOOKUP(0.5,worksheet1!A1:$C4,3,FALSE)")]
        [InlineData("worksheet1", 2.22, "$A10:XFD999999", 200, true, "VLOOKUP(2.22,worksheet1!$A10:XFD999999,200,TRUE)")]
        [InlineData("worksheet1", 0, "X100:A1", 5, true, "VLOOKUP(0,worksheet1!A1:X100,5,TRUE)")]
        public void VLookupTest2(string worksheetName, object number, string rangeExpression, int columnIndex, bool exactMatch, string expectedFormula)
        {
            Worksheet worksheet = new Worksheet(worksheetName);
            Range range = new Range(rangeExpression);
            Cell formula = BasicFormulas.VLookup(number, worksheet, range, columnIndex, exactMatch);
            Assert.Equal(expectedFormula, formula.Value.ToString());
            Assert.Equal(Cell.CellType.FORMULA, formula.DataType);
        }

        [Theory(DisplayName = "Test of the VLookup function on a Range object with reference address, the column index and the option of an exact match")]
        [InlineData("C5", "A1:$A$1", 1, false, "VLOOKUP(C5,A1:$A$1,1,FALSE)")]
        [InlineData("A1", "A1:C$4", 3, false, "VLOOKUP(A1,A1:C$4,3,FALSE)")]
        [InlineData("$F4", "A10:XFD999999", 200, true, "VLOOKUP($F4,A10:XFD999999,200,TRUE)")]
        [InlineData("$XFD$99999", "X100:A1", 5, true, "VLOOKUP($XFD$99999,A1:X100,5,TRUE)")]
        public void VLookupTest3(string addressExpression, string rangeExpression, int columnIndex, bool exactMatch, string expectedFormula)
        {
            Address address = new Address(addressExpression);
            Range range = new Range(rangeExpression);
            Cell formula = BasicFormulas.VLookup(address, range, columnIndex, exactMatch);
            Assert.Equal(expectedFormula, formula.Value.ToString());
            Assert.Equal(Cell.CellType.FORMULA, formula.DataType);
        }

        [Theory(DisplayName = "Test of the VLookup function on a Range object with reference address, the column index, the option of an exact match and tow target worksheets")]
        [InlineData("worksheet1", "C5", "worksheet1", "A1:$A$1", 1, false, "VLOOKUP(worksheet1!C5,worksheet1!A1:$A$1,1,FALSE)")]
        [InlineData("worksheet2", "A1", "worksheet1", "A1:C$4", 3, false, "VLOOKUP(worksheet2!A1,worksheet1!A1:C$4,3,FALSE)")]
        [InlineData("worksheet1", "$F4", "worksheet2", "A10:XFD999999", 200, true, "VLOOKUP(worksheet1!$F4,worksheet2!A10:XFD999999,200,TRUE)")]
        [InlineData("worksheet2", "$XFD$99999", "worksheet2", "X100:A1", 5, true, "VLOOKUP(worksheet2!$XFD$99999,worksheet2!A1:X100,5,TRUE)")]
        public void VLookupTest4(string valueWorksheetName, string addressExpression, string rangesWorksheetName, string rangeExpression, int columnIndex, bool exactMatch, string expectedFormula)
        {
            Worksheet valueWorksheet = new Worksheet(valueWorksheetName);
            Worksheet rangeWorksheet = new Worksheet(rangesWorksheetName);
            Address address = new Address(addressExpression);
            Range range = new Range(rangeExpression);
            Cell formula = BasicFormulas.VLookup(valueWorksheet, address, rangeWorksheet, range, columnIndex, exactMatch);
            Assert.Equal(expectedFormula, formula.Value.ToString());
            Assert.Equal(Cell.CellType.FORMULA, formula.DataType);
        }

        [Fact(DisplayName = "Test of the VLookup function for byte as value")]
        public void VLookupByteTest()
        {
            AssertVlookup((byte)0, "0");
            AssertVlookup((byte)15, "15");
            AssertVlookup((byte)15, "15");
        }

        [Fact(DisplayName = "Test of the VLookup function for sbyte as value")]
        public void VLookupSbyteTest()
        {
            AssertVlookup((sbyte)0, "0");
            AssertVlookup((sbyte)-77, "-77");
            AssertVlookup((sbyte)77, "77");
        }

        [Fact(DisplayName = "Test of the VLookup function for decimal as value")]
        public void VLookupDecimalTest()
        {
            Decimal d1 = new decimal(0);
            Decimal d2 = new decimal(-0.005);
            Decimal d3 = new decimal(22.78);
            AssertVlookup(d1, "0");
            AssertVlookup(d2, "-0.005");
            AssertVlookup(d3, "22.78");
        }

        [Fact(DisplayName = "Test of the VLookup function for double as value")]
        public void VLookupDoubleTest()
        {
            AssertVlookup(0.0d, "0");
            AssertVlookup(222.5d, "222.5");
            AssertVlookup(-0.101d, "-0.101");
        }

        [Fact(DisplayName = "Test of the VLookup function for float as value")]
        public void VLookupFloatTest()
        {
            AssertVlookup(0.0f, "0");
            AssertVlookup(22.5f, "22.5");
            AssertVlookup(-0.01f, "-0.01");
        }

        [Fact(DisplayName = "Test of the VLookup function for int as value")]
        public void VLookupIntTest()
        {
            AssertVlookup((int)0, "0");
            AssertVlookup((int)-77, "-77");
            AssertVlookup((int)77, "77");
        }

        [Fact(DisplayName = "Test of the VLookup function for uint as value")]
        public void VLookupUintTest()
        {
            AssertVlookup((uint)0, "0");
            AssertVlookup((uint)999999, "999999");
        }

        [Fact(DisplayName = "Test of the VLookup function for long as value")]
        public void VLookupLongTest()
        {
            AssertVlookup(0L, "0");
            AssertVlookup(-999999L, "-999999");
            AssertVlookup(999999L, "999999");
        }

        [Fact(DisplayName = "Test of the VLookup function for ulong as value")]
        public void VLookupUlongTest()
        {
            AssertVlookup((ulong)0, "0");
            AssertVlookup((ulong)999999, "999999");
        }

        [Fact(DisplayName = "Test of the VLookup function for short as value")]
        public void VLookupShortTest()
        {
            AssertVlookup((short)0, "0");
            AssertVlookup((short)-128, "-128");
            AssertVlookup((short)255, "255");
        }

        [Fact(DisplayName = "Test of the VLookup function for ushort as value")]
        public void VLookupUshortTest()
        {
            AssertVlookup((ushort)0, "0");
            AssertVlookup((ushort)128, "128");
        }

        [Fact(DisplayName = "Test of the failing VLookup function on an invalid value type")]
        public void VLookupFailTest()
        {
            Range range = new Range("A1:D100");
            int column = 2;
            bool exactMatch = true;
            Assert.Throws<FormatException>(() => BasicFormulas.VLookup("test", range, column, exactMatch));
            Assert.Throws<FormatException>(() => BasicFormulas.VLookup(false, range, column, exactMatch));
            Assert.Throws<FormatException>(() => BasicFormulas.VLookup(null, range, column, exactMatch));
            Assert.Throws<FormatException>(() => BasicFormulas.VLookup(new DateTime(), range, column, exactMatch));
        }

        [Fact(DisplayName = "Test of the failing VLookup function on an invalid index column")]
        public void VLookupFailTest2()
        {
            Range range = new Range("A1:D100");
            Range range2 = new Range("C1:D100");
            Assert.Throws<FormatException>(() => BasicFormulas.VLookup(22, range, 0, true));
            Assert.Throws<FormatException>(() => BasicFormulas.VLookup(22, range, -2, false));
            Assert.Throws<FormatException>(() => BasicFormulas.VLookup(22, range, 100, true));
            Assert.Throws<FormatException>(() => BasicFormulas.VLookup(22, range2, 3, true));
            Assert.Throws<FormatException>(() => BasicFormulas.VLookup(22, range2, 4, false));
        }

        private void AssertVlookup(object number, string expectedLookupValue)
        {
            Range range = new Range("A1:D100");
            int column = 2;
            bool exactMatch = true;
            Cell formula = BasicFormulas.VLookup(number, range, column, exactMatch);
            string expectedFormula = "VLOOKUP(" + expectedLookupValue + "," + range.ToString() + "," + column.ToString() + "," + exactMatch.ToString().ToUpper() + ")";
            Assert.Equal(expectedFormula, formula.Value.ToString());
            Assert.Equal(Cell.CellType.FORMULA, formula.DataType);
        }

    }
}
