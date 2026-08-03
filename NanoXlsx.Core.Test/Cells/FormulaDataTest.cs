using System;
using NanoXLSX.Enums;
using Xunit;

namespace NanoXLSX.Test.Core.CellTest
{
    public class FormulaDataTest
    {
        [Fact(DisplayName = "Test of the FormulaData default constructor")]
        public void FormulaDataConstructorTest()
        {
            FormulaData data = new FormulaData();

            Assert.Equal(FormulaData.FormulaType.Normal, data.Type);
            Assert.Equal(Cell.CellType.Default, data.CachedValueType);
            Assert.Null(data.CachedValue);
        }

        [Theory(DisplayName = "Test of cached value type inference in the FormulaData constructor")]
        [InlineData(null, Cell.CellType.Default)]
        [InlineData("0", Cell.CellType.String)]
        [InlineData('x', Cell.CellType.String)]
        [InlineData(true, Cell.CellType.Bool)]
        [InlineData((byte)1, Cell.CellType.Number)]
        [InlineData((sbyte)-1, Cell.CellType.Number)]
        [InlineData((short)-2, Cell.CellType.Number)]
        [InlineData((ushort)2, Cell.CellType.Number)]
        [InlineData(-3, Cell.CellType.Number)]
        [InlineData((uint)3, Cell.CellType.Number)]
        [InlineData((long)-4, Cell.CellType.Number)]
        [InlineData((ulong)4, Cell.CellType.Number)]
        [InlineData(1.25f, Cell.CellType.Number)]
        [InlineData(2.5d, Cell.CellType.Number)]
        public void FormulaDataConstructorTest2(object givenCachedValue, Cell.CellType expectedType)
        {
            FormulaData data = new FormulaData("A1", givenCachedValue);

            Assert.Equal("A1", data.Expression);
            Assert.Equal(givenCachedValue, data.CachedValue);
            Assert.Equal(expectedType, data.CachedValueType);
        }

        [Fact(DisplayName = "Test of special cached value type inference in the FormulaData constructor")]
        public void FormulaDataConstructorTest3()
        {
            Assert.Equal(Cell.CellType.Number, new FormulaData("A1", 1.25m).CachedValueType);
            Assert.Equal(Cell.CellType.Date, new FormulaData("A1", new DateTime(2026, 7, 30)).CachedValueType);
            Assert.Equal(Cell.CellType.Time, new FormulaData("A1", TimeSpan.FromHours(2)).CachedValueType);
            Assert.Equal(Cell.CellType.Error, new FormulaData("A1", Errors.FormulaError.DivisionByZero).CachedValueType);
            Assert.Equal(Cell.CellType.String, new FormulaData("A1", new object()).CachedValueType);
        }

        [Theory(DisplayName = "Test of external workbook reference detection in formulas")]
        [InlineData("[1]Sheet1!A1")]
        [InlineData("SUM([12]Sheet_Name!$A$1)")]
        [InlineData("'[1]Sheet 1'!$A$1")]
        [InlineData("[Book.xlsx]Sheet1!A1")]
        [InlineData("SUM('[Book.xlsx]Owner''s Sheet'!A1)")]
        [InlineData("[1]Sheet1!A1+[2]Sheet2!B2")]
        [InlineData("'..\\[Book.xlsx]Sheet1'!$A$1")]
        [InlineData("'../[Book.xlsx]Sheet1'!$A$1")]
        [InlineData("'C:\\temp\\[book one.xlsx]Sheet 1'!$A$1")]
        [InlineData("SUM('C:\\temp\\[book one.xlsx]Sheet 1'!$A$1,'..\\[other.xlsx]Data'!$B$2)")]
        public void ExternalReferenceDetectionTest(string expression)
        {
            FormulaData data = new FormulaData(expression);

            Assert.True(data.HasExternalReferences);
            Assert.True(FormulaData.ContainsExternalReference(expression));
        }

        [Theory(DisplayName = "Test of expressions without external workbook references")]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("SUM(A1:A2)")]
        [InlineData("Table1[Column]")]
        [InlineData("Table1[1]")]
        [InlineData("R[1]C[1]")]
        [InlineData("[")]
        [InlineData("[]Sheet1!A1")]
        [InlineData("[1]")]
        [InlineData("[1]!A1")]
        [InlineData("[1]Sheet1+A1")]
        [InlineData("Table1[Column]+Sheet1!A1")]
        [InlineData("\"[1]Sheet1!A1\"")]
        [InlineData("INDIRECT(\"[1]Sheet1!A1\")")]
        [InlineData("\"escaped \"\"[1]Sheet1!A1\"\" text\"")]
        public void ExternalReferenceDetectionNegativeTest(string expression)
        {
            FormulaData data = new FormulaData(expression);

            Assert.False(data.HasExternalReferences);
            Assert.False(FormulaData.ContainsExternalReference(expression));
        }

        [Fact(DisplayName = "Test that external workbook reference detection follows expression changes")]
        public void ExternalReferenceExpressionChangeTest()
        {
            FormulaData data = new FormulaData("A1");
            Assert.False(data.HasExternalReferences);

            data.Expression = "[1]Sheet1!A1";
            Assert.True(data.HasExternalReferences);

            data.Expression = "Table1[Column]";
            Assert.False(data.HasExternalReferences);
        }

        [Fact(DisplayName = "Test that copying FormulaData preserves external workbook reference detection")]
        public void ExternalReferenceCopyTest()
        {
            FormulaData data = new FormulaData("[1]Sheet1!A1");

            FormulaData copy = data.Copy();

            Assert.True(copy.HasExternalReferences);
            Assert.Equal(data, copy);
        }

        [Fact(DisplayName = "Test of the FormulaData Copy function with cached value metadata")]
        public void CopyTest()
        {
            FormulaData data = new FormulaData("A1", "0")
            {
                CachedValueType = Cell.CellType.Number,
                FormulaRange = "A1:A2",
                MasterCellAddress = "A1",
                Type = FormulaData.FormulaType.Array
            };

            FormulaData copy = data.Copy();

            Assert.NotSame(data, copy);
            Assert.Equal(data, copy);
            Assert.Equal(Cell.CellType.Number, copy.CachedValueType);
        }

        [Fact(DisplayName = "Test of FormulaData equality, comparison and hashing with cached value metadata")]
        public void CachedValueTypeComparisonTest()
        {
            FormulaData number = new FormulaData("A1", "0") { CachedValueType = Cell.CellType.Number };
            FormulaData numberCopy = number.Copy();
            FormulaData text = new FormulaData("A1", "0") { CachedValueType = Cell.CellType.String };

            Assert.True(number.Equals(numberCopy));
            Assert.Equal(0, number.CompareTo(numberCopy));
            Assert.Equal(number.GetHashCode(), numberCopy.GetHashCode());
            Assert.False(number.Equals(text));
            Assert.NotEqual(0, number.CompareTo(text));
            Assert.NotEqual(number.GetHashCode(), text.GetHashCode());
        }

        [Fact(DisplayName = "Test of the FormulaData CompareTo method")]
        public void CompareToTest()
        {
            FormulaData data = CreateFormulaData();

            Assert.Equal(1, data.CompareTo(null));
            Assert.Equal(0, data.CompareTo(data.Copy()));
            Assert.NotEqual(0, data.CompareTo(CreateFormulaData(expression: "B1")));
            Assert.NotEqual(0, data.CompareTo(CreateFormulaData(type: FormulaData.FormulaType.Shared)));
            Assert.NotEqual(0, data.CompareTo(CreateFormulaData(formulaRange: "A1:A3")));
            Assert.NotEqual(0, data.CompareTo(CreateFormulaData(definedName: CreateDefinedName("OtherName"))));
            Assert.NotEqual(0, data.CompareTo(CreateFormulaData(cachedValueType: Cell.CellType.String)));
            Assert.NotEqual(0, data.CompareTo(CreateFormulaData(cachedValue: 2)));
            Assert.NotEqual(0, data.CompareTo(CreateFormulaData(masterCellAddress: "A2")));
        }

        [Fact(DisplayName = "Test of the strongly typed FormulaData Equals method")]
        public void EqualsFormulaDataTest()
        {
            FormulaData data = CreateFormulaData();

            Assert.False(data.Equals((FormulaData)null));
            Assert.True(data.Equals(data));
            Assert.True(data.Equals(data.Copy()));
            Assert.False(data.Equals(CreateFormulaData(expression: "B1")));
            Assert.False(data.Equals(CreateFormulaData(type: FormulaData.FormulaType.Shared)));
            Assert.False(data.Equals(CreateFormulaData(formulaRange: "A1:A3")));
            Assert.False(data.Equals(CreateFormulaData(definedName: CreateDefinedName("OtherName"))));
            Assert.False(data.Equals(CreateFormulaData(cachedValue: 2)));
            Assert.False(data.Equals(CreateFormulaData(cachedValueType: Cell.CellType.String)));
            Assert.False(data.Equals(CreateFormulaData(masterCellAddress: "A2")));
        }

        [Fact(DisplayName = "Test of the object FormulaData Equals method")]
        public void EqualsObjectTest()
        {
            FormulaData data = CreateFormulaData();

            Assert.True(data.Equals((object)data.Copy()));
            Assert.False(data.Equals((object)null));
            Assert.False(data.Equals("Wrong type"));
        }

        [Fact(DisplayName = "Test of the FormulaData GetHashCode method")]
        public void GetHashCodeTest()
        {
            FormulaData data = CreateFormulaData();
            FormulaData copy = data.Copy();
            FormulaData empty = new FormulaData();

            Assert.Equal(data.GetHashCode(), copy.GetHashCode());
            Assert.NotEqual(data.GetHashCode(), empty.GetHashCode());
            Assert.Equal(empty.GetHashCode(), new FormulaData().GetHashCode());
        }

        private static FormulaData CreateFormulaData(
            string expression = "A1",
            FormulaData.FormulaType type = FormulaData.FormulaType.Array,
            string formulaRange = "A1:A2",
            DefinedName definedName = null,
            object cachedValue = null,
            Cell.CellType cachedValueType = Cell.CellType.Number,
            string masterCellAddress = "A1")
        {
            return new FormulaData(expression, cachedValue ?? 1)
            {
                Type = type,
                FormulaRange = formulaRange,
                DefinedNameReference = definedName ?? CreateDefinedName("FormulaName"),
                CachedValueType = cachedValueType,
                MasterCellAddress = masterCellAddress
            };
        }

        private static DefinedName CreateDefinedName(string name)
        {
            return new DefinedName(new Workbook(), DefinedName.NameType.Constant, name, 1, null);
        }
    }
}
