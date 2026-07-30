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
    }
}
