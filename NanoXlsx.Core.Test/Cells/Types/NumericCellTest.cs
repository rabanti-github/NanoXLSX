using System;
using NanoXLSX.Styles;
using NanoXLSX.Test.Core.Utils;
using Xunit;
using static NanoXLSX.Cell;

namespace NanoXLSX.Test.Core.Cells.Types
{
    // Ensure that these tests are executed sequentially, since static repository methods may be called 
    [Collection(nameof(SequentialCollection))]
    public class NumericCellTest
    {

        CellTypeUtils utils;

        public NumericCellTest()
        {
            utils = new CellTypeUtils();
        }


        [Theory(DisplayName = "Byte value cell test: Test of the cell values, as well as proper modification")]
        [InlineData(0)]
        [InlineData(16)]
        [InlineData(byte.MinValue)]
        [InlineData(byte.MaxValue)]
        public void ByteCellTest(byte value)
        {
            utils.AssertCellCreation<byte>(8, value, CellType.NUMBER, (current, other) => { return current.Equals(other); });
        }

        [Fact(DisplayName = "Byte value cell test with style")]
        public void ByteCellTest2()
        {
            Style style = BasicStyles.Italic;
            utils.AssertStyledCellCreation<byte>(0, 8, CellType.NUMBER, (current, other) => { return current.Equals(other); }, style);
        }

        [Theory(DisplayName = "Signed Byte value cell test: Test of the cell values, as well as proper modification")]
        [InlineData(0)]
        [InlineData(-22)]
        [InlineData(sbyte.MinValue)]
        [InlineData(sbyte.MaxValue)]
        public void SByteCellTest(sbyte value)
        {
            utils.AssertCellCreation<sbyte>(8, value, CellType.NUMBER, (current, other) => { return current.Equals(other); });
        }

        [Theory(DisplayName = "Short value cell test: Test of the cell values, as well as proper modification")]
        [InlineData(0)]
        [InlineData(-127)]
        [InlineData(short.MinValue)]
        [InlineData(short.MaxValue)]
        public void ShortCellTest(short value)
        {
            utils.AssertCellCreation<short>(-6, value, CellType.NUMBER, (current, other) => { return current.Equals(other); });
        }

        [Theory(DisplayName = "Unsigned Short value cell test: Test of the cell values, as well as proper modification")]
        [InlineData(0)]
        [InlineData(127)]
        [InlineData(ushort.MinValue)]
        [InlineData(ushort.MaxValue)]
        public void UShortCellTest(ushort value)
        {
            utils.AssertCellCreation<ushort>(3398, value, CellType.NUMBER, (current, other) => { return current.Equals(other); });
        }

        [Theory(DisplayName = "Int value cell test: Test of the cell values, as well as proper modification")]
        [InlineData(0)]
        [InlineData(-42)]
        [InlineData(int.MinValue)]
        [InlineData(int.MaxValue)]
        public void IntCellTest(int value)
        {
            utils.AssertCellCreation<int>(99, value, CellType.NUMBER, (current, other) => { return current.Equals(other); });
        }

        [Theory(DisplayName = "Unsigned Int value cell test: Test of the cell values, as well as proper modification")]
        [InlineData(0)]
        [InlineData(42)]
        [InlineData(uint.MinValue)]
        [InlineData(uint.MaxValue)]
        public void UIntCellTest(uint value)
        {
            utils.AssertCellCreation<uint>(98, value, CellType.NUMBER, (current, other) => { return current.Equals(other); });
        }

        [Theory(DisplayName = "Long value cell test: Test of the cell values, as well as proper modification")]
        [InlineData(0L)]
        [InlineData(-999999999L)]
        [InlineData(long.MinValue)]
        [InlineData(long.MaxValue)]
        public void LongCellTest(long value)
        {
            utils.AssertCellCreation<long>(-6, value, CellType.NUMBER, (current, other) => { return current.Equals(other); });
        }

        [Theory(DisplayName = "Unsigned Long value cell test: Test of the cell values, as well as proper modification")]
        [InlineData(0)]
        [InlineData(55555)]
        [InlineData(ulong.MinValue)]
        [InlineData(ulong.MaxValue)]
        public void ULongCellTest(ulong value)
        {
            utils.AssertCellCreation<ulong>(99, value, CellType.NUMBER, (current, other) => { return current.Equals(other); });
        }

        [Fact(DisplayName = "Decimal value cell test: Test of the cell values, as well as proper modification")]
        public void DecimalCellTest()
        {
            // foalt.MinValue and float.MaxValue are not constants. Test must be hardcoded
            utils.AssertCellCreation<decimal>(-2.338m, 0, CellType.NUMBER, CompareDecimal);
            utils.AssertCellCreation<decimal>(-2.338m, -0.0057m, CellType.NUMBER, CompareDecimal);
            utils.AssertCellCreation<decimal>(-2.338m, decimal.MinValue, CellType.NUMBER, CompareDecimal);
            utils.AssertCellCreation<decimal>(-2.338m, decimal.MaxValue, CellType.NUMBER, CompareDecimal);
        }


        [Theory(DisplayName = "Float value cell test: Test of the cell values, as well as proper modification")]
        [InlineData(0f)]
        [InlineData(779.254f)]
        [InlineData(float.MinValue)]
        [InlineData(float.MaxValue)]
        public void FloatCellTest(float value)
        {
            utils.AssertCellCreation<float>(-2.338f, value, CellType.NUMBER, CompareFloat);
        }

        [Theory(DisplayName = "Double value cell test: Test of the cell values, as well as proper modification")]
        [InlineData(0d)]
        [InlineData(1.22d)]
        [InlineData(double.MinValue)]
        [InlineData(double.MaxValue)]
        public void DoubleCellTest(double value)
        {
            utils.AssertCellCreation<double>(42.778, value, CellType.NUMBER, CompareDouble);
        }

        [Theory(DisplayName = "Test of the byte comparison method on cells")]
        [InlineData(42, 42, 0)]
        [InlineData(100, 24, 1)]
        [InlineData(0, 127, -1)]
        public void ByteCellComparisonTest(byte value1, byte value2, int expectedResult)
        {
            AssertNumericType<byte>(value1, value2, expectedResult);
        }

        [Theory(DisplayName = "Test of the sbyte comparison method on cells")]
        [InlineData(42, 42, 0)]
        [InlineData(100, -20, 1)]
        [InlineData(-127, 127, -1)]
        public void SbyteCellComparisonTest(sbyte value1, sbyte value2, int expectedResult)
        {
            AssertNumericType<sbyte>(value1, value2, expectedResult);
        }

        [Theory(DisplayName = "Test of the int comparison method on cells")]
        [InlineData(42, 42, 0)]
        [InlineData(9999, -999999, 1)]
        [InlineData(0, 18720, -1)]
        public void IntCellComparisonTest(int value1, int value2, int expectedResult)
        {
            AssertNumericType<int>(value1, value2, expectedResult);
        }

        [Theory(DisplayName = "Test of the uint comparison method on cells")]
        [InlineData(42, 42, 0)]
        [InlineData(10000, 2004, 1)]
        [InlineData(5000, 12700, -1)]
        public void UintCellComparisonTest(uint value1, uint value2, int expectedResult)
        {
            AssertNumericType<uint>(value1, value2, expectedResult);
        }


        [Theory(DisplayName = "Test of the long comparison method on cells")]
        [InlineData(42, 42, 0)]
        [InlineData(9999, -999999, 1)]
        [InlineData(0, 18720, -1)]
        public void LongCellComparisonTest(long value1, long value2, int expectedResult)
        {
            AssertNumericType<long>(value1, value2, expectedResult);
        }

        [Theory(DisplayName = "Test of the ulong comparison method on cells")]
        [InlineData(42, 42, 0)]
        [InlineData(10000, 2004, 1)]
        [InlineData(5000, 12700, -1)]
        public void UlongCellComparisonTest(ulong value1, ulong value2, int expectedResult)
        {
            AssertNumericType<ulong>(value1, value2, expectedResult);
        }

        [Theory(DisplayName = "Test of the short comparison method on cells")]
        [InlineData(42, 42, 0)]
        [InlineData(9999, -9999, 1)]
        [InlineData(0, 18720, -1)]
        public void ShortCellComparisonTest(short value1, short value2, int expectedResult)
        {
            AssertNumericType<long>(value1, value2, expectedResult);
        }

        [Theory(DisplayName = "Test of the ushort comparison method on cells")]
        [InlineData(42, 42, 0)]
        [InlineData(1000, 204, 1)]
        [InlineData(500, 1200, -1)]
        public void UshortCellComparisonTest(ushort value1, ushort value2, int expectedResult)
        {
            AssertNumericType<ulong>(value1, value2, expectedResult);
        }

        [Theory(DisplayName = "Test of the decimal comparison method on cells")]
        [InlineData(22.223, 22.223, 0)]
        [InlineData(9999.1, -0.001, 1)]
        [InlineData(0.23, 18720.0, -1)]
        public void DecimalCellComparisonTest(decimal value1, decimal value2, int expectedResult)
        {
            AssertNumericType<decimal>(value1, value2, expectedResult);
        }

        [Theory(DisplayName = "Test of the double comparison method on cells")]
        [InlineData(22.223, 22.223, 0)]
        [InlineData(9999.1, -0.001, 1)]
        [InlineData(0.23, 18720.0, -1)]
        public void DoubleCellComparisonTest(double value1, double value2, int expectedResult)
        {
            AssertNumericType<double>(value1, value2, expectedResult);
        }

        [Theory(DisplayName = "Test of the float comparison method on cells")]
        [InlineData(22.223, 22.223, 0)]
        [InlineData(9999.1, -0.001, 1)]
        [InlineData(0.23, 18720.0, -1)]
        public void FloatCellComparisonTest(float value1, float value2, int expectedResult)
        {
            AssertNumericType<float>(value1, value2, expectedResult);
        }

        private void AssertNumericType<T>(T value1, T value2, int expectedResult) where T : IComparable<T>
        {
            Cell cell1 = utils.CreateVariantCell(value1, utils.CellAddress);
            Cell cell2 = utils.CreateVariantCell(value2, utils.CellAddress);
            int comparison = ((T)cell1.Value).CompareTo((T)cell2.Value);
            Assert.True(VariantCompareTo(comparison, expectedResult));
        }


        #region comparers

        private static bool VariantCompareTo(int comparison, int expectedResult)
        {
            if (comparison == expectedResult)
            {
                return true;
            }
            if (comparison < 0 && expectedResult < 0)
            {
                return true;
            }
            if (comparison > 0 && expectedResult > 0)
            {
                return true;
            }
            return false;
        }

        private static bool CompareDouble(double current, double other)
        {
            const double threshold = 0.0000001;
            if (Math.Abs(current - other) < threshold)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private static bool CompareDecimal(decimal current, decimal other)
        {
            const decimal threshold = 0.0000001m;
            if (Math.Abs(current - other) < threshold)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private static bool CompareFloat(float current, float other)
        {
            const float threshold = 0.0000001f;
            if (Math.Abs(current - other) < threshold)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        #endregion
    }
}
