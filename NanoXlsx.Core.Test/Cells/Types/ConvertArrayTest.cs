using NanoXLSX;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace NanoXLSX.Test.Cells.Types
{
    public class ConvertArrayTest
    {
        [Fact(DisplayName = "Test of the ConvertArray method on bools")]
        public void ConvertBoolArrayTest()
        {
            bool[] array = new bool[] { true, true, false, true, false };
            AssertArray<bool>(array, typeof(bool));
        }

        [Fact(DisplayName = "Test of the ConvertArray method on bytes")]
        public void ConvertByteArrayTest()
        {
            byte[] array = new byte[] { 12, 55, 127, 0, 1, 255, byte.MinValue, byte.MaxValue };
            AssertArray<byte>(array, typeof(byte));
        }

        [Fact(DisplayName = "Test of the ConvertArray method on sbytes")]
        public void ConvertSByteArrayTest()
        {
            sbyte[] array = new sbyte[] { 12, 55, 127, -128, -1, 0, sbyte.MinValue, sbyte.MaxValue };
            AssertArray<sbyte>(array, typeof(sbyte));
        }

        [Fact(DisplayName = "Test of the ConvertArray method on decimal")]
        public void ConvertDecimaleArrayTest()
        {
            decimal[] array = new decimal[] { 0, 11.7m, 0.00001m, -22.5m, 100, -99, decimal.MinValue, decimal.MaxValue };
            AssertArray<decimal>(array, typeof(decimal));
        }

        [Fact(DisplayName = "Test of the ConvertArray method on double")]
        public void ConvertDoubleArrayTest()
        {
            double[] array = new double[] { 0, 11.7d, 0.00001d, -22.5d, 100, -99, double.MinValue, double.MaxValue };
            AssertArray<double>(array, typeof(double));
        }

        [Fact(DisplayName = "Test of the ConvertArray method on float")]
        public void ConvertFloatArrayTest()
        {
            float[] array = new float[] { 0, 11.7f, 0.00001f, -22.5f, 100, -99, float.MinValue, float.MaxValue };
            AssertArray<float>(array, typeof(float));
        }

        [Fact(DisplayName = "Test of the ConvertArray method on int")]
        public void ConvertIntArrayTest()
        {
            int[] array = new int[] { 12, 55, -1, 0, int.MaxValue, int.MinValue };
            AssertArray<int>(array, typeof(int));
        }

        [Fact(DisplayName = "Test of the ConvertArray method on uint")]
        public void ConvertUintArrayTest()
        {
            uint[] array = new uint[] { 12, 55, 777, 0, uint.MaxValue, uint.MinValue };
            AssertArray<uint>(array, typeof(uint));
        }

        [Fact(DisplayName = "Test of the ConvertArray method on long")]
        public void ConvertLongArrayTest()
        {
            long[] array = new long[] { 12, 55, -1, 0, long.MaxValue, long.MinValue };
            AssertArray<long>(array, typeof(long));
        }

        [Fact(DisplayName = "Test of the ConvertArray method on ulong")]
        public void ConvertULongArrayTest()
        {
            ulong[] array = new ulong[] { 12, 55, 777, 0, ulong.MaxValue, ulong.MinValue };
            AssertArray<ulong>(array, typeof(ulong));
        }

        [Fact(DisplayName = "Test of the ConvertArray method on short")]
        public void ConvertShortArrayTest()
        {
            short[] array = new short[] { 12, 55, -1, 0, short.MaxValue, short.MinValue };
            AssertArray<short>(array, typeof(short));
        }

        [Fact(DisplayName = "Test of the ConvertArray method on ushort")]
        public void ConvertUShortArrayTest()
        {
            ushort[] array = new ushort[] { 12, 55, 777, 0, ushort.MaxValue, ushort.MinValue };
            AssertArray<ushort>(array, typeof(ushort));
        }

        [Fact(DisplayName = "Test of the ConvertArray method on DateTime")]
        public void ConvertDateTimeArrayTest()
        {
            DateTime[] array = new DateTime[4];
            // Note: Dates before 1.1.1900 are not valid
            array[0] = new DateTime(1901, 01, 12, 12, 12, 12);
            array[1] = new DateTime(2200, 01, 12, 12, 12, 12);
            array[2] = new DateTime(2020, 11, 12);
            array[3] = new DateTime(1950, 5, 1);
            AssertArray<DateTime>(array, typeof(DateTime));
        }

        [Fact(DisplayName = "Test of the ConvertArray method on TimeSpan")]
        public void ConvertTimeSpanArrayTest()
        {
            TimeSpan[] array = new TimeSpan[4];
            array[0] = new TimeSpan(0);
            array[1] = new TimeSpan(12, 10, 50);
            array[2] = new TimeSpan(23, 59, 59);
            array[3] = new TimeSpan(11, 11, 11);
            AssertArray<TimeSpan>(array, typeof(TimeSpan));
        }

        [Fact(DisplayName = "Test of the ConvertArray method on nested Cell objects")]
        public void ConvertCellArrayTest()
        {
            Cell[] array = new Cell[4];
            array[0] = new Cell("", Cell.CellType.STRING);
            array[1] = new Cell("test", Cell.CellType.STRING);
            array[2] = new Cell("x", Cell.CellType.STRING);
            array[3] = new Cell(" ", Cell.CellType.STRING);
            AssertArray<Cell>(array, typeof(string), new string[] { "", "test", "x", " " });
        }

        [Fact(DisplayName = "Test of the ConvertArray method on string")]
        public void ConvertStringArrayTest()
        {
            string[] array = new string[] { "", "test", "X", "Ø", null, " " };

            AssertArray<string>(array, typeof(string));
        }

        [Fact(DisplayName = "Test of the ConvertArray method on other object types")]
        public void ConvertObjectArrayTest()
        {
            DummyArrayClass[] array = new DummyArrayClass[4];
            array[0] = new DummyArrayClass("");
            array[1] = new DummyArrayClass(null);
            array[2] = new DummyArrayClass(" ");
            array[3] = new DummyArrayClass("test");
            string[] actualValues = new string[array.Length];
            for (int i = 0; i < actualValues.Length; i++)
            {
                actualValues[i] = array[i].ToString();
            }
            AssertArray<DummyArrayClass>(array, typeof(string), actualValues);
        }

        [Fact(DisplayName = "Test of the ConvertArray method on null and empty arrays")]
        public void ConvertObjectArrayEmptyTest()
        {
            string[] nullArray = null;
            List<Cell> cells = Cell.ConvertArray<string>(nullArray).ToList();
            Assert.Empty(cells);

            string[] emptyArray = new string[0];
            List<Cell> cells2 = Cell.ConvertArray<string>(emptyArray).ToList();
            Assert.Empty(cells2);
        }

        private static void AssertArray<T>(T[] array, Type expectedValueType, object[] actualValues = null)
        {
            List<T> list = new List<T>();
            foreach (T obj in array)
            {
                list.Add(obj);
            }
            IEnumerable<Cell> ienumerable = Cell.ConvertArray<T>(list);
            List<Cell> cells = ienumerable.ToList();
            Assert.NotNull(cells);
            Assert.Equal(array.Length, cells.Count());
            for (int i = 0; i < array.Length; i++)
            {
                Cell cell = cells[i];
                if (cell.Value != null)
                {
                    Assert.Equal(expectedValueType, cell.Value.GetType());
                }
                if (actualValues == null)
                {
                    Assert.Equal(array[i], cell.Value);
                }
                else
                {
                    Assert.Equal(actualValues[i], cell.Value);
                }

            }
        }

        public class DummyArrayClass
        {
            public string Value { get; set; }

            public DummyArrayClass(string value)
            {
                this.Value = value;
            }

            public override string ToString()
            {
                return this.Value;
            }

        }

    }
}
