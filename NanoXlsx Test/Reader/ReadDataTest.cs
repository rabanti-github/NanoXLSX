using NanoXLSX;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace NanoXLSX_Test.Reader
{
    public class ReadDataTest
    {
        [Fact(DisplayName = "Test of the reader functionality for strings")]
        public void ReadStringTest()
        {
            Dictionary<string, string> cells = new Dictionary<string, string>();
            cells.Add("A1", "Test");
            cells.Add("B2", "22");
            cells.Add("C3", "");
            cells.Add("D4", " ");
            cells.Add("E4", "x ");
            cells.Add("F4", " X");
            cells.Add("G4", " x ");
            cells.Add("H4", "x x");
            cells.Add("E5", "#@+-\"'?!\\(){}[]<>/|.,;:");
            cells.Add("L6", "\t");
            cells.Add("M6", "\tx");
            cells.Add("N6", "x\t");
            cells.Add("E7", "日本語");
            cells.Add("F7", "हिन्दी");
            cells.Add("G7", "한국어");
            cells.Add("H7", "官話");
            cells.Add("I7", "ελληνική γλώσσα");
            cells.Add("J7", "русский язык");
            cells.Add("K7", "עברית");
            cells.Add("L7", "اَلْعَرَبِيَّة");
            AssertValues<String>(cells, AssertEqals);
        }

        [Fact(DisplayName = "Test of the reader functionality for new lines in strings")]
        public void ReadStringNewLineTest()
        {
            Dictionary<string, string> given = new Dictionary<string, string>();
            given.Add("A1", "\r");
            given.Add("A2", "\n");
            given.Add("A3", "\r\n");
            given.Add("A4", "a\n");
            given.Add("A5", "\nx");
            given.Add("A6", "a\r");
            given.Add("A7", "\rx");
            given.Add("A8", "a\r\n");
            given.Add("A9", "\r\nx");
            given.Add("A10", "\n\n\n");
            given.Add("A11", "\r\r\r");
            given.Add("A12", "\n\r"); // irregular use
            Dictionary<string, string> expected = new Dictionary<string, string>();
            expected.Add("A1", "\r\n");
            expected.Add("A2", "\r\n");
            expected.Add("A3", "\r\n");
            expected.Add("A4", "a\r\n");
            expected.Add("A5", "\r\nx");
            expected.Add("A6", "a\r\n");
            expected.Add("A7", "\r\nx");
            expected.Add("A8", "a\r\n");
            expected.Add("A9", "\r\nx");
            expected.Add("A10", "\r\n\r\n\r\n");
            expected.Add("A11", "\r\n\r\n\r\n");
            expected.Add("A12", "\r\n\r\n");
            AssertValues<String>(given, AssertEqals, expected);
        }

        [Fact(DisplayName = "Test of the reader functionality for null / empty values")]
        public void ReadNullTest()
        {
            Dictionary<string, object> cells = new Dictionary<string, object>();
            cells.Add("A1", null);
            cells.Add("A2", null);
            cells.Add("A3", null);
            AssertValues<object>(cells, AssertEqals);
        }

        [Fact(DisplayName = "Test of the reader functionality for long values (above int32 and uint32 range)")]
        public void ReadLongTest()
        {
            Dictionary<string, long> cells = new Dictionary<string, long>();
            cells.Add("A1", 4294967296);
            cells.Add("A2", -2147483649);
            cells.Add("A3", 21474836480);
            cells.Add("A4", -21474836480);
            cells.Add("A5", long.MinValue);
            cells.Add("A6", long.MaxValue);
            AssertValues<long>(cells, AssertEqals);
        }

        [Fact(DisplayName = "Test of the reader functionality for ulong values (above signed int64 range)")]
        public void ReadUlongTest()
        {
            Dictionary<string, ulong> cells = new Dictionary<string, ulong>();
            long lmax = long.MaxValue;
            cells.Add("A1", (ulong)(lmax + 1));
            cells.Add("A2", (ulong)(lmax + 9999));
            cells.Add("A3", ulong.MaxValue);
            AssertValues<ulong>(cells, AssertEqals);
        }

        [Fact(DisplayName = "Test of the reader functionality for int values")]
        public void ReadIntTest()
        {
            Dictionary<string, int> cells = new Dictionary<string, int>();
            cells.Add("A1", 0);
            cells.Add("A2", 10);
            cells.Add("A3", -10);
            cells.Add("A4", 999999);
            cells.Add("A5", -999999);
            cells.Add("A6", int.MinValue);
            cells.Add("A7", int.MaxValue);
            AssertValues<int>(cells, AssertEqals);
        }

        [Fact(DisplayName = "Test of the reader functionality for uint values (above signed int32 range)")]
        public void ReadUintTest()
        {
            Dictionary<string, uint> cells = new Dictionary<string, uint>();
            uint imax = int.MaxValue;
            cells.Add("A1", (uint)(imax + 1));
            cells.Add("A2", (uint)(imax + 9999));
            cells.Add("A3", uint.MaxValue);
            AssertValues<uint>(cells, AssertEqals);
        }

        [Fact(DisplayName = "Test of the reader functionality for float values")]
        public void ReadFloatTest()
        {
            // Numbers without fraction elements are always interpreted as float
            Dictionary<string, float> cells = new Dictionary<string, float>();
            cells.Add("A1", 0.000001f);
            cells.Add("A2", 10.1f);
            cells.Add("A3", -10.22f);
            cells.Add("A4", 999999.9f);
            cells.Add("A5", -999999.9f);
            cells.Add("A6", float.MinValue);
            cells.Add("A7", float.MaxValue);
            AssertValues<float>(cells, AssertAppriximateFloat);
        }

        [Fact(DisplayName = "Test of the reader functionality for double values (above single32 range)")]
        public void ReadDoubleTest()
        {
            Dictionary<string, double> cells = new Dictionary<string, double>();
            cells.Add("A1", 440282346700000000000000000000000000009.1d);
            cells.Add("A2", -440282347600000000000000000000000000009.1d);
            cells.Add("A3", 21474836480648356436538453467583788456343865.227d);
            cells.Add("A4", -21474836480648356436538453467583748856343865.9d);
            cells.Add("A5", double.MinValue);
            cells.Add("A6", double.MaxValue);
            AssertValues<double>(cells, AssertAppriximateDouble);
        }

        [Fact(DisplayName = "Test of the reader functionality for bool values")]
        public void ReadBoolTest()
        {
            Dictionary<string, bool> cells = new Dictionary<string, bool>();
            cells.Add("A1", true);
            cells.Add("A2", false);
            cells.Add("A3", true);
            AssertValues<bool>(cells, AssertEqals);
        }

        [Fact(DisplayName = "Test of the reader functionality for DateTime values")]
        public void ReadDateTimeTest()
        {
            Dictionary<string, DateTime> cells = new Dictionary<string, DateTime>();
            cells.Add("A1", new DateTime(2021, 5, 11, 15, 7, 2));
            cells.Add("A2", new DateTime(1900, 1, 1, 0, 0, 0));
            cells.Add("A3", new DateTime(1960, 12, 12));
            cells.Add("A4", new DateTime(9999, 12, 31, 23, 59, 59));
            AssertValues<DateTime>(cells, AssertEqals);
        }

        [Fact(DisplayName = "Test of the reader functionality for TimeSpan values")]
        public void ReadTimeSpanTest()
        {
            Dictionary<string, TimeSpan> cells = new Dictionary<string, TimeSpan>();
            cells.Add("A1", new TimeSpan(0,0,0));
            cells.Add("A2", new TimeSpan(13,18,22));
            cells.Add("A3", new TimeSpan(12,0,0));
            cells.Add("A4", new TimeSpan(23,59,59));
            AssertValues<TimeSpan>(cells, AssertEqals);
        }

        [Fact(DisplayName = "Test of the reader functionality on invalid / unexpected values")]
        public void ReadInvalidDataTest()
        {
            Stream stream = TestUtils.GetResource("tampered.xlsx");
            Workbook workbook = Workbook.Load(stream);
            int i = 0;
        }


        private static void AssertEqals<T>(T expected, T given)
        {
            Assert.Equal(expected, given);
        }

        private static void AssertValues<T>(Dictionary<string, T> givenCells, Action<T,T> asserionAction, Dictionary<string, T> expectedCells = null)
        {
            Workbook workbook = new Workbook("worksheet1");
            foreach (KeyValuePair<string, T> cell in givenCells)
            {
                workbook.CurrentWorksheet.AddCell(cell.Value, cell.Key);
               // Cell c = new Cell(1.8e+309, Cell.CellType.NUMBER);
            }
            MemoryStream stream = new MemoryStream();
            workbook.SaveAsStream(stream, true);
            workbook.SaveAs("C:\\purge-temp\\debug.xlsx");
            stream.Position = 0;
            Workbook givenWorkbook = Workbook.Load(stream);

            Assert.NotNull(givenWorkbook);
            Worksheet givenWorksheet = givenWorkbook.SetCurrentWorksheet(0);
            Assert.Equal("worksheet1", givenWorksheet.SheetName);
            foreach(string address in givenCells.Keys)
            {
                Cell givenCell = givenWorksheet.GetCell(new Address(address));
                T value;
                if (expectedCells == null)
                {
                    value = givenCells[address];
                }
                else
                {
                    value = expectedCells[address];
                }
                
                if (value == null)
                {
                    Assert.Equal(Cell.CellType.EMPTY, givenCell.DataType);
                }
                else
                {
                    asserionAction.Invoke(value, (T)givenCell.Value);
                    //Assert.Equal(value, givenCell.Value);
                }
            }
        }

        private static void AssertAppriximateDouble(double expected, double given)
        {
            double threshold = 0.00000001;
            Assert.True(Math.Abs(given - expected) < threshold);
        }

        private static void AssertAppriximateFloat(float expected, float given)
        {
            float threshold = 0.00000001f;
            Assert.True(Math.Abs(given - expected) < threshold);
        }


    }
}
