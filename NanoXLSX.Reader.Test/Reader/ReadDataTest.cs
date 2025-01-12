using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using NanoXLSX;
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
            AssertValues<string>(cells, AssertEquals);
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
            AssertValues<string>(given, AssertEquals, expected);
        }

        [Fact(DisplayName = "Test of the reader functionality for null / empty values")]
        public void ReadNullTest()
        {
            Dictionary<string, object> cells = new Dictionary<string, object>();
            cells.Add("A1", null);
            cells.Add("A2", null);
            cells.Add("A3", null);
            AssertValues<object>(cells, AssertEquals);
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
            AssertValues<long>(cells, AssertEquals);
        }

        [Fact(DisplayName = "Test of the reader functionality for ulong values (above signed int64 range)")]
        public void ReadUlongTest()
        {
            Dictionary<string, ulong> cells = new Dictionary<string, ulong>();
            long lmax = long.MaxValue;
#pragma warning disable S3949 // Calculations should not overflow
            cells.Add("A1", (ulong)(lmax + 1));
#pragma warning disable S3949 // Calculations should not overflow
            cells.Add("A2", (ulong)(lmax + 9999));
#pragma warning restore S3949 // Calculations should not overflow
            cells.Add("A3", ulong.MaxValue);
            AssertValues<ulong>(cells, AssertEquals);
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
            AssertValues<int>(cells, AssertEquals);
        }

        [Fact(DisplayName = "Test of the reader functionality for uint values (above signed int32 range)")]
        public void ReadUintTest()
        {
            Dictionary<string, uint> cells = new Dictionary<string, uint>();
            uint imax = int.MaxValue;
            cells.Add("A1", imax + 1);
            cells.Add("A2", imax + 9999);
            cells.Add("A3", uint.MaxValue);
            AssertValues<uint>(cells, AssertEquals);
        }

        [Fact(DisplayName = "Test of the reader functionality for byte values (cast to int)")]
        public void ReadByteTest()
        {
            Dictionary<string, byte> cells = new Dictionary<string, byte>();
            cells.Add("A1", 0);
            cells.Add("A2", 10);
            cells.Add("A3", 255);
            cells.Add("A4", byte.MinValue);
            cells.Add("A5", byte.MaxValue);

            Dictionary<string, int> expected = new Dictionary<string, int>();
            expected.Add("A1", 0);
            expected.Add("A2", 10);
            expected.Add("A3", 255);
            expected.Add("A4", byte.MinValue);
            expected.Add("A5", byte.MaxValue);
            AssertValues<byte, int>(cells, AssertEquals, expected);
        }

        [Fact(DisplayName = "Test of the reader functionality for sbyte values (cast to int)")]
        public void ReadSbyteTest()
        {
            Dictionary<string, sbyte> cells = new Dictionary<string, sbyte>();
            cells.Add("A1", 0);
            cells.Add("A2", 10);
            cells.Add("A3", -10);
            cells.Add("A4", 127);
            cells.Add("A5", -128);
            cells.Add("A6", sbyte.MinValue);
            cells.Add("A7", sbyte.MaxValue);

            Dictionary<string, int> expected = new Dictionary<string, int>();
            expected.Add("A1", 0);
            expected.Add("A2", 10);
            expected.Add("A3", -10);
            expected.Add("A4", 127);
            expected.Add("A5", -128);
            expected.Add("A6", sbyte.MinValue);
            expected.Add("A7", sbyte.MaxValue);
            AssertValues<sbyte, int>(cells, AssertEquals, expected);
        }

        [Fact(DisplayName = "Test of the reader functionality for short values (cast to int)")]
        public void ReadShortTest()
        {
            Dictionary<string, short> cells = new Dictionary<string, short>();
            cells.Add("A1", 0);
            cells.Add("A2", 10);
            cells.Add("A3", 32767);
            cells.Add("A4", -32767);
            cells.Add("A5", short.MinValue);
            cells.Add("A6", short.MaxValue);

            Dictionary<string, int> expected = new Dictionary<string, int>();
            expected.Add("A1", 0);
            expected.Add("A2", 10);
            expected.Add("A3", 32767);
            expected.Add("A4", -32767);
            expected.Add("A5", short.MinValue);
            expected.Add("A6", short.MaxValue);
            AssertValues<short, int>(cells, AssertEquals, expected);
        }

        [Fact(DisplayName = "Test of the reader functionality for ushort values (cast to int)")]
        public void ReadUshortTest()
        {
            Dictionary<string, ushort> cells = new Dictionary<string, ushort>();
            cells.Add("A1", 0);
            cells.Add("A2", 10);
            cells.Add("A3", 56353);
            cells.Add("A4", ushort.MinValue);
            cells.Add("A5", ushort.MaxValue);

            Dictionary<string, int> expected = new Dictionary<string, int>();
            expected.Add("A1", 0);
            expected.Add("A2", 10);
            expected.Add("A3", 56353);
            expected.Add("A4", ushort.MinValue);
            expected.Add("A5", ushort.MaxValue);
            AssertValues<ushort, int>(cells, AssertEquals, expected);
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
            AssertValues<float>(cells, AssertApproximateFloat);
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
            AssertValues<double>(cells, AssertApproximateDouble);
        }

        [Fact(DisplayName = "Test of the reader functionality for bool values")]
        public void ReadBoolTest()
        {
            Dictionary<string, bool> cells = new Dictionary<string, bool>();
            cells.Add("A1", true);
            cells.Add("A2", false);
            cells.Add("A3", true);
            AssertValues<bool>(cells, AssertEquals);
        }

        [Fact(DisplayName = "Test of the reader functionality for DateTime values")]
        public void ReadDateTimeTest()
        {
            Dictionary<string, DateTime> cells = new Dictionary<string, DateTime>();
            cells.Add("A1", new DateTime(2021, 5, 11, 15, 7, 2));
            cells.Add("A2", new DateTime(1900, 1, 1, 0, 0, 0));
            cells.Add("A3", new DateTime(1960, 12, 12));
            cells.Add("A4", new DateTime(9999, 12, 31, 23, 59, 59));
            AssertValues<DateTime>(cells, AssertEquals);
        }

        [Fact(DisplayName = "Test of the reader functionality for TimeSpan values")]
        public void ReadTimeSpanTest()
        {
            Dictionary<string, TimeSpan> cells = new Dictionary<string, TimeSpan>();
            cells.Add("A1", new TimeSpan(0, 0, 0));
            cells.Add("A2", new TimeSpan(13, 18, 22));
            cells.Add("A3", new TimeSpan(12, 0, 0));
            cells.Add("A4", new TimeSpan(23, 59, 59));
            AssertValues<TimeSpan>(cells, AssertEquals);
        }

        [Fact(DisplayName = "Test of the reader functionality for formulas (no formula parsing)")]
        public void ReadFormulaTest()
        {
            Dictionary<string, string> cells = new Dictionary<string, string>();
            cells.Add("A1", "=B2");
            cells.Add("A2", "MIN(C2:D2)");
            cells.Add("A3", "MAX(worksheet2!A1:worksheet2:A100");

            Workbook workbook = new Workbook("worksheet1");
            foreach (KeyValuePair<string, string> cell in cells)
            {
                workbook.CurrentWorksheet.AddCellFormula(cell.Value, cell.Key);
            }
            MemoryStream stream = new MemoryStream();
            workbook.SaveAsStream(stream, true);
            stream.Position = 0;
            Workbook givenWorkbook = WorkbookReader.Load(stream);

            Assert.NotNull(givenWorkbook);
            Worksheet givenWorksheet = givenWorkbook.SetCurrentWorksheet(0);
            Assert.Equal("worksheet1", givenWorksheet.SheetName);
            foreach (string address in cells.Keys)
            {
                Cell givenCell = givenWorksheet.GetCell(new Address(address));
                Assert.Equal(Cell.CellType.FORMULA, givenCell.DataType);
                Assert.Equal(cells[address], givenCell.Value);
            }
        }

        [Theory(DisplayName = "Test of the reader functionality on invalid / unexpected values")]
        [InlineData("A1", Cell.CellType.STRING, "Test")]
        [InlineData("B1", Cell.CellType.STRING, "x")]
        [InlineData("C1", Cell.CellType.NUMBER, -1.8538541667)]
        [InlineData("D1", Cell.CellType.NUMBER, 2)]
        [InlineData("E1", Cell.CellType.STRING, "x")]
        [InlineData("F1", Cell.CellType.STRING, "1")] // Reference 1 is casted to string '1'
        [InlineData("G1", Cell.CellType.NUMBER, -1.5f)]
        [InlineData("H1", Cell.CellType.STRING, "y")]
        [InlineData("I1", Cell.CellType.BOOL, true)]
        [InlineData("J1", Cell.CellType.BOOL, false)]
        [InlineData("K1", Cell.CellType.STRING, "z")]
        [InlineData("L1", Cell.CellType.STRING, "z")]
        [InlineData("M1", Cell.CellType.STRING, "a")]
        public void ReadInvalidDataTest(string cellAddress, Cell.CellType expectedType, object expectedValue)
        {
            // Note: Cell A1 is a valid string
            //       Cell B1 is declared numerical, but contains a string
            //       Cell C1 is defined as date but has a negative number
            //       Cell D1 is defined ad bool but has an invalid value of 2
            //       Cell E1 is defined as bool but has an invalid value of 'x'
            //       Cell F1 is defined as shared string value, but the value does not exist
            //       Cell G1 is defined as time but has a negative number
            //       Cell H1 is defined as the unknown type 'z'
            //       Cell I1 is defined as boolean but has 'true' instead of 1 as XML value
            //       Cell J1 is defined as boolean but has 'FALSE' instead of 0 as XML value
            //       Cell K1 is defined as date but has an invalid value of 'z'
            //       Cell L1 is defined as time but has an invalid value of 'z'
            //       Cell M1 is defined as shared string but has an invalid value of 'a'
            Stream stream = TestUtils.GetResource("tampered.xlsx");
            Workbook workbook = WorkbookReader.Load(stream);
            Assert.Equal(expectedType, workbook.Worksheets[0].Cells[cellAddress].DataType);
            Assert.Equal(expectedValue, workbook.Worksheets[0].Cells[cellAddress].Value);
        }

        [Theory(DisplayName = "Test of the failing reader functionality on invalid XML content")]
        [InlineData("invalid_workbook.xlsx")]
        [InlineData("invalid_workbook_sheet-definition.xlsx")]
        [InlineData("invalid_worksheet.xlsx")]
        [InlineData("invalid_style.xlsx")]
        [InlineData("invalid_metadata_app.xlsx")]
        [InlineData("invalid_metadata_core.xlsx")]
        [InlineData("invalid_sharedStrings.xlsx")]
        [InlineData("invalid_sharedStrings2.xlsx")]
        [InlineData("invalid_relationship.xlsx")]
        [InlineData("missing_worksheet.xlsx")]
        public void FailingReadInvalidDataTest(string invalidFile)
        {
            // Note: all referenced (embedded) files contains invalid XML documents (malformed, missing start or end tags, missing attributes)
            Stream stream = TestUtils.GetResource(invalidFile);
            Assert.Throws<NanoXLSX.Exceptions.IOException>(() => WorkbookReader.Load(stream));
        }

        [Theory(DisplayName = "Test of the failing asynchronous reader functionality on invalid XML content")]
        [InlineData("invalid_workbook.xlsx")]
        [InlineData("invalid_workbook_sheet-definition.xlsx")]
        [InlineData("invalid_worksheet.xlsx")]
        [InlineData("invalid_style.xlsx")]
        [InlineData("invalid_metadata_app.xlsx")]
        [InlineData("invalid_metadata_core.xlsx")]
        [InlineData("invalid_sharedStrings.xlsx")]
        [InlineData("invalid_sharedStrings2.xlsx")]
        [InlineData("invalid_relationship.xlsx")]
        [InlineData("empty_worksheet.xlsx")]
        [InlineData("missing_worksheet.xlsx")]

        public async Task FailingAsyncReadInvalidDataTest(string invalidFile)
        {
            // Note: all referenced (embedded) files contains invalid XML documents (malformed, missing start or end tags, missing attributes)
            Stream stream = TestUtils.GetResource(invalidFile);
            await Assert.ThrowsAsync<NanoXLSX.Exceptions.IOException>(() => WorkbookReader.LoadAsync(stream));
        }

        [Fact(DisplayName = "Test of the workbook reader if the only workbook entry is a chart")]
        public void ReadChartsheetTest()
        {
            Stream stream = TestUtils.GetResource("chartsheet.xlsx");
            Workbook workbook = WorkbookReader.Load(stream);
            Assert.Single(workbook.Worksheets);
            Assert.Empty(workbook.Worksheets[0].Cells);
        }

        [Fact(DisplayName = "Test of the workbook reader if the workbook contains worksheets chats and embedded charts")]
        public void ReadChartsheetTest2()
        {
            // Note: Sheet1 and Sheet3 contains data. Diagram1 (worksheet2) is just a chart and should be empty
            Stream stream = TestUtils.GetResource("chartsheet2.xlsx");
            Workbook workbook = WorkbookReader.Load(stream);
            Assert.Equal(3, workbook.Worksheets.Count);
            Assert.True(workbook.GetWorksheet("Sheet1").Cells.Count > 0);
            Assert.Empty(workbook.GetWorksheet("Diagram1").Cells);
            Assert.True(workbook.GetWorksheet("Sheet3").Cells.Count > 0);
        }

        [Fact(DisplayName = "Test of the reader functionality on an invalid stream")]
        public void ReadInvalidStreamTest()
        {
            Stream nullStream = null;
            Assert.Throws<NanoXLSX.Exceptions.IOException>(() => WorkbookReader.Load(nullStream));
        }

        [Fact(DisplayName = "Test of the AutoFilter behavior when a worksheet is read on an address instead of a range")]
        public void AutoFilterReadTest()
        {
            // The file contains an auto filter on Column "A1", described as an address instead of a range
            Stream stream = TestUtils.GetResource("autofilter.xlsx");
            Workbook workbook = WorkbookReader.Load(stream);
            Assert.Equal(new NanoXLSX.Range("A1:A1").ToString(), workbook.Worksheets[0].AutoFilterRange.Value.ToString());
        }

        private static void AssertEquals<T>(T expected, T given)
        {
            Assert.Equal(expected, given);
        }

        private static void AssertValues<T>(Dictionary<string, T> givenCells, Action<T, T> assertionAction, Dictionary<string, T> expectedCells = null)
        {
            Worksheet givenWorksheet = GetWorksheet(givenCells);
            foreach (string address in givenCells.Keys)
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
                    assertionAction.Invoke(value, (T)givenCell.Value);
                }
            }
        }

        private static void AssertValues<T, D>(Dictionary<string, T> givenCells, Action<D, D> assertionAction, Dictionary<string, D> expectedCells)
        {
            Worksheet givenWorksheet = GetWorksheet(givenCells);
            foreach (string address in givenCells.Keys)
            {
                Cell givenCell = givenWorksheet.GetCell(new Address(address));
                D givenValue = (D)givenCell.Value;
                D expectedvalue = expectedCells[address];
                if (givenValue == null)
                {
                    Assert.Equal(Cell.CellType.EMPTY, givenCell.DataType);
                }
                else
                {
                    assertionAction.Invoke(givenValue, expectedvalue);
                }
            }
        }

        private static Worksheet GetWorksheet<T>(Dictionary<string, T> givenCells)
        {
            Workbook workbook = new Workbook("worksheet1");
            foreach (KeyValuePair<string, T> cell in givenCells)
            {
                workbook.CurrentWorksheet.AddCell(cell.Value, cell.Key);
            }
            MemoryStream stream = new MemoryStream();
            workbook.SaveAsStream(stream, true);
            stream.Position = 0;
            Workbook givenWorkbook = WorkbookReader.Load(stream);

            Assert.NotNull(givenWorkbook);
            Worksheet givenWorksheet = givenWorkbook.SetCurrentWorksheet(0);
            Assert.Equal("worksheet1", givenWorksheet.SheetName);
            return givenWorksheet;
        }



        private static void AssertApproximateDouble(double expected, double given)
        {
            double threshold = 0.00000001;
            Assert.True(Math.Abs(given - expected) < threshold);
        }

        private static void AssertApproximateFloat(float expected, float given)
        {
            float threshold = 0.00000001f;
            Assert.True(Math.Abs(given - expected) < threshold);
        }


    }
}
