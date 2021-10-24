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
    public class ImportOptionTest
    {

        [Fact( DisplayName = "Test of the reader functionality with the global import option to cast everything to string")]
        public void CastAllToStringTest()
        {
            Dictionary<string, Object> cells = new Dictionary<string, object>();
            cells.Add("A1", "test");
            cells.Add("A2", true);
            cells.Add("A3", false);
            cells.Add("A4", 42);
            cells.Add("A5", 0.55f);
            cells.Add("A6", -0.111d);
            cells.Add("A7", new DateTime(2020, 11, 10, 9, 8, 7, 0));
            cells.Add("A8", new TimeSpan(18, 15, 12));
            cells.Add("A9", null);
            Dictionary<string, string> expectedCells = new Dictionary<string, string>();
            expectedCells.Add("A1", "test");
            expectedCells.Add("A2", "True");
            expectedCells.Add("A3", "False");
            expectedCells.Add("A4", "42");
            expectedCells.Add("A5", "0.55");
            expectedCells.Add("A6", "-0.111");
            expectedCells.Add("A7", "2020-11-10 09:08:07");
            expectedCells.Add("A8", "18:15:12");
            expectedCells.Add("A9", null);

            ImportOptions options = new ImportOptions();
            options.GlobalEnforcingType = ImportOptions.GlobalType.EverythingToString;
            AssertValues<object, string>(cells, options, AssertEquals, expectedCells);
        }

        [Fact(DisplayName = "Test of the reader functionality with the global import option to cast all number to double")]
        public void CastToDoubleTest()
        {
            Dictionary<string, Object> cells = new Dictionary<string, object>();
            cells.Add("A1", "test");
            cells.Add("A2", true);
            cells.Add("A3", false);
            cells.Add("A4", 42);
            cells.Add("A5", 0.55f);
            cells.Add("A6", -0.111d);
            cells.Add("A7", new DateTime(2020, 11, 10, 9, 8, 7, 0));
            cells.Add("A8", new TimeSpan(18, 15, 12));
            cells.Add("A9", null);
            cells.Add("A10", "27");
            Dictionary<string, object> expectedCells = new Dictionary<string, object>();
            expectedCells.Add("A1", "test");
            expectedCells.Add("A2", 1d);
            expectedCells.Add("A3", 0d);
            expectedCells.Add("A4", 42d);
            expectedCells.Add("A5", 0.55d);
            expectedCells.Add("A6", -0.111d);
            expectedCells.Add("A7", double.Parse(Utils.GetOADateTimeString(new DateTime(2020,11,10,9,8,7,0))));
            expectedCells.Add("A8", double.Parse(Utils.GetOATimeString(new TimeSpan(18,15,12))));
            expectedCells.Add("A9", null);
            expectedCells.Add("A10", 27d);
            ImportOptions options = new ImportOptions();
            options.GlobalEnforcingType = ImportOptions.GlobalType.AllNumbersToDouble;
            AssertValues<object, object>(cells, options, AssertApproximate, expectedCells);
        }

        [Fact(DisplayName = "Test of the reader functionality with the global import option to cast all number to int")]
        public void CastToIntTest()
        {
            Dictionary<string, Object> cells = new Dictionary<string, object>();
            cells.Add("A1", "test");
            cells.Add("A2", true);
            cells.Add("A3", false);
            cells.Add("A4", 42);
            cells.Add("A5", 0.55f);
            cells.Add("A6", -3.111d);
            cells.Add("A7", new DateTime(2020, 11, 10, 9, 8, 7, 0));
            cells.Add("A8", new TimeSpan(18,15,12));
            cells.Add("A9", -4.9f);
            cells.Add("A10", 0.49d);
            cells.Add("A11", null);
            cells.Add("A12", "28");
            Dictionary<string, object> expectedCells = new Dictionary<string, object>();
            expectedCells.Add("A1", "test");
            expectedCells.Add("A2", 1);
            expectedCells.Add("A3", 0);
            expectedCells.Add("A4", 42);
            expectedCells.Add("A5", 1);
            expectedCells.Add("A6", -3);
            expectedCells.Add("A7", (int)Math.Round(double.Parse(Utils.GetOADateTimeString(new DateTime(2020, 11, 10, 9, 8, 7, 0))),0));
            expectedCells.Add("A8", (int)Math.Round(double.Parse(Utils.GetOATimeString(new TimeSpan(18,15,12))), 0));
            expectedCells.Add("A9", -5);
            expectedCells.Add("A10", 0);
            expectedCells.Add("A11", null);
            expectedCells.Add("A12", 28);
            ImportOptions options = new ImportOptions();
            options.GlobalEnforcingType = ImportOptions.GlobalType.AllNumbersToInt;
            AssertValues<object, object>(cells, options, AssertApproximate, expectedCells);
        }

        [Fact(DisplayName = "Test of the reader functionality with the import option EnforceEmptyValuesAsString")]
        public void EnforceEmptyValuesAsStringTest()
        {
            Dictionary<string, Object> cells = new Dictionary<string, object>();
            cells.Add("A1", "test");
            cells.Add("A2", true);
            cells.Add("A3", 22.2d);
            cells.Add("A4", null);
            cells.Add("A5", "");
            Dictionary<string, object> expectedCells = new Dictionary<string, object>();
            expectedCells.Add("A1", "test");
            expectedCells.Add("A2", true);
            expectedCells.Add("A3", 22.2f); // Import will go to the smallest float unit (float 32 / single)
            expectedCells.Add("A4", "");
            expectedCells.Add("A5", "");
            ImportOptions options = new ImportOptions();
            options.EnforceEmptyValuesAsString = true;
            AssertValues<object, object>(cells, options, AssertApproximate, expectedCells);
        }

        [Fact(DisplayName = "Test of the EnforcingStartRowNumber functionality on global enforcing rules")]
        public void EnforcingStartRowNumberTest()
        {
            Dictionary<string, Object> cells = new Dictionary<string, object>();
            cells.Add("A1", 22);
            cells.Add("A2", true);
            cells.Add("A3", 22);
            cells.Add("A4", true);
            cells.Add("A5", 22.5d);
            Dictionary<string, object> expectedCells = new Dictionary<string, object>();
            expectedCells.Add("A1", 22);
            expectedCells.Add("A2", true);
            expectedCells.Add("A3", "22");
            expectedCells.Add("A4", "True");
            expectedCells.Add("A5", "22.5");
            ImportOptions options = new ImportOptions();
            options.EnforcingStartRowNumber = 2;
            options.GlobalEnforcingType = ImportOptions.GlobalType.EverythingToString;
            AssertValues<object, object>(cells, options, AssertApproximate, expectedCells);
        }

        [Theory(DisplayName = "Test of the import options for the import column type: Double")]
        [InlineData("B")]
        [InlineData(1)]
        public void EnforcingColumnAsNumberTest(object column)
        {
            TimeSpan time = new TimeSpan(11, 12, 13);
            DateTime date = new DateTime(2021, 8, 14, 18, 22, 13, 0);
            Dictionary<string, Object> cells = new Dictionary<string, object>();
            cells.Add("A1", 22);
            cells.Add("A2", "21");
            cells.Add("A3", true);
            cells.Add("B1", 23);
            cells.Add("B2", "20");
            cells.Add("B3", true);
            cells.Add("B4", time);
            cells.Add("B5", date);
            cells.Add("C1", "2");
            cells.Add("C2", new TimeSpan(12, 14, 16));
            Dictionary<string, object> expectedCells = new Dictionary<string, object>();
            expectedCells.Add("A1", 22);
            expectedCells.Add("A2", "21");
            expectedCells.Add("A3", true);
            expectedCells.Add("B1", 23d);
            expectedCells.Add("B2", 20d);
            expectedCells.Add("B3", 1d);
            expectedCells.Add("B4",  double.Parse(Utils.GetOATimeString(time)));
            expectedCells.Add("B5", double.Parse(Utils.GetOADateTimeString(date)));
            expectedCells.Add("C1", "2");
            expectedCells.Add("C2", new TimeSpan(12, 14, 16));
            ImportOptions options = new ImportOptions();
            if (column is string)
            {
                options.AddEnforcedColumn(column as string, ImportOptions.ColumnType.Double);
            }
            else
            {
                options.AddEnforcedColumn((int)column, ImportOptions.ColumnType.Double);
            }
            AssertValues<object, object>(cells, options, AssertApproximate, expectedCells);
        }

        [Theory(DisplayName = "Test of the import options for the import column type: Numeric")]
        [InlineData("B")]
        [InlineData(1)]
        public void EnforcingColumnAsNumberTest2(object column)
        {
            TimeSpan time = new TimeSpan(11, 12, 13);
            DateTime date = new DateTime(2021, 8, 14, 18, 22, 13, 0);
            Dictionary<string, Object> cells = new Dictionary<string, object>();
            cells.Add("A1", 22);
            cells.Add("A2", "21");
            cells.Add("A3", true);
            cells.Add("B1", 23);
            cells.Add("B2", "20.1");
            cells.Add("B3", true);
            cells.Add("B4", time);
            cells.Add("B5", date);
            cells.Add("C1", "2");
            cells.Add("C2", new TimeSpan(12, 14, 16));
            Dictionary<string, object> expectedCells = new Dictionary<string, object>();
            expectedCells.Add("A1", 22);
            expectedCells.Add("A2", "21");
            expectedCells.Add("A3", true);
            expectedCells.Add("B1", 23);
            expectedCells.Add("B2", 20.1f);
            expectedCells.Add("B3", 1);
            expectedCells.Add("B4", float.Parse(Utils.GetOATimeString(time)));
            expectedCells.Add("B5", float.Parse(Utils.GetOADateTimeString(date)));
            expectedCells.Add("C1", "2");
            expectedCells.Add("C2", new TimeSpan(12, 14, 16));
            ImportOptions options = new ImportOptions();
            if (column is string)
            {
                options.AddEnforcedColumn(column as string, ImportOptions.ColumnType.Numeric);
            }
            else
            {
                options.AddEnforcedColumn((int)column, ImportOptions.ColumnType.Numeric);
            }
            AssertValues<object, object>(cells, options, AssertApproximate, expectedCells);
        }

        [Theory(DisplayName = "Test of the import options for the import column type: Bool")]
        [InlineData("B")]
        [InlineData(1)]
        public void EnforcingColumnAsBoolTest(object column)
        {
            TimeSpan time = new TimeSpan(11, 12, 13);
            DateTime date = new DateTime(2021, 8, 14, 18, 22, 13, 0);
            Dictionary<string, Object> cells = new Dictionary<string, object>();
            cells.Add("A1", 1);
            cells.Add("A2", "21");
            cells.Add("A3", true);
            cells.Add("B1", 1);
            cells.Add("B2", "true");
            cells.Add("B3", false);
            cells.Add("B4", time);
            cells.Add("B5", date);
            cells.Add("B6", 0f);
            cells.Add("B7", "1");
            cells.Add("B8", "Test");
            cells.Add("B9", 1.0d);
            cells.Add("C1", "0");
            cells.Add("C2", new TimeSpan(12, 14, 16));
            Dictionary<string, object> expectedCells = new Dictionary<string, object>();
            expectedCells.Add("A1", 1);
            expectedCells.Add("A2", "21");
            expectedCells.Add("A3", true);
            expectedCells.Add("B1", true);
            expectedCells.Add("B2", true);
            expectedCells.Add("B3", false);
            expectedCells.Add("B4", time);
            expectedCells.Add("B5", date);
            expectedCells.Add("B6", false);
            expectedCells.Add("B7", true);
            expectedCells.Add("B8", "Test");
            expectedCells.Add("B9", true);
            expectedCells.Add("C1", "0");
            expectedCells.Add("C2", new TimeSpan(12, 14, 16));
            ImportOptions options = new ImportOptions();
            if (column is string)
            {
                options.AddEnforcedColumn(column as string, ImportOptions.ColumnType.Bool);
            }
            else
            {
                options.AddEnforcedColumn((int)column, ImportOptions.ColumnType.Bool);
            }
            AssertValues<object, object>(cells, options, AssertApproximate, expectedCells);
        }

        [Theory(DisplayName = "Test of the import options for the import column type: String")]
        [InlineData("B")]
        [InlineData(1)]
        public void EnforcingColumnAsStringTest(object column)
        {
            TimeSpan time = new TimeSpan(11, 12, 13);
            DateTime date = new DateTime(2021, 8, 14, 18, 22, 13, 0);
            Dictionary<string, Object> cells = new Dictionary<string, object>();
            cells.Add("A1", 1);
            cells.Add("A2", "21");
            cells.Add("A3", true);
            cells.Add("B1", 1);
            cells.Add("B2", "Test");
            cells.Add("B3", false);
            cells.Add("B4", time);
            cells.Add("B5", date);
            cells.Add("B6", 0f);
            cells.Add("B7", true);
            cells.Add("B8", -10);
            cells.Add("B9", 1.111d);
            cells.Add("C1", "0");
            cells.Add("C2", new TimeSpan(12, 14, 16));
            Dictionary<string, object> expectedCells = new Dictionary<string, object>();
            expectedCells.Add("A1", 1);
            expectedCells.Add("A2", "21");
            expectedCells.Add("A3", true);
            expectedCells.Add("B1", "1");
            expectedCells.Add("B2", "Test");
            expectedCells.Add("B3", "False");
            expectedCells.Add("B4", time.ToString(ImportOptions.DEFAULT_TIMESPAN_FORMAT));
            expectedCells.Add("B5", date.ToString(ImportOptions.DEFAULT_DATETIME_FORMAT));
            expectedCells.Add("B6", "0");
            expectedCells.Add("B7", "True");
            expectedCells.Add("B8", "-10");
            expectedCells.Add("B9", "1.111");
            expectedCells.Add("C1", "0");
            expectedCells.Add("C2", new TimeSpan(12, 14, 16));
            ImportOptions options = new ImportOptions();
            if (column is string)
            {
                options.AddEnforcedColumn(column as string, ImportOptions.ColumnType.String);
            }
            else
            {
                options.AddEnforcedColumn((int)column, ImportOptions.ColumnType.String);
            }
            AssertValues<object, object>(cells, options, AssertApproximate, expectedCells);
        }

        private static void AssertValues<T,D>(Dictionary<string, T> givenCells, ImportOptions importOptions, Action<object, object> assertionAction, Dictionary<string, D> expectedCells = null)
        {
            Workbook workbook = new Workbook("worksheet1");
            foreach (KeyValuePair<string, T> cell in givenCells)
            {
                workbook.CurrentWorksheet.AddCell(cell.Value, cell.Key);
            }
            MemoryStream stream = new MemoryStream();
            workbook.SaveAsStream(stream, true);
            stream.Position = 0;
            Workbook givenWorkbook = Workbook.Load(stream, importOptions);

            Assert.NotNull(givenWorkbook);
            Worksheet givenWorksheet = givenWorkbook.SetCurrentWorksheet(0);
            Assert.Equal("worksheet1", givenWorksheet.SheetName);
            foreach (string address in givenCells.Keys)
            {
                Cell givenCell = givenWorksheet.GetCell(new Address(address));
                D expectedValue = expectedCells[address];
                if (expectedValue == null)
                {
                    Assert.Equal(Cell.CellType.EMPTY, givenCell.DataType);
                }
                else
                {
                    assertionAction.Invoke(expectedValue, (D)givenCell.Value);
                }
            }
        }
        private static void AssertEquals<T>(T expected, T given)
        {
            Assert.Equal(expected, given);
        }

        private static void AssertApproximate(object expected, object given)
        {
            double threshold = 0.000012; // The precision may vary (roughly one second)
            if (given is double)
            {
                Assert.True(Math.Abs((double)given - (double)expected) < threshold);
            }
            else if (given is float)
            {
                Assert.True(Math.Abs((float)given - (float)expected) < threshold);
            }
            else if (given is DateTime)
            {
                double e = double.Parse(Utils.GetOADateTimeString((DateTime)expected));
                double g = double.Parse(Utils.GetOADateTimeString((DateTime)given));
                AssertApproximate(e, g);
            }
            else if (given is TimeSpan)
            {
                double g = double.Parse(Utils.GetOATimeString((TimeSpan)given));
                double e = double.Parse(Utils.GetOATimeString((TimeSpan)expected));
                AssertApproximate(e, g);
            }
            else
            {
                AssertEquals<object>(expected,given);
            }
            
        }

    }
}
