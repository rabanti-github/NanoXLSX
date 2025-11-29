using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using NanoXLSX.Extensions;
using NanoXLSX.Styles;
using NanoXLSX.Test.Writer_Reader.Utils;
using NanoXLSX.Utils;
using Xunit;

namespace NanoXLSX.Test.Writer_Reader.ReaderTest
{
    public class ReaderOptionTest
    {


        [Fact(DisplayName = "Test of the reader functionality with the global import option to cast everything to string")]
        public void CastAllToStringTest()
        {
            Dictionary<string, Object> cells = new Dictionary<string, object>
            {
                { "A1", "test" },
                { "A2", true },
                { "A3", false },
                { "A4", 42 },
                { "A5", 0.55f },
                { "A6", -0.111d },
                { "A7", new DateTime(2020, 11, 10, 9, 8, 7, 0) },
                { "A8", new TimeSpan(18, 15, 12) },
                { "A9", null },
                { "A10", new Cell("=A1", Cell.CellType.Formula, "A10") },
                { "A11", 8294967296.25d } // high-range double
            };
            Dictionary<string, string> expectedCells = new Dictionary<string, string>
            {
                { "A1", "test" },
                { "A2", "True" },
                { "A3", "False" },
                { "A4", "42" },
                { "A5", "0.55" },
                { "A6", "-0.111" },
                { "A7", "2020-11-10 09:08:07" },
                { "A8", "18:15:12" },
                { "A9", null }, // Empty remains empty
                { "A10", "=A1" },
                { "A11", "8294967296.25" }
            };

            ReaderOptions options = new ReaderOptions
            {
                GlobalEnforcingType = ReaderOptions.GlobalType.EverythingToString
            };
            AssertValues<object, string>(cells, options, AssertEquals, expectedCells);
        }


        [Fact(DisplayName = "Test of the reader functionality with the global import option to cast all number to decimal")]
        public void CastToDecimalTest()
        {
            Dictionary<string, Object> cells = new Dictionary<string, object>
            {
                { "A1", "test" },
                { "A2", true },
                { "A3", false },
                { "A4", 42 },
                { "A5", 0.55f },
                { "A6", -0.111d },
                { "A7", new DateTime(2020, 11, 10, 9, 8, 7, 0) },
                { "A8", new TimeSpan(18, 15, 12) },
                { "A9", null },
                { "A10", "27" },
                { "A11", new Cell("=A1", Cell.CellType.Formula, "A11") }
            };
            Dictionary<string, object> expectedCells = new Dictionary<string, object>
            {
                { "A1", "test" },
                { "A2", decimal.One },
                { "A3", decimal.Zero },
                { "A4", 42m },
                { "A5", 0.55m },
                { "A6", -0.111m },
                { "A7", (decimal)DataUtils.GetOADateTime(new DateTime(2020, 11, 10, 9, 8, 7, 0)) },
                { "A8", (decimal)DataUtils.GetOATime(new TimeSpan(18, 15, 12)) },
                { "A9", null },
                { "A10", 27m },
                { "A11", new Cell("=A1", Cell.CellType.Formula, "A11") }
            };
            ReaderOptions options = new ReaderOptions
            {
                GlobalEnforcingType = ReaderOptions.GlobalType.AllNumbersToDecimal
            };
            AssertValues<object, object>(cells, options, AssertApproximate, expectedCells);
        }


        [Fact(DisplayName = "Test of the reader functionality with the global import option to cast all number to double")]
        public void CastToDoubleTest()
        {
            Dictionary<string, Object> cells = new Dictionary<string, object>
            {
                { "A1", "test" },
                { "A2", true },
                { "A3", false },
                { "A4", 42 },
                { "A5", 0.55f },
                { "A6", -0.111d },
                { "A7", new DateTime(2020, 11, 10, 9, 8, 7, 0) },
                { "A8", new TimeSpan(18, 15, 12) },
                { "A9", null },
                { "A10", "27" },
                { "A11", new Cell("=A1", Cell.CellType.Formula, "A11") }
            };
            Dictionary<string, object> expectedCells = new Dictionary<string, object>
            {
                { "A1", "test" },
                { "A2", 1d },
                { "A3", 0d },
                { "A4", 42d },
                { "A5", 0.55d },
                { "A6", -0.111d },
                { "A7", DataUtils.GetOADateTime(new DateTime(2020, 11, 10, 9, 8, 7, 0)) },
                { "A8", DataUtils.GetOATime(new TimeSpan(18, 15, 12)) },
                { "A9", null },
                { "A10", 27d },
                { "A11", new Cell("=A1", Cell.CellType.Formula, "A11") }
            };
            ReaderOptions options = new ReaderOptions
            {
                GlobalEnforcingType = ReaderOptions.GlobalType.AllNumbersToDouble
            };
            AssertValues<object, object>(cells, options, AssertApproximate, expectedCells);
        }

        [Fact(DisplayName = "Test of the reader functionality with the global import option to cast all number to int")]
        public void CastToIntTest()
        {
            Dictionary<string, Object> cells = new Dictionary<string, object>
            {
                { "A1", "test" },
                { "A2", true },
                { "A3", false },
                { "A4", 42 },
                { "A5", 0.55f },
                { "A6", -3.111d },
                { "A7", new DateTime(2020, 11, 10, 9, 8, 7, 0) },
                { "A8", new TimeSpan(18, 15, 12) },
                { "A9", -4.9f },
                { "A10", 0.49d },
                { "A11", null },
                { "A12", "28" },
                { "A13", new Cell("=A1", Cell.CellType.Formula, "A13") },
                { "A14", 8589934592L },
                { "A15", 2147483650.6d },
                { "A16", 4294967294u },
                { "A17", 18446744073709551614 }
            };
            Dictionary<string, object> expectedCells = new Dictionary<string, object>
            {
                { "A1", "test" },
                { "A2", 1 },
                { "A3", 0 },
                { "A4", 42 },
                { "A5", 1 },
                { "A6", -3 },
                { "A7", (int)Math.Round(DataUtils.GetOADateTime(new DateTime(2020, 11, 10, 9, 8, 7, 0)), 0) },
                { "A8", (int)Math.Round(DataUtils.GetOATime(new TimeSpan(18, 15, 12)), 0) },
                { "A9", -5 },
                { "A10", 0 },
                { "A11", null },
                { "A12", 28 },
                { "A13", new Cell("=A1", Cell.CellType.Formula, "A13") },
                { "A14", 8589934592L },
                { "A15", 2147483650.6 },
                { "A16", 4294967294u },
                { "A17", 18446744073709551614 }
            };
            ReaderOptions options = new ReaderOptions
            {
                GlobalEnforcingType = ReaderOptions.GlobalType.AllNumbersToInt
            };
            AssertValues<object, object>(cells, options, AssertApproximate, expectedCells);
        }

        [Fact(DisplayName = "Test of the reader functionality with the import option EnforceEmptyValuesAsString")]
        public void EnforceEmptyValuesAsStringTest()
        {
            Dictionary<string, Object> cells = new Dictionary<string, object>
            {
                { "A1", "test" },
                { "A2", true },
                { "A3", 22.2d },
                { "A4", null },
                { "A5", "" },
                { "A6", new Cell("=A1", Cell.CellType.Formula, "A6") }
            };
            Dictionary<string, object> expectedCells = new Dictionary<string, object>
            {
                { "A1", "test" },
                { "A2", true },
                { "A3", 22.2f }, // Import will go to the smallest float unit (float 32 / single)
                { "A4", "" },
                { "A5", "" },
                { "A6", new Cell("=A1", Cell.CellType.Formula, "A6") }
            };
            ReaderOptions options = new ReaderOptions
            {
                EnforceEmptyValuesAsString = true
            };
            AssertValues<object, object>(cells, options, AssertApproximate, expectedCells);
        }

        [Fact(DisplayName = "Test of the EnforcingStartRowNumber functionality on global enforcing rules")]
        public void EnforcingStartRowNumberTest()
        {
            Dictionary<string, Object> cells = new Dictionary<string, object>
            {
                { "A1", 22 },
                { "A2", true },
                { "A3", new Cell("=A1", Cell.CellType.Formula, "A3") },
                { "A4", 22 },
                { "A5", true },
                { "A6", 22.5d },
                { "A7", new Cell("=A1", Cell.CellType.Formula, "A7") }
            };
            Dictionary<string, object> expectedCells = new Dictionary<string, object>
            {
                { "A1", 22 },
                { "A2", true },
                { "A3", new Cell("=A1", Cell.CellType.Formula, "A3") },
                { "A4", "22" },
                { "A5", "True" },
                { "A6", "22.5" },
                { "A7", "=A1" }
            };
            ReaderOptions options = new ReaderOptions
            {
                EnforcingStartRowNumber = 3,
                GlobalEnforcingType = ReaderOptions.GlobalType.EverythingToString
            };
            AssertValues<object, object>(cells, options, AssertApproximate, expectedCells);
        }

        [Fact(DisplayName = "Test of the EnforceDateTimesAsNumbers functionality on global enforcing rules")]
        public void EnforceDateTimesAsNumbersTest()
        {
            DateTime date = new DateTime(2021, 8, 17, 11, 12, 13, 0);
            TimeSpan time = new TimeSpan(18, 14, 10);
            Dictionary<string, Object> cells = new Dictionary<string, object>
            {
                { "A1", 22 },
                { "A2", true },
                { "A3", date },
                { "A4", time },
                { "A5", 22.5f },
                { "A6", new Cell("=A1", Cell.CellType.Formula, "A6") }
            };
            Dictionary<string, object> expectedCells = new Dictionary<string, object>
            {
                { "A1", 22 },
                { "A2", true },
                { "A3", DataUtils.GetOADateTime(date) },
                { "A4", DataUtils.GetOATime(time) },
                { "A5", 22.5f },
                { "A6", new Cell("=A1", Cell.CellType.Formula, "A6") }
            };
            ReaderOptions options = new ReaderOptions
            {
                EnforceDateTimesAsNumbers = true
            };
            AssertValues<object, object>(cells, options, AssertApproximate, expectedCells);
        }

        [Theory(DisplayName = "Test of the EnforceDateTimesAsNumbers functionality for the import column types: Date or Time")]
        [InlineData(ReaderOptions.ColumnType.Date, 22.5f, 22.5d)]
        [InlineData(ReaderOptions.ColumnType.Time, 22.5d, 22.5d)]
        public void EnforceDateTimesAsNumbersTest2(ReaderOptions.ColumnType columnType, object givenLowNumber, object expectedLowNumber)
        {
            DateTime date = new DateTime(2021, 8, 17, 11, 12, 13, 0);
            TimeSpan time = new TimeSpan(18, 14, 10);
            Dictionary<string, Object> cells = new Dictionary<string, object>
            {
                { "A1", 22 },
                { "A2", true },
                { "A3", date },
                { "A4", time },
                { "B1", date },
                { "B2", time },
                { "B3", givenLowNumber },
                { "B4", new Cell("=A1", Cell.CellType.Formula, "B4") }
            };
            Dictionary<string, object> expectedCells = new Dictionary<string, object>
            {
                { "A1", 22 },
                { "A2", true },
                { "A3", DataUtils.GetOADateTime(date) },
                { "A4", DataUtils.GetOATime(time) },
                { "B1", DataUtils.GetOADateTime(date) },
                { "B2", DataUtils.GetOATime(time) },
                { "B3", expectedLowNumber },
                { "B4", new Cell("=A1", Cell.CellType.Formula, "B4") }
            };
            ReaderOptions options = new ReaderOptions
            {
                EnforceDateTimesAsNumbers = true
            };
            options.AddEnforcedColumn(1, columnType);
            AssertValues<object, object>(cells, options, AssertApproximate, expectedCells);
        }

        [Theory(DisplayName = "Test of the reader options for the import column type: Double")]
        [InlineData("B")]
        [InlineData(1)]
        public void EnforcingColumnAsNumberTest(object column)
        {
            TimeSpan time = new TimeSpan(11, 12, 13);
            DateTime date = new DateTime(2021, 8, 14, 18, 22, 13, 0);
            Dictionary<string, Object> cells = new Dictionary<string, object>
            {
                { "A1", 22 },
                { "A2", "21" },
                { "A3", true },
                { "B1", 23 },
                { "B2", "20" },
                { "B3", true },
                { "B4", time },
                { "B5", date },
                { "B6", null },
                { "B7", new Cell("=A1", Cell.CellType.Formula, "B7") },
                { "C1", "2" },
                { "C2", new TimeSpan(12, 14, 16) },
                { "C3", new Cell("=A1", Cell.CellType.Formula, "C3") }
            };
            Dictionary<string, object> expectedCells = new Dictionary<string, object>
            {
                { "A1", 22 },
                { "A2", "21" },
                { "A3", true },
                { "B1", 23d },
                { "B2", 20d },
                { "B3", 1d },
                { "B4", DataUtils.GetOATime(time) },
                { "B5", DataUtils.GetOADateTime(date) },
                { "B6", null },
                { "B7", new Cell("=A1", Cell.CellType.Formula, "B7") },
                { "C1", "2" },
                { "C2", new TimeSpan(12, 14, 16) },
                { "C3", new Cell("=A1", Cell.CellType.Formula, "C3") }
            };
            ReaderOptions options = new ReaderOptions();
            if (column is string)
            {
                options.AddEnforcedColumn(column as string, ReaderOptions.ColumnType.Double);
            }
            else
            {
                options.AddEnforcedColumn((int)column, ReaderOptions.ColumnType.Double);
            }
            AssertValues<object, object>(cells, options, AssertApproximate, expectedCells);
        }

        [Theory(DisplayName = "Test of the reader options for the import column type: Numeric")]
        [InlineData("B")]
        [InlineData(1)]
        public void EnforcingColumnAsNumberTest2(object column)
        {
            TimeSpan time = new TimeSpan(11, 12, 13);
            DateTime date = new DateTime(2021, 8, 14, 18, 22, 13, 0);
            Dictionary<string, Object> cells = new Dictionary<string, object>
            {
                { "A1", 22 },
                { "A2", "21" },
                { "A3", true },
                { "B1", 23 },
                { "B2", "20.1" },
                { "B3", time },
                { "B4", date },
                { "B5", null },
                { "B6", new Cell("=A1", Cell.CellType.Formula, "B6") },
                { "B7", "true" },
                { "B8", "false" },
                { "B9", true },
                { "B10", false },
                { "B11", "XYZ" },
                { "C1", "2" },
                { "C2", new TimeSpan(12, 14, 16) },
                { "C3", new Cell("=A1", Cell.CellType.Formula, "C3") }
            };
            Dictionary<string, object> expectedCells = new Dictionary<string, object>
            {
                { "A1", 22 },
                { "A2", "21" },
                { "A3", true },
                { "B1", 23 },
                { "B2", 20.1f },
                { "B3", DataUtils.GetOATime(time) },
                { "B4", DataUtils.GetOADateTime(date) },
                { "B5", null },
                { "B6", new Cell("=A1", Cell.CellType.Formula, "B6") },
                { "B7", 1 },
                { "B8", 0 },
                { "B9", 1 },
                { "B10", 0 },
                { "B11", "XYZ" },
                { "C1", "2" },
                { "C2", new TimeSpan(12, 14, 16) },
                { "C3", new Cell("=A1", Cell.CellType.Formula, "C3") }
            };
            ReaderOptions options = new ReaderOptions();
            if (column is string)
            {
                options.AddEnforcedColumn(column as string, ReaderOptions.ColumnType.Numeric);
            }
            else
            {
                options.AddEnforcedColumn((int)column, ReaderOptions.ColumnType.Numeric);
            }
            AssertValues<object, object>(cells, options, AssertApproximate, expectedCells);
        }

        [Theory(DisplayName = "Test of the reader options for the import column types Numeric, Decimal and Double on parsed dates and times")]
        [InlineData(ReaderOptions.ColumnType.Double, "2021-10-31 12:11:10", 44500.5077546296d)]
        [InlineData(ReaderOptions.ColumnType.Double, "18:20:22", 0.764143518518519d)]
        [InlineData(ReaderOptions.ColumnType.Decimal, "2021-10-31 12:11:10", "44500.5077546296")]
        [InlineData(ReaderOptions.ColumnType.Decimal, "18:20:22", "0.764143518518519")]
        [InlineData(ReaderOptions.ColumnType.Numeric, "2021-10-31 12:11:10", 44500.5077546296d)]
        [InlineData(ReaderOptions.ColumnType.Numeric, "18:20:22", 0.764143518518519d)]
        public void EnforcingColumnAsNumberTest3(ReaderOptions.ColumnType columnType, string givenValue, object expectedValue)
        {
            Dictionary<string, Object> cells = new Dictionary<string, object>
            {
                { "A1", true },
                { "B1", givenValue },
                { "C1", "2" }
            };

            Dictionary<string, object> expectedCells = new Dictionary<string, object>
            {
                { "A1", true }
            };
            if (columnType == ReaderOptions.ColumnType.Decimal)
            {
                // m-suffix is not working
                expectedCells.Add("B1", Convert.ToDecimal(expectedValue));
            }
            else
            {
                expectedCells.Add("B1", expectedValue);
            }
            expectedCells.Add("C1", "2");
            ReaderOptions options = new ReaderOptions
            {
                EnforceDateTimesAsNumbers = true
            };
            options.AddEnforcedColumn(1, columnType);
            AssertValues<object, object>(cells, options, AssertApproximate, expectedCells);
        }

        [Theory(DisplayName = "Test of the reader options for the import column type with wrong style information: Double and Decimal")]
        [InlineData("B", ReaderOptions.ColumnType.Double)]
        [InlineData(1, ReaderOptions.ColumnType.Double)]
        [InlineData("B", ReaderOptions.ColumnType.Decimal)]
        [InlineData(1, ReaderOptions.ColumnType.Decimal)]
        public void EnforcingColumnAsNumberTest4a(object column, ReaderOptions.ColumnType type)
        {
            object ob1, ob2, ob4, ob5;
            object exOb1, exOb2, exOb4, exOb5;
            string ob3 = "5-7";
            string ob6 = "1870-06-01 12:12:00";
            if (type == ReaderOptions.ColumnType.Double)
            {
                ob1 = -10d;
                ob2 = -5.5d;
                ob4 = -1d;
                ob5 = float.MaxValue;
                exOb1 = -10d;
                exOb2 = -5.5d;
                exOb4 = -1d;
                exOb5 = (double)float.MaxValue;
            }
            else
            {
                ob1 = -10m;
                ob2 = -5.5m;
                ob4 = -1m;
                ob5 = float.MaxValue;
                exOb1 = -10m;
                exOb2 = -5.5m;
                exOb4 = -1m;
                exOb5 = float.MaxValue;
            }
            Dictionary<string, Object> cells = new Dictionary<string, object>();
            Cell a1 = new Cell(1, Cell.CellType.Number, "A1");
            Cell b1 = new Cell(ob1, Cell.CellType.Number, "B1");
            b1.SetStyle(BasicStyles.DateFormat);
            Cell b2 = new Cell(ob2, Cell.CellType.Number, "B2");
            b1.SetStyle(BasicStyles.TimeFormat);
            Cell b3 = new Cell(ob3, Cell.CellType.String, "B3");
            b1.SetStyle(BasicStyles.DateFormat);
            Cell b4 = new Cell(ob4, Cell.CellType.String, "B4");
            b1.SetStyle(BasicStyles.DateFormat);
            Cell b5 = new Cell(ob5, Cell.CellType.Number, "B5");
            b1.SetStyle(BasicStyles.DateFormat);
            Cell b6 = new Cell(ob6, Cell.CellType.String, "B6");
            b5.SetStyle(BasicStyles.DateFormat);
            Cell c1 = new Cell(10, Cell.CellType.Number, "C1");
            cells.Add("A1", a1);
            cells.Add("B1", b1);
            cells.Add("B2", b2);
            cells.Add("B3", b3);
            cells.Add("B4", b4);
            cells.Add("B5", b5);
            cells.Add("B6", b6);
            cells.Add("C1", c1);
            Dictionary<string, Cell> expectedCells = new Dictionary<string, Cell>();
            Cell exA1 = new Cell(1, Cell.CellType.Number, "A1");
            Cell exB1 = new Cell(exOb1, Cell.CellType.Number, "B1");
            Cell exB2 = new Cell(exOb2, Cell.CellType.Number, "B2");
            Cell exB3 = new Cell(ob3, Cell.CellType.String, "B3");
            Cell exB4 = new Cell(exOb4, Cell.CellType.String, "B4");
            Cell exB5 = new Cell(exOb5, Cell.CellType.Number, "B5");
            Cell exB6 = new Cell(ob6, Cell.CellType.String, "B6");
            Cell exC1 = new Cell(10, Cell.CellType.Number, "C1");
            expectedCells.Add("A1", exA1);
            expectedCells.Add("B1", exB1);
            expectedCells.Add("B2", exB2);
            expectedCells.Add("B3", exB3);
            expectedCells.Add("B4", exB4);
            expectedCells.Add("B5", exB5);
            expectedCells.Add("B6", exB6);
            expectedCells.Add("C1", exC1);
            ReaderOptions options = new ReaderOptions();
            if (column is string)
            {
                options.AddEnforcedColumn(column as string, type);
            }
            else
            {
                options.AddEnforcedColumn((int)column, type);
            }
            AssertValues<object, Cell>(cells, options, AssertApproximate, expectedCells);
        }


        [Theory(DisplayName = "Test of the reader options for the import column type: Bool")]
        [InlineData("B")]
        [InlineData(1)]
        public void EnforcingColumnAsBoolTest(object column)
        {
            TimeSpan time = new TimeSpan(11, 12, 13);
            DateTime date = new DateTime(2021, 8, 14, 18, 22, 13, 0);
            Dictionary<string, Object> cells = new Dictionary<string, object>
            {
                { "A1", 1 },
                { "A2", "21" },
                { "A3", true },
                { "B1", 1 },
                { "B2", "true" },
                { "B3", false },
                { "B4", time },
                { "B5", date },
                { "B6", 0f },
                { "B7", "1" },
                { "B8", "Test" },
                { "B9", 1.0d },
                { "B10", null },
                { "B11", new Cell("=A1", Cell.CellType.Formula, "B11") },
                { "B12", 2 },
                { "B13", "0" },
                { "B14", "" },
                { "C1", "0" },
                { "C2", new TimeSpan(12, 14, 16) },
                { "C3", new Cell("=A1", Cell.CellType.Formula, "C3") }
            };
            Dictionary<string, object> expectedCells = new Dictionary<string, object>
            {
                { "A1", 1 },
                { "A2", "21" },
                { "A3", true },
                { "B1", true },
                { "B2", true },
                { "B3", false },
                { "B4", time },
                { "B5", date },
                { "B6", false },
                { "B7", true },
                { "B8", "Test" },
                { "B9", true },
                { "B10", null },
                { "B11", new Cell("=A1", Cell.CellType.Formula, "B11") },
                { "B12", 2 },
                { "B13", false },
                { "B14", "" },
                { "C1", "0" },
                { "C2", new TimeSpan(12, 14, 16) },
                { "C3", new Cell("=A1", Cell.CellType.Formula, "C3") }
            };
            ReaderOptions options = new ReaderOptions();
            if (column is string)
            {
                options.AddEnforcedColumn(column as string, ReaderOptions.ColumnType.Bool);
            }
            else
            {
                options.AddEnforcedColumn((int)column, ReaderOptions.ColumnType.Bool);
            }
            AssertValues<object, object>(cells, options, AssertApproximate, expectedCells);
        }

        [Theory(DisplayName = "Test of the reader options for the import column type: String")]
        [InlineData("B")]
        [InlineData(1)]
        public void EnforcingColumnAsStringTest(object column)
        {
            TimeSpan time = new TimeSpan(11, 12, 13);
            DateTime date = new DateTime(2021, 8, 14, 18, 22, 13, 0);
            Dictionary<string, Object> cells = new Dictionary<string, object>
            {
                { "A1", 1 },
                { "A2", "21" },
                { "A3", true },
                { "B1", 1 },
                { "B2", "Test" },
                { "B3", false },
                { "B4", time },
                { "B5", date },
                { "B6", 0f },
                { "B7", true },
                { "B8", -10 },
                { "B9", 1.111d },
                { "B10", null },
                { "B11", new Cell("=A1", Cell.CellType.Formula, "B11") },
                { "B12", 2147483650 },
                { "B13", 9223372036854775806 },
                { "B14", 18446744073709551614 },
                { "B15", (short)32766 },
                { "B16", (ushort)65534 },
                { "B17", 0.000000001d },
                { "B18", 0.123f },
                { "B19", (byte)17 },
                { "C1", "0" },
                { "C2", new TimeSpan(12, 14, 16) },
                { "C3", new Cell("=A1", Cell.CellType.Formula, "C3") }
            };
            Dictionary<string, object> expectedCells = new Dictionary<string, object>
            {
                { "A1", 1 },
                { "A2", "21" },
                { "A3", true },
                { "B1", "1" },
                { "B2", "Test" },
                { "B3", "False" },
                { "B4", time.ToString(ReaderOptions.DefaultTimeSpanFormat) },
                { "B5", date.ToString(ReaderOptions.DefaultDateTimeFormat) },
                { "B6", "0" },
                { "B7", "True" },
                { "B8", "-10" },
                { "B9", "1.111" },
                { "B10", null },
                { "B11", "=A1" },
                { "B12", "2147483650" },
                { "B13", "9223372036854775806" },
                { "B14", "18446744073709551614" },
                { "B15", "32766" },
                { "B16", "65534" },
                { "B17", "1E-09" }, // Currently handled without option to format the number
                { "B18", "0.123" },
                { "B19", "17" },
                { "C1", "0" },
                { "C2", new TimeSpan(12, 14, 16) },
                { "C3", new Cell("=A1", Cell.CellType.Formula, "C3") }
            };
            ReaderOptions options = new ReaderOptions();
            if (column is string)
            {
                options.AddEnforcedColumn(column as string, ReaderOptions.ColumnType.String);
            }
            else
            {
                options.AddEnforcedColumn((int)column, ReaderOptions.ColumnType.String);
            }
            AssertValues<object, object>(cells, options, AssertApproximate, expectedCells);
        }

        [Theory(DisplayName = "Test of the reader options for the import column type: Date")]
        [InlineData("B")]
        [InlineData(1)]
        public void EnforcingColumnAsDateTest(object column)
        {
            TimeSpan time = new TimeSpan(11, 12, 13);
            DateTime date = new DateTime(2021, 8, 14, 18, 22, 13, 0);
            Dictionary<string, Object> cells = new Dictionary<string, object>
            {
                { "A1", 1 },
                { "A2", "21" },
                { "A3", true },
                { "B1", 1 },
                { "B2", "Test" },
                { "B3", false },
                { "B4", time },
                { "B5", date },
                { "B6", 44494.5209490741d },
                { "B7", "2021-10-25 12:30:10" },
                { "B8", -10 },
                { "B9", 44494.5f },
                { "B10", null },
                { "B11", new Cell("=A1", Cell.CellType.Formula, "B11") },
                { "B12", 2147483650 },
                { "B13", 2958466 },
                { "C1", "0" },
                { "C2", new TimeSpan(12, 14, 16) },
                { "C3", new Cell("=A1", Cell.CellType.Formula, "C3") }
            };
            Dictionary<string, object> expectedCells = new Dictionary<string, object>
            {
                { "A1", 1 },
                { "A2", "21" },
                { "A3", true },
                { "B1", new DateTime(1900, 1, 1, 0, 0, 0, 0) },
                { "B2", "Test" },
                { "B3", false },
                { "B4", new DateTime(1900, 1, 1, 11, 12, 13, 0) },
                { "B5", new DateTime(2021, 8, 14, 18, 22, 13, 0) },
                { "B6", new DateTime(2021, 10, 25, 12, 30, 10, 0) },
                { "B7", new DateTime(2021, 10, 25, 12, 30, 10, 0) },
                { "B8", -10 },
                { "B9", new DateTime(2021, 10, 25, 12, 0, 0, 0) },
                { "B10", null },
                { "B11", new Cell("=A1", Cell.CellType.Formula, "B11") },
                { "B12", 2147483650 },
                { "B13", 2958466 }, // Exceeds year 9999
                { "C1", "0" },
                { "C2", new TimeSpan(12, 14, 16) },
                { "C3", new Cell("=A1", Cell.CellType.Formula, "C3") }
            };
            ReaderOptions options = new ReaderOptions();
            if (column is string)
            {
                options.AddEnforcedColumn(column as string, ReaderOptions.ColumnType.Date);
            }
            else
            {
                options.AddEnforcedColumn((int)column, ReaderOptions.ColumnType.Date);
            }
            AssertValues<object, object>(cells, options, AssertApproximate, expectedCells);
        }

        [Theory(DisplayName = "Test of the reader options for the import column type (without casting to numbers) with missing formats for DateTime and TimeSpan")]
        [InlineData("B", "C")]
        [InlineData(1, 2)]
        void enforcingColumnAsDateTest2(object column1, object column2)
        {

            Dictionary<string, object> cells = new Dictionary<string, object>
            {
                { "A1", 1 },
                { "B1", "11:12:13" },
                { "C1", "2021-08-14 18:22:13" },
                { "D1", "0" }
            };

            Dictionary<string, object> expectedCells = new Dictionary<string, object>
            {
                { "A1", 1 },
                { "B1", new TimeSpan(11, 12, 13) },
                { "C1", new DateTime(2021, 8, 14, 18, 22, 13) },
                { "D1", "0" }
            };
            ReaderOptions options = new ReaderOptions
            {
                DateTimeFormat = null,
                TimeSpanFormat = null
            };
            if (column1 is String)
            {
                options.AddEnforcedColumn(column1 as string, ReaderOptions.ColumnType.Time);
                options.AddEnforcedColumn(column2 as string, ReaderOptions.ColumnType.Date);
            }
            else
            {
                options.AddEnforcedColumn((int)column1, ReaderOptions.ColumnType.Time);
                options.AddEnforcedColumn((int)column2, ReaderOptions.ColumnType.Date);
            }
            AssertValues<object, object>(cells, options, AssertApproximate, expectedCells);
        }

        [Theory(DisplayName = "Test of the reader options for the import column type with wrong style information: Date")]
        [InlineData("B", ReaderOptions.ColumnType.Date)]
        [InlineData(1, ReaderOptions.ColumnType.Date)]
        [InlineData("B", ReaderOptions.ColumnType.Time)]
        [InlineData(1, ReaderOptions.ColumnType.Time)]
        public void EnforcingColumnAsDateTest2(object column, ReaderOptions.ColumnType type)
        {
            Dictionary<string, Object> cells = new Dictionary<string, object>();
            Cell a1 = new Cell(1, Cell.CellType.Number, "A1");
            Cell b1 = new Cell(-10, Cell.CellType.Number, "B1");
            b1.SetStyle(BasicStyles.DateFormat);
            Cell b2 = new Cell(-5.5f, Cell.CellType.Number, "B2");
            b2.SetStyle(BasicStyles.TimeFormat);
            Cell b3 = new Cell("5-7", Cell.CellType.String, "B3");
            b3.SetStyle(BasicStyles.DateFormat);
            Cell b4 = new Cell("-1", Cell.CellType.String, "B4");
            b4.SetStyle(BasicStyles.TimeFormat);
            Cell b5 = new Cell("1870-06-06 12:12:00", Cell.CellType.String, "B5");
            b5.SetStyle(BasicStyles.DateFormat);
            Cell c1 = new Cell(10, Cell.CellType.Number, "C1");
            cells.Add("A1", a1);
            cells.Add("B1", b1);
            cells.Add("B2", b2);
            cells.Add("B3", b3);
            cells.Add("B4", b4);
            cells.Add("B5", b5);
            cells.Add("C1", c1);
            Dictionary<string, Cell> expectedCells = new Dictionary<string, Cell>();
            Cell exA1 = new Cell(1, Cell.CellType.Number, "A1");
            Cell exB1 = new Cell(-10, Cell.CellType.Number, "B1");
            Cell exB2 = new Cell(-5.5f, Cell.CellType.Number, "B2");
            Cell exB3 = new Cell("5-7", Cell.CellType.String, "B3");
            Cell exB4 = new Cell("-1", Cell.CellType.String, "B4");
            Cell exB5 = new Cell("1870-06-06 12:12:00", Cell.CellType.String, "B5");
            Cell exC1 = new Cell(10, Cell.CellType.Number, "C1");
            expectedCells.Add("A1", exA1);
            expectedCells.Add("B1", exB1);
            expectedCells.Add("B2", exB2);
            expectedCells.Add("B3", exB3);
            expectedCells.Add("B4", exB4);
            expectedCells.Add("B5", exB5);
            expectedCells.Add("C1", exC1);
            ReaderOptions options = new ReaderOptions();
            if (column is string)
            {
                options.AddEnforcedColumn(column as string, type);
            }
            else
            {
                options.AddEnforcedColumn((int)column, type);
            }
            AssertValues<object, Cell>(cells, options, AssertApproximate, expectedCells);
        }

        [Theory(DisplayName = "Test of the reader options for the import column type: Time")]
        [InlineData("B")]
        [InlineData(1)]
        public void EnforcingColumnAsTimeTest(object column)
        {
            TimeSpan time = new TimeSpan(11, 12, 13);
            DateTime date = new DateTime(2021, 8, 14, 18, 22, 13, 0);
            Dictionary<string, Object> cells = new Dictionary<string, object>
            {
                { "A1", 1 },
                { "A2", "21" },
                { "A3", true },
                { "B1", 1 },
                { "B2", "Test" },
                { "B3", false },
                { "B4", time },
                { "B5", date },
                { "B6", 44494.5209490741d },
                { "B7", "2021-10-25 12:30:10" },
                { "B8", -10 },
                { "B9", 44494.5f },
                { "B10", null },
                { "B11", new Cell("=A1", Cell.CellType.Formula, "B11") },
                { "C1", "0" },
                { "C2", new TimeSpan(12, 14, 16) },
                { "C3", new Cell("=A1", Cell.CellType.Formula, "C3") }
            };
            Dictionary<string, object> expectedCells = new Dictionary<string, object>
            {
                { "A1", 1 },
                { "A2", "21" },
                { "A3", true },
                { "B1", new TimeSpan(1, 0, 0, 0) },
                { "B2", "Test" },
                { "B3", false },
                { "B4", time },
                { "B5", new TimeSpan(44422, 18, 22, 13) },
                { "B6", new TimeSpan(44494, 12, 30, 10) },
                { "B7", new TimeSpan(44494, 12, 30, 10) },
                { "B8", -10 },
                { "B9", new TimeSpan(44494, 12, 0, 0) },
                { "B10", null },
                { "B11", new Cell("=A1", Cell.CellType.Formula, "B11") },
                { "C1", "0" },
                { "C2", new TimeSpan(12, 14, 16) },
                { "C3", new Cell("=A1", Cell.CellType.Formula, "C3") }
            };
            ReaderOptions options = new ReaderOptions();
            if (column is string)
            {
                options.AddEnforcedColumn(column as string, ReaderOptions.ColumnType.Time);
            }
            else
            {
                options.AddEnforcedColumn((int)column, ReaderOptions.ColumnType.Time);
            }
            AssertValues<object, object>(cells, options, AssertApproximate, expectedCells);
        }

        [Theory(DisplayName = "Test of the reader options for the combination of a start row and a enforced column")]
        [InlineData(ReaderOptions.ColumnType.Bool, "1", true)]
        [InlineData(ReaderOptions.ColumnType.Bool, false, false)]
        [InlineData(ReaderOptions.ColumnType.Double, "-2.5", -2.5d)]
        [InlineData(ReaderOptions.ColumnType.Double, 13, 13d)]
        [InlineData(ReaderOptions.ColumnType.Numeric, "12.5", 12.5f)]
        [InlineData(ReaderOptions.ColumnType.Numeric, 13, 13)]
        [InlineData(ReaderOptions.ColumnType.String, 16.5f, "16.5")]
        [InlineData(ReaderOptions.ColumnType.String, true, "True")]
        public void EnforcingColumnStartRowTest(ReaderOptions.ColumnType columnType, object givenValue, object expectedValue)
        {
            TimeSpan time = new TimeSpan(11, 12, 13);
            Dictionary<string, Object> cells = new Dictionary<string, object>
            {
                { "A1", "test" },
                { "A2", 23 },
                { "A3", time },
                { "B1", null },
                { "B2", givenValue },
                { "B3", givenValue },
                { "C1", 28 },
                { "C2", false },
                { "C3", "Test" }
            };
            Dictionary<string, object> expectedCells = new Dictionary<string, object>
            {
                { "A1", "test" },
                { "A2", 23 },
                { "A3", time },
                { "B1", null },
                { "B2", givenValue },
                { "B3", expectedValue },
                { "C1", 28 },
                { "C2", false },
                { "C3", "Test" }
            };
            ReaderOptions options = new ReaderOptions();
            options.AddEnforcedColumn(1, columnType);
            options.EnforcingStartRowNumber = 2;
            AssertValues<object, object>(cells, options, AssertApproximate, expectedCells);
            ReaderOptions options2 = new ReaderOptions();
            options2.AddEnforcedColumn("B", columnType);
            options2.EnforcingStartRowNumber = 2;
            AssertValues<object, object>(cells, options2, AssertApproximate, expectedCells);
        }

        [Theory(DisplayName = "Test of the reader options for the combination of a start row and a enforced column on types Date and Time")]
        [InlineData(ReaderOptions.ColumnType.Date)]
        [InlineData(ReaderOptions.ColumnType.Time)]
        public void EnforcingColumnStartRowTest2(ReaderOptions.ColumnType columnType)
        {
            TimeSpan time = new TimeSpan(11, 12, 13);
            TimeSpan expectedTime = new TimeSpan(12, 13, 14);
            DateTime expectedDate = new DateTime(2021, 8, 14, 18, 22, 13, 0);

            Dictionary<string, Object> cells = new Dictionary<string, object>
            {
                { "A1", "test" },
                { "A2", 23 },
                { "A3", time },
                { "B1", null }
            };
            if (columnType == ReaderOptions.ColumnType.Time)
            {
                cells.Add("B2", "12:13:14");
                cells.Add("B3", "12:13:14");
            }
            else if (columnType == ReaderOptions.ColumnType.Date)
            {
                cells.Add("B2", "2021-08-14 18:22:13");
                cells.Add("B3", "2021-08-14 18:22:13");
            }
            cells.Add("C1", 28);
            cells.Add("C2", false);
            cells.Add("C3", "Test");
            Dictionary<string, object> expectedCells = new Dictionary<string, object>
            {
                { "A1", "test" },
                { "A2", 23 },
                { "A3", time },
                { "B1", null }
            };
            if (columnType == ReaderOptions.ColumnType.Time)
            {
                expectedCells.Add("B2", "12:13:14");
                expectedCells.Add("B3", expectedTime);
            }
            else if (columnType == ReaderOptions.ColumnType.Date)
            {
                expectedCells.Add("B2", "2021-08-14 18:22:13");
                expectedCells.Add("B3", expectedDate);
            }
            expectedCells.Add("C1", 28);
            expectedCells.Add("C2", false);
            expectedCells.Add("C3", "Test");
            ReaderOptions options = new ReaderOptions();
            options.AddEnforcedColumn(1, columnType);
            options.EnforcingStartRowNumber = 2;
            AssertValues<object, object>(cells, options, AssertApproximate, expectedCells);
            ReaderOptions options2 = new ReaderOptions();
            options2.AddEnforcedColumn("B", columnType);
            options2.EnforcingStartRowNumber = 2;
            AssertValues<object, object>(cells, options2, AssertApproximate, expectedCells);
        }

        [Theory(DisplayName = "Test of enforced import types when the same type overlaps globally and on a column")]
        [InlineData(ReaderOptions.GlobalType.AllNumbersToDecimal, ReaderOptions.ColumnType.Decimal, "7", "23", "1.1", typeof(decimal))]
        [InlineData(ReaderOptions.GlobalType.AllNumbersToDouble, ReaderOptions.ColumnType.Double, "7", "23", "1.1", typeof(double))]
        [InlineData(ReaderOptions.GlobalType.AllNumbersToInt, ReaderOptions.ColumnType.Numeric, "7", "23", "1.1", typeof(int))]
        [InlineData(ReaderOptions.GlobalType.EverythingToString, ReaderOptions.ColumnType.String, "7", "23", "1.1", typeof(string))]
        public void ImportEnforceOverlappingTest(ReaderOptions.GlobalType globalType, ReaderOptions.ColumnType columnType, string givenA2Value, string givenB1Value, string givenB2Value, Type expectedType)
        {
            Dictionary<string, object> cells = new Dictionary<string, object>();
            Dictionary<string, object> expectedCells = new Dictionary<string, object>();
            cells.Add("A1", "test");
            cells.Add("A2", givenA2Value);
            cells.Add("B1", givenB1Value);
            cells.Add("B2", givenB2Value);
            expectedCells.Add("A1", "test");
            expectedCells.Add("A2", TestUtils.CreateInstance(expectedType, givenA2Value));
            expectedCells.Add("B1", TestUtils.CreateInstance(expectedType, givenB1Value));
            expectedCells.Add("B2", TestUtils.CreateInstance(expectedType, givenB2Value));
            ReaderOptions ReaderOptions = new ReaderOptions();
            ReaderOptions.AddEnforcedColumn(1, columnType);
            ReaderOptions.GlobalEnforcingType = globalType;
            AssertValues<object, object>(cells, ReaderOptions, AssertApproximate, expectedCells);
        }


        [Theory(DisplayName = "Test of enforced import types when the global type overrules the column type")]
        [InlineData(ReaderOptions.ColumnType.Decimal, ReaderOptions.GlobalType.AllNumbersToDouble, typeof(decimal), "7", typeof(double), "7")]
        [InlineData(ReaderOptions.ColumnType.Double, ReaderOptions.GlobalType.AllNumbersToDecimal, typeof(double), "7", typeof(decimal), "7")]
        public void ImportEnforceOverruleTest(ReaderOptions.ColumnType columnType, ReaderOptions.GlobalType globalType, Type givenType, string givenValue, Type expectedType, string expectedValue)
        {
            Dictionary<string, object> cells = new Dictionary<string, object>();
            Dictionary<string, object> expectedCells = new Dictionary<string, object>();
            cells.Add("A1", TestUtils.CreateInstance(givenType, givenValue));
            expectedCells.Add("A1", TestUtils.CreateInstance(expectedType, expectedValue));
            ReaderOptions ReaderOptions = new ReaderOptions();
            ReaderOptions.AddEnforcedColumn(1, columnType);
            ReaderOptions.GlobalEnforcingType = globalType;
            AssertValues<object, object>(cells, ReaderOptions, AssertApproximate, expectedCells);
        }


        [Theory(DisplayName = "Test of the reader options for custom date and time formats and culture info")]
        [InlineData(ReaderOptions.ColumnType.Date, "en-US", "yyyy-MM-dd HH:mm:ss", "2021-08-12 12:11:10", "2021-08-12 12:11:10")]
        [InlineData(ReaderOptions.ColumnType.Date, "de-DE", "dd.MM.yyyy HH:mm:ss", "12.08.2021 12:11:10", "2021-08-12 12:11:10")]
        [InlineData(ReaderOptions.ColumnType.Date, "fr-FR", "dd/MM/yyyy", "12/08/2021", "2021-08-12 00:00:00")]
        [InlineData(ReaderOptions.ColumnType.Date, null, null, "12.08.2021 12:11:10", "2021-12-08 12:11:10")]
        [InlineData(ReaderOptions.ColumnType.Time, "en-US", "hh\\:mm\\:ss", "18:11:10", "18:11:10")]
        [InlineData(ReaderOptions.ColumnType.Time, "", "hh", "12", "12:00:00")]
        [InlineData(ReaderOptions.ColumnType.Time, null, null, "18:11:10", "18:11:10")]
        public void ParseDateTimeTest(ReaderOptions.ColumnType columnType, string cultureInfo, string pattern, string givenValue, string expectedValue)
        {
            Dictionary<string, Object> cells = new Dictionary<string, object>();
            Dictionary<string, object> expectedCells = new Dictionary<string, object>();
            ReaderOptions ReaderOptions = new ReaderOptions();
            if (columnType == ReaderOptions.ColumnType.Date)
            {
                DateTime expected = DateTime.ParseExact(expectedValue, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                expectedCells.Add("A1", expected);
                ReaderOptions.DateTimeFormat = pattern;
                ReaderOptions.AddEnforcedColumn(0, ReaderOptions.ColumnType.Date);
            }
            else
            {
                TimeSpan expected = TimeSpan.ParseExact(expectedValue, "hh\\:mm\\:ss", CultureInfo.InvariantCulture);
                expectedCells.Add("A1", expected);
                ReaderOptions.TimeSpanFormat = pattern;
                ReaderOptions.AddEnforcedColumn(0, ReaderOptions.ColumnType.Time);
            }
            if (cultureInfo != null)
            {
                CultureInfo givenCultureInfo = new CultureInfo(cultureInfo); // empty will lead to invariant
                ReaderOptions.TemporalCultureInfo = givenCultureInfo;
            }
            cells.Add("A1", givenValue);
            AssertValues<object, object>(cells, ReaderOptions, AssertApproximate, expectedCells);
        }

        [Theory(DisplayName = "Test of the failing casting on an invalid Date or TimeSpan value")]
        [InlineData(ReaderOptions.ColumnType.Time, "55.81.202x")]
        [InlineData(ReaderOptions.ColumnType.Date, "2022-18-22 15:6x:00")]
        [InlineData(ReaderOptions.ColumnType.Time, "10000-01-01 00:00:00")]
        [InlineData(ReaderOptions.ColumnType.Date, "10000-01-01 00:00:00")]
        [InlineData(ReaderOptions.ColumnType.Time, "1800-01-01 00:00:00")]
        [InlineData(ReaderOptions.ColumnType.Time, "-10:00:00")]
        void InvalidDateCastingTest(ReaderOptions.ColumnType columnType, string value)
        {
            ReaderOptions options = new ReaderOptions
            {
                EnforceDateTimesAsNumbers = true
            };
            options.AddEnforcedColumn("A", columnType);
            Dictionary<string, object> cells = new Dictionary<string, object>();
            Dictionary<string, object> expectedCells = new Dictionary<string, object>();
            cells.Add("A1", value);
            expectedCells.Add("A1", value);
            AssertValues<object, object>(cells, options, AssertApproximate, expectedCells);
        }

        [Fact(DisplayName = "Test of the failing casting on an TimeSpan with an invalid (too high) number of days")]
        void InvalidDateCastingTest2()
        {
            ReaderOptions options = new ReaderOptions
            {
                EnforceDateTimesAsNumbers = true
            };
            options.AddEnforcedColumn("A", ReaderOptions.ColumnType.Time);
            options.TimeSpanFormat = "HH:mm:ss d";
            Dictionary<string, object> cells = new Dictionary<string, object>();
            Dictionary<string, object> expectedCells = new Dictionary<string, object>();
            cells.Add("A1", "00:00:00 2958467");
            expectedCells.Add("A1", "00:00:00 2958467");
            AssertValues<object, object>(cells, options, AssertApproximate, expectedCells);
        }

        [Theory(DisplayName = "Test of the reader option to process or discard phonetic characters in strings")]
        [InlineData(0, 1, false)]
        [InlineData(0, 2, true)]
        void PhnoneticCharactersImportOptionTest(int givenValuesColumn, int expectedValuesColumn, bool importOptionValue)
        {
            // Note: Cells in column A contains the strings with phonetic characters.
            // The corresponding Cells in column B contains the values without enabled import option.
            // The cells in column C contains the values with the import option enabled.
            // The values starts at Row 2 (Index 1)

            ReaderOptions options = new ReaderOptions
            {
                EnforcePhoneticCharacterImport = importOptionValue
            };
            Stream stream = TestUtils.GetResource("phonetics.xlsx");
            Workbook workbook = WorkbookReader.Load(stream, options);

            int lastRow = workbook.Worksheets[0].GetLastDataRowNumber();
            for (int r = 1; r <= lastRow; r++)
            {
                string given = workbook.Worksheets[0].GetCell(new Address(givenValuesColumn, r)).Value.ToString();
                string expected = workbook.Worksheets[0].GetCell(new Address(expectedValuesColumn, r)).Value.ToString();
                Assert.Equal(expected, given);
            }
        }

        [Theory(DisplayName = "Test of the reader option to ignore not supported password algorithms for worksheet protection")]
        [InlineData(false, true)]
        [InlineData(true, false)]
        void IgnoreNotSupportedPasswordAlgorithmsTest(bool importOptionValue, bool expectedError)
        {
            ReaderOptions options = new ReaderOptions
            {
                IgnoreNotSupportedPasswordAlgorithms = importOptionValue
            };
            Stream stream = TestUtils.GetResource("contemporary_password.xlsx");
            if (expectedError)
            {
                Assert.Throws<Exceptions.NotSupportedContentException>(() => WorkbookReader.Load(stream, options));
            }
            else
            {
                Workbook workbook = WorkbookReader.Load(stream, options);
                Assert.NotNull(workbook);
            }
        }

        [Theory(DisplayName = "Test of the reader option to ignore not supported password algorithms for workbook protection")]
        [InlineData(false, true)]
        [InlineData(true, false)]
        void IgnoreNotSupportedPasswordAlgorithmsTest2(bool importOptionValue, bool expectedError)
        {
            ReaderOptions options = new ReaderOptions
            {
                IgnoreNotSupportedPasswordAlgorithms = importOptionValue
            };
            Stream stream = TestUtils.GetResource("contemporary_password2.xlsx");
            if (expectedError)
            {
                Assert.Throws<Exceptions.NotSupportedContentException>(() => WorkbookReader.Load(stream, options));
            }
            else
            {
                Workbook workbook = WorkbookReader.Load(stream, options);
                Assert.NotNull(workbook);
            }
        }

        [Theory(DisplayName = "Test of the reader option property EnforceStrictValidation")]
        [InlineData("valid_column_row_dimensions.xlsx", true, false, 0)]
        [InlineData("invalid_column_width_min.xlsx", true, true, -1)]
        [InlineData("invalid_column_width_max.xlsx", true, true, 1)]
        [InlineData("invalid_row_height_min.xlsx", true, true, 0)]
        [InlineData("invalid_row_height_max.xlsx", true, true, 0)]
        [InlineData("valid_column_row_dimensions.xlsx", false, false, 0)]
        [InlineData("invalid_row_height_min.xlsx", false, false, 0)]
        [InlineData("invalid_row_height_max.xlsx", false, false, 0)]
        [InlineData("invalid_column_width_min.xlsx", false, false, -1)]
        [InlineData("invalid_column_width_max.xlsx", false, false, 1)]
        public void EnforceValidColumnDimensionsTest(string fileName, bool givenOptionValue, bool expectedThrow, int columnFlag)
        {
            ReaderOptions options = new ReaderOptions
            {
                EnforceStrictValidation = givenOptionValue
            };
            using Stream stream = TestUtils.GetResource(fileName);

            if (expectedThrow)
            {
                Assert.ThrowsAny<Exception>(() => WorkbookReader.Load(stream, options));
            }
            else
            {
                Workbook workbook = WorkbookReader.Load(stream, options);
                if (columnFlag == -1)
                {
                    Assert.Equal(Worksheet.MinColumnWidth, workbook.GetWorksheet(0).Columns[0].Width);
                }
                else if (columnFlag == 1)
                {
                    Assert.Equal(Worksheet.MaxColumnWidth, workbook.GetWorksheet(0).Columns[0].Width);
                }
                else
                {
                    Assert.True(true);
                }
            }
        }

        private static void AssertValues<T, D>(Dictionary<string, T> givenCells, ReaderOptions ReaderOptions, Action<object, object> assertionAction, Dictionary<string, D> expectedCells = null)
        {
            Workbook workbook = new Workbook("worksheet1");
            foreach (KeyValuePair<string, T> cell in givenCells)
            {
                workbook.CurrentWorksheet.AddCell(cell.Value, cell.Key);
            }
            MemoryStream stream = new MemoryStream();
            workbook.SaveAsStream(stream, true);
            stream.Position = 0;
            Workbook givenWorkbook = WorkbookReader.Load(stream, ReaderOptions);

            Assert.NotNull(givenWorkbook);
            Worksheet givenWorksheet = givenWorkbook.SetCurrentWorksheet(0);
            Assert.Equal("worksheet1", givenWorksheet.SheetName);
            foreach (string address in givenCells.Keys)
            {
                Cell givenCell = givenWorksheet.GetCell(new Address(address));
                D expectedValue = expectedCells[address];
                if (expectedValue == null)
                {
                    Assert.Equal(Cell.CellType.Empty, givenCell.DataType);
                }
                else if (expectedValue is Cell && (givenCell != null))
                {
                    assertionAction.Invoke((expectedValue as Cell).Value, (givenCell).Value);
                }
                //  else if (expectedValue is Cell)
                //  {
                //      assertionAction.Invoke((D)(expectedValue as Cell).Value, (D)givenCell.Value);
                //      Assert.Equal(Cell.CellType.Formula, givenCell.DataType);
                //  }
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
            double doubleThreshold = 0.000012; // The precision may vary (roughly one second)
            decimal decimalThreshold = 0.00000012m;
            if (given is decimal)
            {
                Assert.True(Math.Abs((decimal)given - (decimal)expected) < decimalThreshold);
            }
            else if (given is double)
            {
                Assert.True(Math.Abs((double)given - (double)expected) < doubleThreshold);
            }
            else if (given is float)
            {
                Assert.True(Math.Abs((float)given - (float)expected) < doubleThreshold);
            }
            else if (given is DateTime)
            {
                double e = DataUtils.GetOADateTime((DateTime)expected);
                double g = DataUtils.GetOADateTime((DateTime)given);
                AssertApproximate(e, g);
            }
            else if (given is TimeSpan)
            {
                double g = DataUtils.GetOATime((TimeSpan)given);
                double e = DataUtils.GetOATime((TimeSpan)expected);
                AssertApproximate(e, g);
            }
            else
            {
                AssertEquals<object>(expected, given);
            }

        }

    }
}
