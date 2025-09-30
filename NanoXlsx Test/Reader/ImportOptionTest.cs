using NanoXLSX;
using NanoXLSX.Styles;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Xunit;

namespace NanoXLSX_Test.Reader
{
    public class ImportOptionTest
    {


        [Fact(DisplayName = "Test of the reader functionality with the global import option to cast everything to string")]
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
            cells.Add("A10", new Cell("=A1", Cell.CellType.FORMULA, "A10"));
            Dictionary<string, string> expectedCells = new Dictionary<string, string>();
            expectedCells.Add("A1", "test");
            expectedCells.Add("A2", "True");
            expectedCells.Add("A3", "False");
            expectedCells.Add("A4", "42");
            expectedCells.Add("A5", "0.55");
            expectedCells.Add("A6", "-0.111");
            expectedCells.Add("A7", "2020-11-10 09:08:07");
            expectedCells.Add("A8", "18:15:12");
            expectedCells.Add("A9", null); // Empty remains empty
            expectedCells.Add("A10", "=A1");

            ImportOptions options = new ImportOptions();
            options.GlobalEnforcingType = ImportOptions.GlobalType.EverythingToString;
            AssertValues<object, string>(cells, options, AssertEquals, expectedCells);
        }


        [Fact(DisplayName = "Test of the reader functionality with the global import option to cast all number to decimal")]
        public void CastToDecimalTest()
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
            cells.Add("A11", new Cell("=A1", Cell.CellType.FORMULA, "A11"));
            Dictionary<string, object> expectedCells = new Dictionary<string, object>();
            expectedCells.Add("A1", "test");
            expectedCells.Add("A2", decimal.One);
            expectedCells.Add("A3", decimal.Zero);
            expectedCells.Add("A4", 42m);
            expectedCells.Add("A5", 0.55m);
            expectedCells.Add("A6", -0.111m);
            expectedCells.Add("A7", (decimal)Utils.GetOADateTime(new DateTime(2020, 11, 10, 9, 8, 7, 0)));
            expectedCells.Add("A8", (decimal)Utils.GetOATime(new TimeSpan(18, 15, 12)));
            expectedCells.Add("A9", null);
            expectedCells.Add("A10", 27m);
            expectedCells.Add("A11", new Cell("=A1", Cell.CellType.FORMULA, "A11"));
            ImportOptions options = new ImportOptions();
            options.GlobalEnforcingType = ImportOptions.GlobalType.AllNumbersToDecimal;
            AssertValues<object, object>(cells, options, AssertApproximate, expectedCells);
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
            cells.Add("A11", new Cell("=A1", Cell.CellType.FORMULA, "A11"));
            Dictionary<string, object> expectedCells = new Dictionary<string, object>();
            expectedCells.Add("A1", "test");
            expectedCells.Add("A2", 1d);
            expectedCells.Add("A3", 0d);
            expectedCells.Add("A4", 42d);
            expectedCells.Add("A5", 0.55d);
            expectedCells.Add("A6", -0.111d);
            expectedCells.Add("A7", Utils.GetOADateTime(new DateTime(2020, 11, 10, 9, 8, 7, 0)));
            expectedCells.Add("A8", Utils.GetOATime(new TimeSpan(18, 15, 12)));
            expectedCells.Add("A9", null);
            expectedCells.Add("A10", 27d);
            expectedCells.Add("A11", new Cell("=A1", Cell.CellType.FORMULA, "A11"));
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
            cells.Add("A8", new TimeSpan(18, 15, 12));
            cells.Add("A9", -4.9f);
            cells.Add("A10", 0.49d);
            cells.Add("A11", null);
            cells.Add("A12", "28");
            cells.Add("A13", new Cell("=A1", Cell.CellType.FORMULA, "A13"));
            cells.Add("A14", 8589934592L);
            cells.Add("A15", 2147483650.6f);
            cells.Add("A16", 4294967294u);
            cells.Add("A17", 18446744073709551614);
            Dictionary<string, object> expectedCells = new Dictionary<string, object>();
            expectedCells.Add("A1", "test");
            expectedCells.Add("A2", 1);
            expectedCells.Add("A3", 0);
            expectedCells.Add("A4", 42);
            expectedCells.Add("A5", 1);
            expectedCells.Add("A6", -3);
            expectedCells.Add("A7", (int)Math.Round(Utils.GetOADateTime(new DateTime(2020, 11, 10, 9, 8, 7, 0)), 0));
            expectedCells.Add("A8", (int)Math.Round(Utils.GetOATime(new TimeSpan(18, 15, 12)), 0));
            expectedCells.Add("A9", -5);
            expectedCells.Add("A10", 0);
            expectedCells.Add("A11", null);
            expectedCells.Add("A12", 28);
            expectedCells.Add("A13", new Cell("=A1", Cell.CellType.FORMULA, "A13"));
            expectedCells.Add("A14", 8589934592L);
            expectedCells.Add("A15", 2147483650.6f);
            expectedCells.Add("A16", 4294967294u);
            expectedCells.Add("A17", 18446744073709551614);
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
            cells.Add("A6", new Cell("=A1", Cell.CellType.FORMULA, "A6"));
            Dictionary<string, object> expectedCells = new Dictionary<string, object>();
            expectedCells.Add("A1", "test");
            expectedCells.Add("A2", true);
            expectedCells.Add("A3", 22.2f); // Import will go to the smallest float unit (float 32 / single)
            expectedCells.Add("A4", "");
            expectedCells.Add("A5", "");
            expectedCells.Add("A6", new Cell("=A1", Cell.CellType.FORMULA, "A6"));
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
            cells.Add("A3", new Cell("=A1", Cell.CellType.FORMULA, "A3"));
            cells.Add("A4", 22);
            cells.Add("A5", true);
            cells.Add("A6", 22.5d);
            cells.Add("A7", new Cell("=A1", Cell.CellType.FORMULA, "A7"));
            Dictionary<string, object> expectedCells = new Dictionary<string, object>();
            expectedCells.Add("A1", 22);
            expectedCells.Add("A2", true);
            expectedCells.Add("A3", new Cell("=A1", Cell.CellType.FORMULA, "A3"));
            expectedCells.Add("A4", "22");
            expectedCells.Add("A5", "True");
            expectedCells.Add("A6", "22.5");
            expectedCells.Add("A7", "=A1");
            ImportOptions options = new ImportOptions();
            options.EnforcingStartRowNumber = 3;
            options.GlobalEnforcingType = ImportOptions.GlobalType.EverythingToString;
            AssertValues<object, object>(cells, options, AssertApproximate, expectedCells);
        }

        [Fact(DisplayName = "Test of the EnforceDateTimesAsNumbers functionality on global enforcing rules")]
        public void EnforceDateTimesAsNumbersTest()
        {
            DateTime date = new DateTime(2021, 8, 17, 11, 12, 13, 0);
            TimeSpan time = new TimeSpan(18, 14, 10);
            Dictionary<string, Object> cells = new Dictionary<string, object>();
            cells.Add("A1", 22);
            cells.Add("A2", true);
            cells.Add("A3", date);
            cells.Add("A4", time);
            cells.Add("A5", 22.5f);
            cells.Add("A6", new Cell("=A1", Cell.CellType.FORMULA, "A6"));
            Dictionary<string, object> expectedCells = new Dictionary<string, object>();
            expectedCells.Add("A1", 22);
            expectedCells.Add("A2", true);
            expectedCells.Add("A3", Utils.GetOADateTime(date));
            expectedCells.Add("A4", Utils.GetOATime(time));
            expectedCells.Add("A5", 22.5f);
            expectedCells.Add("A6", new Cell("=A1", Cell.CellType.FORMULA, "A6"));
            ImportOptions options = new ImportOptions();
            options.EnforceDateTimesAsNumbers = true;
            AssertValues<object, object>(cells, options, AssertApproximate, expectedCells);
        }

        [Theory(DisplayName = "Test of the EnforceDateTimesAsNumbers functionality for the import column types: Date or Time")]
        [InlineData(ImportOptions.ColumnType.Date, 22.5f, 22.5d)]
        [InlineData(ImportOptions.ColumnType.Time, 22.5d, 22.5d)]
        public void EnforceDateTimesAsNumbersTest2(ImportOptions.ColumnType columnType, object givenLowNumber, object expectedLowNumber)
        {
            DateTime date = new DateTime(2021, 8, 17, 11, 12, 13, 0);
            TimeSpan time = new TimeSpan(18, 14, 10);
            Dictionary<string, Object> cells = new Dictionary<string, object>();
            cells.Add("A1", 22);
            cells.Add("A2", true);
            cells.Add("A3", date);
            cells.Add("A4", time);
            cells.Add("B1", date);
            cells.Add("B2", time);
            cells.Add("B3", givenLowNumber);
            cells.Add("B4", new Cell("=A1", Cell.CellType.FORMULA, "B4"));
            Dictionary<string, object> expectedCells = new Dictionary<string, object>();
            expectedCells.Add("A1", 22);
            expectedCells.Add("A2", true);
            expectedCells.Add("A3", Utils.GetOADateTime(date));
            expectedCells.Add("A4", Utils.GetOATime(time));
            expectedCells.Add("B1", Utils.GetOADateTime(date));
            expectedCells.Add("B2", Utils.GetOATime(time));
            expectedCells.Add("B3", expectedLowNumber);
            expectedCells.Add("B4", new Cell("=A1", Cell.CellType.FORMULA, "B4"));
            ImportOptions options = new ImportOptions();
            options.EnforceDateTimesAsNumbers = true;
            options.AddEnforcedColumn(1, columnType);
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
            cells.Add("B6", null);
            cells.Add("B7", new Cell("=A1", Cell.CellType.FORMULA, "B7"));
            cells.Add("C1", "2");
            cells.Add("C2", new TimeSpan(12, 14, 16));
            cells.Add("C3", new Cell("=A1", Cell.CellType.FORMULA, "C3"));
            Dictionary<string, object> expectedCells = new Dictionary<string, object>();
            expectedCells.Add("A1", 22);
            expectedCells.Add("A2", "21");
            expectedCells.Add("A3", true);
            expectedCells.Add("B1", 23d);
            expectedCells.Add("B2", 20d);
            expectedCells.Add("B3", 1d);
            expectedCells.Add("B4", Utils.GetOATime(time));
            expectedCells.Add("B5", Utils.GetOADateTime(date));
            expectedCells.Add("B6", null);
            expectedCells.Add("B7", new Cell("=A1", Cell.CellType.FORMULA, "B7"));
            expectedCells.Add("C1", "2");
            expectedCells.Add("C2", new TimeSpan(12, 14, 16));
            expectedCells.Add("C3", new Cell("=A1", Cell.CellType.FORMULA, "C3"));
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
            cells.Add("B3", time);
            cells.Add("B4", date);
            cells.Add("B5", null);
            cells.Add("B6", new Cell("=A1", Cell.CellType.FORMULA, "B6"));
            cells.Add("B7", "true");
            cells.Add("B8", "false");
            cells.Add("B9", true);
            cells.Add("B10", false);
            cells.Add("B11", "XYZ");
            cells.Add("C1", "2");
            cells.Add("C2", new TimeSpan(12, 14, 16));
            cells.Add("C3", new Cell("=A1", Cell.CellType.FORMULA, "C3"));
            Dictionary<string, object> expectedCells = new Dictionary<string, object>();
            expectedCells.Add("A1", 22);
            expectedCells.Add("A2", "21");
            expectedCells.Add("A3", true);
            expectedCells.Add("B1", 23);
            expectedCells.Add("B2", 20.1f);
            expectedCells.Add("B3", Utils.GetOATime(time));
            expectedCells.Add("B4", Utils.GetOADateTime(date));
            expectedCells.Add("B5", null);
            expectedCells.Add("B6", new Cell("=A1", Cell.CellType.FORMULA, "B6"));
            expectedCells.Add("B7", 1);
            expectedCells.Add("B8", 0);
            expectedCells.Add("B9", 1);
            expectedCells.Add("B10", 0);
            expectedCells.Add("B11", "XYZ");
            expectedCells.Add("C1", "2");
            expectedCells.Add("C2", new TimeSpan(12, 14, 16));
            expectedCells.Add("C3", new Cell("=A1", Cell.CellType.FORMULA, "C3"));
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

        [Theory(DisplayName = "Test of the import options for the import column types Numeric, Decimal and Double on parsed dates and times")]
        [InlineData(ImportOptions.ColumnType.Double, "2021-10-31 12:11:10", 44500.5077546296d)]
        [InlineData(ImportOptions.ColumnType.Double, "18:20:22", 0.764143518518519d)]
        [InlineData(ImportOptions.ColumnType.Decimal, "2021-10-31 12:11:10", "44500.5077546296")]
        [InlineData(ImportOptions.ColumnType.Decimal, "18:20:22", "0.764143518518519")]
        [InlineData(ImportOptions.ColumnType.Numeric, "2021-10-31 12:11:10", 44500.5077546296d)]
        [InlineData(ImportOptions.ColumnType.Numeric, "18:20:22", 0.764143518518519d)]
        public void EnforcingColumnAsNumberTest3(ImportOptions.ColumnType columnType, string givenValue, object expectedValue)
        {
            Dictionary<string, Object> cells = new Dictionary<string, object>();
            cells.Add("A1", true);
            cells.Add("B1", givenValue);
            cells.Add("C1", "2");

            Dictionary<string, object> expectedCells = new Dictionary<string, object>();
            expectedCells.Add("A1", true);
            if (columnType == ImportOptions.ColumnType.Decimal)
            {
                // m-suffix is not working
                expectedCells.Add("B1", Convert.ToDecimal(expectedValue));
            }
            else
            {
                expectedCells.Add("B1", expectedValue);
            }
            expectedCells.Add("C1", "2");
            ImportOptions options = new ImportOptions();
            options.EnforceDateTimesAsNumbers = true;
            options.AddEnforcedColumn(1, columnType);
            AssertValues<object, object>(cells, options, AssertApproximate, expectedCells);
        }

        [Theory(DisplayName = "Test of the import options for the import column type with wrong style information: Double and Decimal")]
        [InlineData("B", ImportOptions.ColumnType.Double)]
        [InlineData(1, ImportOptions.ColumnType.Double)]
        [InlineData("B", ImportOptions.ColumnType.Decimal)]
        [InlineData(1, ImportOptions.ColumnType.Decimal)]
        public void EnforcingColumnAsNumberTest4a(object column, ImportOptions.ColumnType type)
        {
            object ob1, ob2, ob4, ob5;
            object exOb1, exOb2, exOb4, exOb5;
            string ob3 = "5-7";
            string ob6 = "1870-06-01 12:12:00";
            if (type == ImportOptions.ColumnType.Double)
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
            Cell a1 = new Cell(1, Cell.CellType.NUMBER, "A1");
            Cell b1 = new Cell(ob1, Cell.CellType.NUMBER, "B1");
            b1.SetStyle(BasicStyles.DateFormat);
            Cell b2 = new Cell(ob2, Cell.CellType.NUMBER, "B2");
            b1.SetStyle(BasicStyles.TimeFormat);
            Cell b3 = new Cell(ob3, Cell.CellType.STRING, "B3");
            b1.SetStyle(BasicStyles.DateFormat);
            Cell b4 = new Cell(ob4, Cell.CellType.STRING, "B4");
            b1.SetStyle(BasicStyles.DateFormat);
            Cell b5 = new Cell(ob5, Cell.CellType.NUMBER, "B5");
            b1.SetStyle(BasicStyles.DateFormat);
            Cell b6 = new Cell(ob6, Cell.CellType.STRING, "B6");
            b5.SetStyle(BasicStyles.DateFormat);
            Cell c1 = new Cell(10, Cell.CellType.NUMBER, "C1");
            cells.Add("A1", a1);
            cells.Add("B1", b1);
            cells.Add("B2", b2);
            cells.Add("B3", b3);
            cells.Add("B4", b4);
            cells.Add("B5", b5);
            cells.Add("B6", b6);
            cells.Add("C1", c1);
            Dictionary<string, Cell> expectedCells = new Dictionary<string, Cell>();
            Cell exA1 = new Cell(1, Cell.CellType.NUMBER, "A1");
            Cell exB1 = new Cell(exOb1, Cell.CellType.NUMBER, "B1");
            Cell exB2 = new Cell(exOb2, Cell.CellType.NUMBER, "B2");
            Cell exB3 = new Cell(ob3, Cell.CellType.STRING, "B3");
            Cell exB4 = new Cell(exOb4, Cell.CellType.STRING, "B4");
            Cell exB5 = new Cell(exOb5, Cell.CellType.NUMBER, "B5");
            Cell exB6 = new Cell(ob6, Cell.CellType.STRING, "B6");
            Cell exC1 = new Cell(10, Cell.CellType.NUMBER, "C1");
            expectedCells.Add("A1", exA1);
            expectedCells.Add("B1", exB1);
            expectedCells.Add("B2", exB2);
            expectedCells.Add("B3", exB3);
            expectedCells.Add("B4", exB4);
            expectedCells.Add("B5", exB5);
            expectedCells.Add("B6", exB6);
            expectedCells.Add("C1", exC1);
            ImportOptions options = new ImportOptions();
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
            cells.Add("B10", null);
            cells.Add("B11", new Cell("=A1", Cell.CellType.FORMULA, "B11"));
            cells.Add("B12", 2);
            cells.Add("B13", "0");
            cells.Add("B14", "");
            cells.Add("C1", "0");
            cells.Add("C2", new TimeSpan(12, 14, 16));
            cells.Add("C3", new Cell("=A1", Cell.CellType.FORMULA, "C3"));
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
            expectedCells.Add("B10", null);
            expectedCells.Add("B11", new Cell("=A1", Cell.CellType.FORMULA, "B11"));
            expectedCells.Add("B12", 2);
            expectedCells.Add("B13", false);
            expectedCells.Add("B14", "");
            expectedCells.Add("C1", "0");
            expectedCells.Add("C2", new TimeSpan(12, 14, 16));
            expectedCells.Add("C3", new Cell("=A1", Cell.CellType.FORMULA, "C3"));
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
            cells.Add("B10", null);
            cells.Add("B11", new Cell("=A1", Cell.CellType.FORMULA, "B11"));
            cells.Add("B12", 2147483650);
            cells.Add("B13", 9223372036854775806);
            cells.Add("B14", 18446744073709551614);
            cells.Add("B15", (short)32766);
            cells.Add("B16", (ushort)65534);
            cells.Add("B17", 0.000000001d);
            cells.Add("B18", 0.123f);
            cells.Add("B19", (byte)17);
            cells.Add("C1", "0");
            cells.Add("C2", new TimeSpan(12, 14, 16));
            cells.Add("C3", new Cell("=A1", Cell.CellType.FORMULA, "C3"));
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
            expectedCells.Add("B10", null);
            expectedCells.Add("B11", "=A1");
            expectedCells.Add("B12", "2147483650");
            expectedCells.Add("B13", "9223372036854775806");
            expectedCells.Add("B14", "18446744073709551614");
            expectedCells.Add("B15", "32766");
            expectedCells.Add("B16", "65534");
            expectedCells.Add("B17", "1E-09"); // Currently handled without option to format the number
            expectedCells.Add("B18", "0.123");
            expectedCells.Add("B19", "17");
            expectedCells.Add("C1", "0");
            expectedCells.Add("C2", new TimeSpan(12, 14, 16));
            expectedCells.Add("C3", new Cell("=A1", Cell.CellType.FORMULA, "C3"));
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

        [Theory(DisplayName = "Test of the import options for the import column type: Date")]
        [InlineData("B")]
        [InlineData(1)]
        public void EnforcingColumnAsDateTest(object column)
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
            cells.Add("B6", 44494.5209490741d);
            cells.Add("B7", "2021-10-25 12:30:10");
            cells.Add("B8", -10);
            cells.Add("B9", 44494.5f);
            cells.Add("B10", null);
            cells.Add("B11", new Cell("=A1", Cell.CellType.FORMULA, "B11"));
            cells.Add("B12", 2147483650);
            cells.Add("B13", 2958466);
            cells.Add("C1", "0");
            cells.Add("C2", new TimeSpan(12, 14, 16));
            cells.Add("C3", new Cell("=A1", Cell.CellType.FORMULA, "C3"));
            Dictionary<string, object> expectedCells = new Dictionary<string, object>();
            expectedCells.Add("A1", 1);
            expectedCells.Add("A2", "21");
            expectedCells.Add("A3", true);
            expectedCells.Add("B1", new DateTime(1900, 1, 1, 0, 0, 0, 0));
            expectedCells.Add("B2", "Test");
            expectedCells.Add("B3", false);
            expectedCells.Add("B4", new DateTime(1900, 1, 1, 11, 12, 13, 0));
            expectedCells.Add("B5", new DateTime(2021, 8, 14, 18, 22, 13, 0));
            expectedCells.Add("B6", new DateTime(2021, 10, 25, 12, 30, 10, 0));
            expectedCells.Add("B7", new DateTime(2021, 10, 25, 12, 30, 10, 0));
            expectedCells.Add("B8", -10);
            expectedCells.Add("B9", new DateTime(2021, 10, 25, 12, 0, 0, 0));
            expectedCells.Add("B10", null);
            expectedCells.Add("B11", new Cell("=A1", Cell.CellType.FORMULA, "B11"));
            expectedCells.Add("B12", 2147483650);
            expectedCells.Add("B13", 2958466); // Exceeds year 9999
            expectedCells.Add("C1", "0");
            expectedCells.Add("C2", new TimeSpan(12, 14, 16));
            expectedCells.Add("C3", new Cell("=A1", Cell.CellType.FORMULA, "C3"));
            ImportOptions options = new ImportOptions();
            if (column is string)
            {
                options.AddEnforcedColumn(column as string, ImportOptions.ColumnType.Date);
            }
            else
            {
                options.AddEnforcedColumn((int)column, ImportOptions.ColumnType.Date);
            }
            AssertValues<object, object>(cells, options, AssertApproximate, expectedCells);
        }

        [Theory(DisplayName = "Test of the import options for the import column type (without casting to numbers) with missing formats for DateTime and TimeSpan")]
        [InlineData("B", "C")]
        [InlineData(1, 2)]
        void enforcingColumnAsDateTest2(object column1, object column2)
        {

            Dictionary<string, object> cells = new Dictionary<string, object>();
            cells.Add("A1", 1);
            cells.Add("B1", "11:12:13");
            cells.Add("C1", "2021-08-14 18:22:13");
            cells.Add("D1", "0");

            Dictionary<string, object> expectedCells = new Dictionary<string, object>();
            expectedCells.Add("A1", 1);
            expectedCells.Add("B1", new TimeSpan(11, 12, 13));
            expectedCells.Add("C1", new DateTime(2021, 8, 14, 18, 22, 13));
            expectedCells.Add("D1", "0");
            ImportOptions options = new ImportOptions();
            options.DateTimeFormat = null;
            options.TimeSpanFormat = null;
            if (column1 is String)
            {
                options.AddEnforcedColumn(column1 as string, ImportOptions.ColumnType.Time);
                options.AddEnforcedColumn(column2 as string, ImportOptions.ColumnType.Date);
            }
            else
            {
                options.AddEnforcedColumn((int)column1, ImportOptions.ColumnType.Time);
                options.AddEnforcedColumn((int)column2, ImportOptions.ColumnType.Date);
            }
            AssertValues<object, object>(cells, options, AssertApproximate, expectedCells);
        }

        [Theory(DisplayName = "Test of the import options for the import column type with wrong style information: Date")]
        [InlineData("B", ImportOptions.ColumnType.Date)]
        [InlineData(1, ImportOptions.ColumnType.Date)]
        [InlineData("B", ImportOptions.ColumnType.Time)]
        [InlineData(1, ImportOptions.ColumnType.Time)]
        public void EnforcingColumnAsDateTest2(object column, ImportOptions.ColumnType type)
        {
            Dictionary<string, Object> cells = new Dictionary<string, object>();
            Cell a1 = new Cell(1, Cell.CellType.NUMBER, "A1");
            Cell b1 = new Cell(-10, Cell.CellType.NUMBER, "B1");
            b1.SetStyle(BasicStyles.DateFormat);
            Cell b2 = new Cell(-5.5f, Cell.CellType.NUMBER, "B2");
            b2.SetStyle(BasicStyles.TimeFormat);
            Cell b3 = new Cell("5-7", Cell.CellType.STRING, "B3");
            b3.SetStyle(BasicStyles.DateFormat);
            Cell b4 = new Cell("-1", Cell.CellType.STRING, "B4");
            b4.SetStyle(BasicStyles.TimeFormat);
            Cell b5 = new Cell("1870-06-06 12:12:00", Cell.CellType.STRING, "B5");
            b5.SetStyle(BasicStyles.DateFormat);
            Cell c1 = new Cell(10, Cell.CellType.NUMBER, "C1");
            cells.Add("A1", a1);
            cells.Add("B1", b1);
            cells.Add("B2", b2);
            cells.Add("B3", b3);
            cells.Add("B4", b4);
            cells.Add("B5", b5);
            cells.Add("C1", c1);
            Dictionary<string, Cell> expectedCells = new Dictionary<string, Cell>();
            Cell exA1 = new Cell(1, Cell.CellType.NUMBER, "A1");
            Cell exB1 = new Cell(-10, Cell.CellType.NUMBER, "B1");
            Cell exB2 = new Cell(-5.5f, Cell.CellType.NUMBER, "B2");
            Cell exB3 = new Cell("5-7", Cell.CellType.STRING, "B3");
            Cell exB4 = new Cell("-1", Cell.CellType.STRING, "B4");
            Cell exB5 = new Cell("1870-06-06 12:12:00", Cell.CellType.STRING, "B5");
            Cell exC1 = new Cell(10, Cell.CellType.NUMBER, "C1");
            expectedCells.Add("A1", exA1);
            expectedCells.Add("B1", exB1);
            expectedCells.Add("B2", exB2);
            expectedCells.Add("B3", exB3);
            expectedCells.Add("B4", exB4);
            expectedCells.Add("B5", exB5);
            expectedCells.Add("C1", exC1);
            ImportOptions options = new ImportOptions();
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

        [Theory(DisplayName = "Test of the import options for the import column type: Time")]
        [InlineData("B")]
        [InlineData(1)]
        public void EnforcingColumnAsTimeTest(object column)
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
            cells.Add("B6", 44494.5209490741d);
            cells.Add("B7", "2021-10-25 12:30:10");
            cells.Add("B8", -10);
            cells.Add("B9", 44494.5f);
            cells.Add("B10", null);
            cells.Add("B11", new Cell("=A1", Cell.CellType.FORMULA, "B11"));
            cells.Add("C1", "0");
            cells.Add("C2", new TimeSpan(12, 14, 16));
            cells.Add("C3", new Cell("=A1", Cell.CellType.FORMULA, "C3"));
            Dictionary<string, object> expectedCells = new Dictionary<string, object>();
            expectedCells.Add("A1", 1);
            expectedCells.Add("A2", "21");
            expectedCells.Add("A3", true);
            expectedCells.Add("B1", new TimeSpan(1, 0, 0, 0));
            expectedCells.Add("B2", "Test");
            expectedCells.Add("B3", false);
            expectedCells.Add("B4", time);
            expectedCells.Add("B5", new TimeSpan(44422, 18, 22, 13));
            expectedCells.Add("B6", new TimeSpan(44494, 12, 30, 10));
            expectedCells.Add("B7", new TimeSpan(44494, 12, 30, 10));
            expectedCells.Add("B8", -10);
            expectedCells.Add("B9", new TimeSpan(44494, 12, 0, 0));
            expectedCells.Add("B10", null);
            expectedCells.Add("B11", new Cell("=A1", Cell.CellType.FORMULA, "B11"));
            expectedCells.Add("C1", "0");
            expectedCells.Add("C2", new TimeSpan(12, 14, 16));
            expectedCells.Add("C3", new Cell("=A1", Cell.CellType.FORMULA, "C3"));
            ImportOptions options = new ImportOptions();
            if (column is string)
            {
                options.AddEnforcedColumn(column as string, ImportOptions.ColumnType.Time);
            }
            else
            {
                options.AddEnforcedColumn((int)column, ImportOptions.ColumnType.Time);
            }
            AssertValues<object, object>(cells, options, AssertApproximate, expectedCells);
        }

        [Theory(DisplayName = "Test of the import options for the combination of a start row and a enforced column")]
        [InlineData(ImportOptions.ColumnType.Bool, "1", true)]
        [InlineData(ImportOptions.ColumnType.Bool, false, false)]
        [InlineData(ImportOptions.ColumnType.Double, "-2.5", -2.5d)]
        [InlineData(ImportOptions.ColumnType.Double, 13, 13d)]
        [InlineData(ImportOptions.ColumnType.Numeric, "12.5", 12.5f)]
        [InlineData(ImportOptions.ColumnType.Numeric, 13, 13)]
        [InlineData(ImportOptions.ColumnType.String, 16.5f, "16.5")]
        [InlineData(ImportOptions.ColumnType.String, true, "True")]
        public void EnforcingColumnStartRowTest(ImportOptions.ColumnType columnType, object givenValue, object expectedValue)
        {
            TimeSpan time = new TimeSpan(11, 12, 13);
            Dictionary<string, Object> cells = new Dictionary<string, object>();
            cells.Add("A1", "test");
            cells.Add("A2", 23);
            cells.Add("A3", time);
            cells.Add("B1", null);
            cells.Add("B2", givenValue);
            cells.Add("B3", givenValue);
            cells.Add("C1", 28);
            cells.Add("C2", false);
            cells.Add("C3", "Test");
            Dictionary<string, object> expectedCells = new Dictionary<string, object>();
            expectedCells.Add("A1", "test");
            expectedCells.Add("A2", 23);
            expectedCells.Add("A3", time);
            expectedCells.Add("B1", null);
            expectedCells.Add("B2", givenValue);
            expectedCells.Add("B3", expectedValue);
            expectedCells.Add("C1", 28);
            expectedCells.Add("C2", false);
            expectedCells.Add("C3", "Test");
            ImportOptions options = new ImportOptions();
            options.AddEnforcedColumn(1, columnType);
            options.EnforcingStartRowNumber = 2;
            AssertValues<object, object>(cells, options, AssertApproximate, expectedCells);
            ImportOptions options2 = new ImportOptions();
            options2.AddEnforcedColumn("B", columnType);
            options2.EnforcingStartRowNumber = 2;
            AssertValues<object, object>(cells, options2, AssertApproximate, expectedCells);
        }

        [Theory(DisplayName = "Test of the import options for the combination of a start row and a enforced column on types Date and Time")]
        [InlineData(ImportOptions.ColumnType.Date)]
        [InlineData(ImportOptions.ColumnType.Time)]
        public void EnforcingColumnStartRowTest2(ImportOptions.ColumnType columnType)
        {
            TimeSpan time = new TimeSpan(11, 12, 13);
            TimeSpan expectedTime = new TimeSpan(12, 13, 14);
            DateTime expectedDate = new DateTime(2021, 8, 14, 18, 22, 13, 0);

            Dictionary<string, Object> cells = new Dictionary<string, object>();
            cells.Add("A1", "test");
            cells.Add("A2", 23);
            cells.Add("A3", time);
            cells.Add("B1", null);
            if (columnType == ImportOptions.ColumnType.Time)
            {
                cells.Add("B2", "12:13:14");
                cells.Add("B3", "12:13:14");
            }
            else if (columnType == ImportOptions.ColumnType.Date)
            {
                cells.Add("B2", "2021-08-14 18:22:13");
                cells.Add("B3", "2021-08-14 18:22:13");
            }
            cells.Add("C1", 28);
            cells.Add("C2", false);
            cells.Add("C3", "Test");
            Dictionary<string, object> expectedCells = new Dictionary<string, object>();
            expectedCells.Add("A1", "test");
            expectedCells.Add("A2", 23);
            expectedCells.Add("A3", time);
            expectedCells.Add("B1", null);
            if (columnType == ImportOptions.ColumnType.Time)
            {
                expectedCells.Add("B2", "12:13:14");
                expectedCells.Add("B3", expectedTime);
            }
            else if (columnType == ImportOptions.ColumnType.Date)
            {
                expectedCells.Add("B2", "2021-08-14 18:22:13");
                expectedCells.Add("B3", expectedDate);
            }
            expectedCells.Add("C1", 28);
            expectedCells.Add("C2", false);
            expectedCells.Add("C3", "Test");
            ImportOptions options = new ImportOptions();
            options.AddEnforcedColumn(1, columnType);
            options.EnforcingStartRowNumber = 2;
            AssertValues<object, object>(cells, options, AssertApproximate, expectedCells);
            ImportOptions options2 = new ImportOptions();
            options2.AddEnforcedColumn("B", columnType);
            options2.EnforcingStartRowNumber = 2;
            AssertValues<object, object>(cells, options2, AssertApproximate, expectedCells);
        }

        [Theory(DisplayName = "Test of enforced import types when the same type overlaps globally and on a column")]
        [InlineData(ImportOptions.GlobalType.AllNumbersToDecimal, ImportOptions.ColumnType.Decimal, "7", "23", "1.1", typeof(decimal))]
        [InlineData(ImportOptions.GlobalType.AllNumbersToDouble, ImportOptions.ColumnType.Double, "7", "23", "1.1", typeof(double))]
        [InlineData(ImportOptions.GlobalType.AllNumbersToInt, ImportOptions.ColumnType.Numeric, "7", "23", "1.1", typeof(int))]
        [InlineData(ImportOptions.GlobalType.EverythingToString, ImportOptions.ColumnType.String, "7", "23", "1.1", typeof(string))]
        public void ImportEnforceOverlappingTest(ImportOptions.GlobalType globalType, ImportOptions.ColumnType columnType, string givenA2Value, string givenB1Value, string givenB2Value, Type expectedType)
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
            ImportOptions importOptions = new ImportOptions();
            importOptions.AddEnforcedColumn(1, columnType);
            importOptions.GlobalEnforcingType = globalType;
            AssertValues<object, object>(cells, importOptions, AssertApproximate, expectedCells);
        }


        [Theory(DisplayName = "Test of enforced import types when the global type overrules the column type")]
        [InlineData(ImportOptions.ColumnType.Decimal, ImportOptions.GlobalType.AllNumbersToDouble, typeof(decimal), "7", typeof(double), "7")]
        [InlineData(ImportOptions.ColumnType.Double, ImportOptions.GlobalType.AllNumbersToDecimal, typeof(double), "7", typeof(decimal), "7")]
        public void ImportEnforceOverruleTest(ImportOptions.ColumnType columnType, ImportOptions.GlobalType globalType, Type givenType, string givenValue, Type expectedType, string expectedValue)
        {
            Dictionary<string, object> cells = new Dictionary<string, object>();
            Dictionary<string, object> expectedCells = new Dictionary<string, object>();
            cells.Add("A1", TestUtils.CreateInstance(givenType, givenValue));
            expectedCells.Add("A1", TestUtils.CreateInstance(expectedType, expectedValue));
            ImportOptions importOptions = new ImportOptions();
            importOptions.AddEnforcedColumn(1, columnType);
            importOptions.GlobalEnforcingType = globalType;
            AssertValues<object, object>(cells, importOptions, AssertApproximate, expectedCells);
        }


        [Theory(DisplayName = "Test of the import options for custom date and time formats and culture info")]
        [InlineData(ImportOptions.ColumnType.Date, "en-US", "yyyy-MM-dd HH:mm:ss", "2021-08-12 12:11:10", "2021-08-12 12:11:10")]
        [InlineData(ImportOptions.ColumnType.Date, "de-DE", "dd.MM.yyyy HH:mm:ss", "12.08.2021 12:11:10", "2021-08-12 12:11:10")]
        [InlineData(ImportOptions.ColumnType.Date, "fr-FR", "dd/MM/yyyy", "12/08/2021", "2021-08-12 00:00:00")]
        [InlineData(ImportOptions.ColumnType.Date, null, null, "12.08.2021 12:11:10", "2021-12-08 12:11:10")]
        [InlineData(ImportOptions.ColumnType.Time, "en-US", "hh\\:mm\\:ss", "18:11:10", "18:11:10")]
        [InlineData(ImportOptions.ColumnType.Time, "", "hh", "12", "12:00:00")]
        [InlineData(ImportOptions.ColumnType.Time, null, null, "18:11:10", "18:11:10")]
        public void ParseDateTimeTest(ImportOptions.ColumnType columnType, string cultureInfo, string pattern, string givenValue, string expectedValue)
        {
            Dictionary<string, Object> cells = new Dictionary<string, object>();
            Dictionary<string, object> expectedCells = new Dictionary<string, object>();
            ImportOptions importOptions = new ImportOptions();
            if (columnType == ImportOptions.ColumnType.Date)
            {
                DateTime expected = DateTime.ParseExact(expectedValue, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                expectedCells.Add("A1", expected);
                importOptions.DateTimeFormat = pattern;
                importOptions.AddEnforcedColumn(0, ImportOptions.ColumnType.Date);
            }
            else
            {
                TimeSpan expected = TimeSpan.ParseExact(expectedValue, "hh\\:mm\\:ss", CultureInfo.InvariantCulture);
                expectedCells.Add("A1", expected);
                importOptions.TimeSpanFormat = pattern;
                importOptions.AddEnforcedColumn(0, ImportOptions.ColumnType.Time);
            }
            if (cultureInfo != null)
            {
                CultureInfo givenCultureInfo = new CultureInfo(cultureInfo); // empty will lead to invariant
                importOptions.TemporalCultureInfo = givenCultureInfo;
            }
            cells.Add("A1", givenValue);
            AssertValues<object, object>(cells, importOptions, AssertApproximate, expectedCells);
        }

        [Theory(DisplayName = "Test of the failing casting on an invalid Date or TimeSpan value")]
        [InlineData(ImportOptions.ColumnType.Time, "55.81.202x")]
        [InlineData(ImportOptions.ColumnType.Date, "2022-18-22 15:6x:00")]
        [InlineData(ImportOptions.ColumnType.Time, "10000-01-01 00:00:00")]
        [InlineData(ImportOptions.ColumnType.Date, "10000-01-01 00:00:00")]
        [InlineData(ImportOptions.ColumnType.Time, "1800-01-01 00:00:00")]
        [InlineData(ImportOptions.ColumnType.Time, "-10:00:00")]
        void InvalidDateCastingTest(ImportOptions.ColumnType columnType, string value)
        {
            ImportOptions options = new ImportOptions();
            options.EnforceDateTimesAsNumbers = true;
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
            ImportOptions options = new ImportOptions();
            options.EnforceDateTimesAsNumbers = true;
            options.AddEnforcedColumn("A", ImportOptions.ColumnType.Time);
            options.TimeSpanFormat = "HH:mm:ss d";
            Dictionary<string, object> cells = new Dictionary<string, object>();
            Dictionary<string, object> expectedCells = new Dictionary<string, object>();
            cells.Add("A1", "00:00:00 2958467");
            expectedCells.Add("A1", "00:00:00 2958467");
            AssertValues<object, object>(cells, options, AssertApproximate, expectedCells);
        }

        [Theory(DisplayName = "Test of the import option to process or discard phonetic characters in strings")]
        [InlineData(0, 1, false)]
        [InlineData(0, 2, true)]
        void PhnoneticCharactersImportOptionTest(int givenValuesColumn, int expectedValuesColumn, bool importOptionValue)
        {
            // Note: Cells in column A contains the strings with phonetic characters.
            // The corresponding Cells in column B contains the values without enabled import option.
            // The cells in column C contains the values with the import option enabled.
            // The values starts at Row 2 (Index 1)

            ImportOptions options = new ImportOptions();
            options.EnforcePhoneticCharacterImport = importOptionValue;
            Stream stream = TestUtils.GetResource("phonetics.xlsx");
            Workbook workbook = Workbook.Load(stream, options);

            int lastRow = workbook.Worksheets[0].GetLastDataRowNumber();
            for (int r = 1; r <= lastRow; r++)
            {
                string given = workbook.Worksheets[0].GetCell(new Address(givenValuesColumn, r)).Value.ToString();
                string expected = workbook.Worksheets[0].GetCell(new Address(expectedValuesColumn, r)).Value.ToString();
                Assert.Equal(expected, given);
            }
        }

        [Theory(DisplayName = "Test of the ImportOption property EnforceValidColumnDimensions")]
        [InlineData("valid_column_row_dimensions.xlsx", true, false, 0)]
        [InlineData("invalid_column_width_min.xlsx", true, true, -1)]
        [InlineData("invalid_column_width_max.xlsx", true, true, 1)]
        [InlineData("invalid_row_height_min.xlsx", true, false, 0)]
        [InlineData("invalid_row_height_max.xlsx", true, false, 0)]
        [InlineData("valid_column_row_dimensions.xlsx", false, false, 0)]
        [InlineData("invalid_column_width_min.xlsx", false, false, 0)]
        [InlineData("invalid_column_width_max.xlsx", false, false, 0)]
        [InlineData("invalid_row_height_min.xlsx", false, false, 0)]
        [InlineData("invalid_row_height_max.xlsx", false, false, 0)]
        public void EnforceValidColumnDimensionsTest(string fileName, bool givenOptionValue, bool expectedThrow, int columnFlag)
        {
            ImportOptions options = new ImportOptions();
            options.EnforceValidColumnDimensions = givenOptionValue;
            options.EnforceValidRowDimensions = false;
            using Stream stream = TestUtils.GetResource(fileName);

            if (expectedThrow)
            {
                Assert.ThrowsAny<Exception>(() => Workbook.Load(stream, options));
            }
            else
            {
                Workbook workbook = Workbook.Load(stream, options);
                if (columnFlag == -1)
                {
                    Assert.Equal(Worksheet.MIN_COLUMN_WIDTH, workbook.GetWorksheet(0).Columns[0].Width);
                }
                else if (columnFlag == 1)
                {
                    Assert.Equal(Worksheet.MAX_COLUMN_WIDTH, workbook.GetWorksheet(0).Columns[0].Width);
                }
                else
                {
                    Assert.True(true);
                }
            }
        }

        [Theory(DisplayName = "Test of the ImportOption property EnforceValidRowDimensions")]
        [InlineData("valid_column_row_dimensions.xlsx", true, false, 0)]
        [InlineData("invalid_row_height_min.xlsx", true, true, -1)]
        [InlineData("invalid_row_height_max.xlsx", true, true, 1)]
        [InlineData("invalid_column_width_min.xlsx", true, false, 0)]
        [InlineData("invalid_column_width_max.xlsx", true, false, 0)]
        [InlineData("valid_column_row_dimensions.xlsx", false, false, 0)]
        [InlineData("invalid_row_height_min.xlsx", false, false, 0)]
        [InlineData("invalid_row_height_max.xlsx", false, false, 0)]
        [InlineData("invalid_column_width_min.xlsx", false, false, 0)]
        [InlineData("invalid_column_width_max.xlsx", false, false, 0)]
        public void EnforceValidRowDimensionsTest(string fileName, bool givenOptionValue, bool expectedThrow, int rowFlag)
        {
            ImportOptions options = new ImportOptions();
            options.EnforceValidRowDimensions = givenOptionValue;
            options.EnforceValidColumnDimensions = false;
            using Stream stream = TestUtils.GetResource(fileName);

            if (expectedThrow)
            {
                Assert.ThrowsAny<Exception>(() => Workbook.Load(stream, options));
            }
            else
            {
                Workbook workbook = Workbook.Load(stream, options);
                if (rowFlag == -1)
                {
                    Assert.Equal(Worksheet.MIN_ROW_HEIGHT, workbook.GetWorksheet(0).RowHeights[0]);
                }
                else if (rowFlag == 1)
                {
                    Assert.Equal(Worksheet.MAX_ROW_HEIGHT, workbook.GetWorksheet(0).RowHeights[0]);
                }
                else
                {
                    Assert.True(true);
                }
            }
        }



        private static void AssertValues<T, D>(Dictionary<string, T> givenCells, ImportOptions importOptions, Action<object, object> assertionAction, Dictionary<string, D> expectedCells = null)
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
                else if (expectedValue is Cell && (givenCell != null))
                {
                    assertionAction.Invoke((expectedValue as Cell).Value, (givenCell).Value);
                }
                else if (expectedValue is Cell)
                {
                    assertionAction.Invoke((D)(expectedValue as Cell).Value, (D)givenCell.Value);
                    Assert.Equal(Cell.CellType.FORMULA, givenCell.DataType);
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
                double e = Utils.GetOADateTime((DateTime)expected);
                double g = Utils.GetOADateTime((DateTime)given);
                AssertApproximate(e, g);
            }
            else if (given is TimeSpan)
            {
                double g = Utils.GetOATime((TimeSpan)given);
                double e = Utils.GetOATime((TimeSpan)expected);
                AssertApproximate(e, g);
            }
            else
            {
                AssertEquals<object>(expected, given);
            }

        }

    }
}
