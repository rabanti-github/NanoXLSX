using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using NanoXLSX.Exceptions;
using NanoXLSX.Styles;
using Xunit;

namespace NanoXLSX.Test.Core.WorksheetTest
{
    public class AddCellRangeTest
    {
        public enum RangeType
        {
            OneColumn,
            OneRow,
            ThreeColumnsFourRows,
            FourColumnsThreeRows
        }

        public enum TestType
        {
            RandomList,
            CellList
        }

        [Theory(DisplayName = "Test of the AddCellRange function for a random list or list of nested cells with start and end address")]
        [InlineData(0, 0, RangeType.OneColumn, TestType.RandomList)]
        [InlineData(7, 27, RangeType.OneRow, TestType.RandomList)]
        [InlineData(5, 13, RangeType.FourColumnsThreeRows, TestType.RandomList)]
        [InlineData(22, 11, RangeType.ThreeColumnsFourRows, TestType.RandomList)]
        [InlineData(0, 0, RangeType.OneColumn, TestType.CellList)]
        [InlineData(7, 27, RangeType.OneRow, TestType.CellList)]
        [InlineData(5, 13, RangeType.FourColumnsThreeRows, TestType.CellList)]
        [InlineData(22, 11, RangeType.ThreeColumnsFourRows, TestType.CellList)]
        public void AddCellRangeTest1(int startColumn, int startRow, RangeType type, TestType testType)
        {
            ListTuple data = GetList(startColumn, startRow, type, testType);

            Worksheet worksheet = new Worksheet();
            Address startAddress = new Address(startColumn, startRow);
            Address endAddress = ListTuple.GetEndAddress(startColumn, startRow, type);

            Assert.Empty(worksheet.Cells);
            worksheet.AddCellRange(data.Values, startAddress, endAddress);
            AssertRange(worksheet, data);
        }

        [Theory(DisplayName = "Test of the AddCellRange function for a random list or list of nested cells with start and end address and a style")]
        [InlineData(0, 0, RangeType.OneColumn, TestType.RandomList)]
        [InlineData(7, 27, RangeType.OneRow, TestType.RandomList)]
        [InlineData(5, 13, RangeType.FourColumnsThreeRows, TestType.RandomList)]
        [InlineData(22, 11, RangeType.ThreeColumnsFourRows, TestType.RandomList)]
        [InlineData(0, 0, RangeType.OneColumn, TestType.CellList)]
        [InlineData(7, 27, RangeType.OneRow, TestType.CellList)]
        [InlineData(5, 13, RangeType.FourColumnsThreeRows, TestType.CellList)]
        [InlineData(22, 11, RangeType.ThreeColumnsFourRows, TestType.CellList)]
        public void AddCellRangeTest2(int startColumn, int startRow, RangeType type, TestType testType)
        {
            ListTuple data = GetList(startColumn, startRow, type, testType);
            Worksheet worksheet = new Worksheet();
            Address startAddress = new Address(startColumn, startRow);
            Address endAddress = ListTuple.GetEndAddress(startColumn, startRow, type);

            Assert.Empty(worksheet.Cells);
            worksheet.AddCellRange(data.Values, startAddress, endAddress, BasicStyles.Bold);
            AssertRange(worksheet, data);
            AssertRangeStyle(worksheet, data, BasicStyles.Bold);
        }

        [Theory(DisplayName = "Test of the AddCellRange function for a random list or list of nested cells with start and end address and a active style on the workbook")]
        [InlineData(0, 0, RangeType.OneColumn, TestType.RandomList)]
        [InlineData(7, 27, RangeType.OneRow, TestType.RandomList)]
        [InlineData(5, 13, RangeType.FourColumnsThreeRows, TestType.RandomList)]
        [InlineData(22, 11, RangeType.ThreeColumnsFourRows, TestType.RandomList)]
        [InlineData(0, 0, RangeType.OneColumn, TestType.CellList)]
        [InlineData(7, 27, RangeType.OneRow, TestType.CellList)]
        [InlineData(5, 13, RangeType.FourColumnsThreeRows, TestType.CellList)]
        [InlineData(22, 11, RangeType.ThreeColumnsFourRows, TestType.CellList)]
        public void AddCellRangeTest3(int startColumn, int startRow, RangeType type, TestType testType)
        {
            ListTuple data = GetList(startColumn, startRow, type, testType);
            Worksheet worksheet = new Worksheet();
            worksheet.SetActiveStyle(BasicStyles.BorderFrame);
            Address startAddress = new Address(startColumn, startRow);
            Address endAddress = ListTuple.GetEndAddress(startColumn, startRow, type);

            Assert.Empty(worksheet.Cells);
            worksheet.AddCellRange(data.Values, startAddress, endAddress);
            AssertRange(worksheet, data);
            AssertRangeStyle(worksheet, data, BasicStyles.BorderFrame);
        }

        [Theory(DisplayName = "Test of the AddCellRange function for a random list or list of nested cells with a range as string")]
        [InlineData("A1:A12", RangeType.OneColumn, TestType.RandomList)]
        [InlineData("H28:S28", RangeType.OneRow, TestType.RandomList)]
        [InlineData("F14:I16", RangeType.FourColumnsThreeRows, TestType.RandomList)]
        [InlineData("T12:V15", RangeType.ThreeColumnsFourRows, TestType.RandomList)]
        [InlineData("A1:A12", RangeType.OneColumn, TestType.CellList)]
        [InlineData("H28:S28", RangeType.OneRow, TestType.CellList)]
        [InlineData("F14:I16", RangeType.FourColumnsThreeRows, TestType.CellList)]
        [InlineData("T12:V15", RangeType.ThreeColumnsFourRows, TestType.CellList)]
        public void AddCellRangeTest4(string range, RangeType type, TestType testType)
        {
            NanoXLSX.Range cellRange = Cell.ResolveCellRange(range);
            ListTuple data = GetList(cellRange.StartAddress.Column, cellRange.StartAddress.Row, type, testType);
            Worksheet worksheet = new Worksheet();

            Assert.Empty(worksheet.Cells);
            worksheet.AddCellRange(data.Values, range);
            AssertRange(worksheet, data);
        }

        [Theory(DisplayName = "Test of the AddCellRange function for a random list or list of nested cells with a range as range object")]
        [InlineData("A1:A12", RangeType.OneColumn, TestType.RandomList)]
        [InlineData("H28:S28", RangeType.OneRow, TestType.RandomList)]
        [InlineData("F14:I16", RangeType.FourColumnsThreeRows, TestType.RandomList)]
        [InlineData("T12:V15", RangeType.ThreeColumnsFourRows, TestType.RandomList)]
        [InlineData("A1:A12", RangeType.OneColumn, TestType.CellList)]
        [InlineData("H28:S28", RangeType.OneRow, TestType.CellList)]
        [InlineData("F14:I16", RangeType.FourColumnsThreeRows, TestType.CellList)]
        [InlineData("T12:V15", RangeType.ThreeColumnsFourRows, TestType.CellList)]
        public void AddCellRangeTest5(string range, RangeType type, TestType testType)
        {
            NanoXLSX.Range cellRange = Cell.ResolveCellRange(range);
            ListTuple data = GetList(cellRange.StartAddress.Column, cellRange.StartAddress.Row, type, testType);
            Worksheet worksheet = new Worksheet();

            Assert.Empty(worksheet.Cells);
            worksheet.AddCellRange(data.Values, cellRange);
            AssertRange(worksheet, data);
        }


        [Theory(DisplayName = "Test of the AddCellRange function for a random list or list of nested cells with a range as string and a style")]
        [InlineData("A1:A12", RangeType.OneColumn, TestType.RandomList)]
        [InlineData("H28:S28", RangeType.OneRow, TestType.RandomList)]
        [InlineData("F14:I16", RangeType.FourColumnsThreeRows, TestType.RandomList)]
        [InlineData("T12:V15", RangeType.ThreeColumnsFourRows, TestType.RandomList)]
        [InlineData("A1:A12", RangeType.OneColumn, TestType.CellList)]
        [InlineData("H28:S28", RangeType.OneRow, TestType.CellList)]
        [InlineData("F14:I16", RangeType.FourColumnsThreeRows, TestType.CellList)]
        [InlineData("T12:V15", RangeType.ThreeColumnsFourRows, TestType.CellList)]
        public void AddCellRangeTest6(string range, RangeType type, TestType testType)
        {
            NanoXLSX.Range cellRange = Cell.ResolveCellRange(range);
            ListTuple data = GetList(cellRange.StartAddress.Column, cellRange.StartAddress.Row, type, testType);
            Worksheet worksheet = new Worksheet();

            Assert.Empty(worksheet.Cells);
            worksheet.AddCellRange(data.Values, range, BasicStyles.BoldItalic);
            AssertRange(worksheet, data);
            AssertRangeStyle(worksheet, data, BasicStyles.BoldItalic);
        }

        [Theory(DisplayName = "Test of the AddCellRange function for a random list  or list of nested cells with a range as string and an active style on the worksheet")]
        [InlineData("A1:A12", RangeType.OneColumn, TestType.RandomList)]
        [InlineData("H28:S28", RangeType.OneRow, TestType.RandomList)]
        [InlineData("F14:I16", RangeType.FourColumnsThreeRows, TestType.RandomList)]
        [InlineData("T12:V15", RangeType.ThreeColumnsFourRows, TestType.RandomList)]
        [InlineData("A1:A12", RangeType.OneColumn, TestType.CellList)]
        [InlineData("H28:S28", RangeType.OneRow, TestType.CellList)]
        [InlineData("F14:I16", RangeType.FourColumnsThreeRows, TestType.CellList)]
        [InlineData("T12:V15", RangeType.ThreeColumnsFourRows, TestType.CellList)]
        public void AddCellRangeTest7(string range, RangeType type, TestType testType)
        {
            NanoXLSX.Range cellRange = Cell.ResolveCellRange(range);
            ListTuple data = GetList(cellRange.StartAddress.Column, cellRange.StartAddress.Row, type, testType);
            Worksheet worksheet = new Worksheet();
            worksheet.SetActiveStyle(BasicStyles.BorderFrameHeader);

            Assert.Empty(worksheet.Cells);
            worksheet.AddCellRange(data.Values, range, BasicStyles.BoldItalic);
            AssertRange(worksheet, data);
            AssertRangeStyle(worksheet, data, BasicStyles.BorderFrameHeader);
        }


        [Theory(DisplayName = "Test of the AddCellRange function for a random list  or list of nested cells with a range as range object and an active style on the worksheet")]
        [InlineData("A1:A12", RangeType.OneColumn, TestType.RandomList)]
        [InlineData("H28:S28", RangeType.OneRow, TestType.RandomList)]
        [InlineData("F14:I16", RangeType.FourColumnsThreeRows, TestType.RandomList)]
        [InlineData("T12:V15", RangeType.ThreeColumnsFourRows, TestType.RandomList)]
        [InlineData("A1:A12", RangeType.OneColumn, TestType.CellList)]
        [InlineData("H28:S28", RangeType.OneRow, TestType.CellList)]
        [InlineData("F14:I16", RangeType.FourColumnsThreeRows, TestType.CellList)]
        [InlineData("T12:V15", RangeType.ThreeColumnsFourRows, TestType.CellList)]
        public void AddCellRangeTest8(string range, RangeType type, TestType testType)
        {
            NanoXLSX.Range cellRange = Cell.ResolveCellRange(range);
            ListTuple data = GetList(cellRange.StartAddress.Column, cellRange.StartAddress.Row, type, testType);
            Worksheet worksheet = new Worksheet();
            worksheet.SetActiveStyle(BasicStyles.BorderFrameHeader);

            Assert.Empty(worksheet.Cells);
            worksheet.AddCellRange(data.Values, cellRange, BasicStyles.BoldItalic);
            AssertRange(worksheet, data);
            AssertRangeStyle(worksheet, data, BasicStyles.BorderFrameHeader);
        }

        [Theory(DisplayName = "Test of the failing AddCellRange function with a deviating row and column number")]
        [InlineData(0, 0, RangeType.OneColumn)]
        [InlineData(7, 27, RangeType.OneRow)]
        [InlineData(5, 13, RangeType.FourColumnsThreeRows)]
        [InlineData(22, 11, RangeType.ThreeColumnsFourRows)]
        public void AddCellRangeFailingTest(int startColumn, int startRow, RangeType type)
        {
            ListTuple data = GetRandomList(0, 0, type);
            Worksheet worksheet = new Worksheet();
            Address startAddress = new Address(startColumn, startRow);
            Address endAddress = ListTuple.GetEndAddress(startColumn + 1, startRow + 1, type);

            Assert.Empty(worksheet.Cells);
            Assert.Throws<RangeException>(() => worksheet.AddCellRange(data.Values, startAddress, endAddress));
        }

        [Theory(DisplayName = "Test of the failing AddCellRange function with a deviating range definition (string)")]
        [InlineData("A1:A12", "A1:A13", RangeType.OneColumn)]
        [InlineData("H28:S28", "H28:S29", RangeType.OneRow)]
        [InlineData("F14:I16", "F14:J16", RangeType.FourColumnsThreeRows)]
        [InlineData("T12:V15", "T12:W15", RangeType.ThreeColumnsFourRows)]
        public void AddCellRangeFailingTest2(string givenRange, string falseRange, RangeType type)
        {
            NanoXLSX.Range cellRange = Cell.ResolveCellRange(givenRange);
            ListTuple data = GetRandomList(cellRange.StartAddress.Column, cellRange.StartAddress.Row, type);
            Worksheet worksheet = new Worksheet();

            Assert.Empty(worksheet.Cells);
            Assert.Throws<RangeException>(() => worksheet.AddCellRange(data.Values, falseRange));
        }

        [Fact(DisplayName = "Test of the failing AddCellRange function with a null as passed list")]
        public void AddCellRangeFailingTest3()
        {
            NanoXLSX.Range cellRange = "A1:A1";
            Worksheet worksheet = new Worksheet();
            Assert.Empty(worksheet.Cells);
            Assert.Throws<RangeException>(() => worksheet.AddCellRange(null, cellRange));
        }

        private void AssertRange(Worksheet worksheet, ListTuple expectedData)
        {
            Assert.Equal(expectedData.Count, worksheet.Cells.Count);
            for (int i = 0; i < expectedData.Count; i++)
            {
                string expectedAddress = expectedData.Addresses[i].GetAddress();
                Assert.True(worksheet.Cells.ContainsKey(expectedAddress));
                Assert.Equal(expectedData.Values[i], worksheet.Cells[expectedAddress].Value);
                Assert.Equal(expectedData.Types[i], worksheet.Cells[expectedAddress].DataType);
            }
        }

        private void AssertRangeStyle(Worksheet worksheet, ListTuple expectedData, Style expectedSourceStyle)
        {
            for (int i = 0; i < expectedData.Count; i++)
            {
                string expectedAddress = expectedData.Addresses[i].GetAddress();
                if (expectedData.Styles[i] == null)
                {
                    Assert.True(expectedSourceStyle.Equals(worksheet.Cells[expectedAddress].CellStyle));
                }
                else
                {
                    Style mergedStyle = (Style)expectedSourceStyle.Copy();
                    mergedStyle.Append(expectedData.Styles[i]);
                    Assert.True(mergedStyle.Equals(worksheet.Cells[expectedAddress].CellStyle));
                }
            }
        }

        private static ListTuple GetList(int startColumn, int startRow, RangeType type, TestType testType)
        {
            ListTuple data;
            if (testType == TestType.RandomList)
            {
                data = GetRandomList(startColumn, startRow, type);
            }
            else
            {
                data = GetCellList(startColumn, startRow, type);
            }
            return data;
        }

        [ExcludeFromCodeCoverage]
        private static ListTuple GetRandomList(int startColumn, int startRow, RangeType type, bool addNull = true)
        {
            ListTuple list = new ListTuple(startColumn, startRow, type);
            // IMPORTANT: The list must contain 12 entries
            list.Add(22, Cell.CellType.Number);
            list.Add(-0.55f, Cell.CellType.Number);
            list.Add(22.22d, Cell.CellType.Number);
            list.Add(true, Cell.CellType.Bool);
            list.Add(false, Cell.CellType.Bool);
            list.Add("", Cell.CellType.String);
            list.Add("test", Cell.CellType.String);
            list.Add((byte)64, Cell.CellType.Number);
            list.Add(77777L, Cell.CellType.Number);
            list.Add(new DateTime(2020, 11, 1, 11, 22, 13, 99), Cell.CellType.Date);
            list.Add(new TimeSpan(13, 16, 22), Cell.CellType.Time);
            if (addNull)
            {
                list.Add(null, Cell.CellType.Empty);
            }
            else
            {
                list.Add("substitute", Cell.CellType.String);
            }
            return list;
        }

        [ExcludeFromCodeCoverage]
        private static ListTuple GetCellList(int startColumn, int startRow, RangeType type, bool addNull = true)
        {
            ListTuple list = new ListTuple(startColumn, startRow, type);
            // IMPORTANT: The list must contain 12 entries
            list.Add(new Cell(23, Cell.CellType.Default, "X1"), Cell.CellType.Number);
            list.Add(new Cell(-0.65f, Cell.CellType.Default, "X2"), Cell.CellType.Number);
            list.Add(new Cell(42.22d, Cell.CellType.Default, "X3"), Cell.CellType.Number);
            list.Add(new Cell(false, Cell.CellType.Default, "X4"), Cell.CellType.Bool);
            list.Add(new Cell(true, Cell.CellType.Default, "X5"), Cell.CellType.Bool);
            list.Add(new Cell("test2", Cell.CellType.Default, "X6"), Cell.CellType.String);
            list.Add(new Cell("", Cell.CellType.Default, "X7"), Cell.CellType.String);
            list.Add(new Cell((byte)64, Cell.CellType.Default, "X8"), Cell.CellType.Number);
            list.Add(new Cell(88888L, Cell.CellType.Default, "X9"), Cell.CellType.Number);
            list.Add(new Cell(new DateTime(2020, 10, 1, 11, 22, 13, 99), Cell.CellType.Default, "X10"), Cell.CellType.Date);
            list.Add(new Cell(new TimeSpan(13, 15, 22), Cell.CellType.Default, "X11"), Cell.CellType.Time);
            if (addNull)
            {
                list.Add(new Cell(null, Cell.CellType.Default, "X12"), Cell.CellType.Empty);
            }
            else
            {
                list.Add(new Cell("substitute2", Cell.CellType.Default, "X12"), Cell.CellType.String);
            }
            return list;
        }

        public class ListTuple
        {
            private readonly List<Address> preparedAddresses;
            private int currentIndex = 0;

            public List<object> Values { get; private set; }
            public List<Cell.CellType> Types { get; private set; }
            public List<Address> Addresses { get; private set; }
            public List<Style> Styles { get; set; }
            public int Count { get; private set; }

            public ListTuple(int startColumn, int startRow, RangeType rangeType)
            {
                Values = new List<object>();
                Types = new List<Cell.CellType>();
                Addresses = new List<Address>();
                Styles = new List<Style>();

                Count = 12;
                Address endAddress = GetEndAddress(startColumn, startRow, rangeType);
                preparedAddresses = Cell.GetCellRange(startColumn, startRow, endAddress.Column, endAddress.Row).ToList();

            }

            public void Add(object value, Cell.CellType type)
            {
                if (value is Cell cell)
                {
                    Values.Add(cell.Value);
                }
                else
                {
                    Values.Add(value);
                }
                Types.Add(type);
                Addresses.Add(preparedAddresses[currentIndex]);
                if (type.Equals(Cell.CellType.Date))
                {
                    Styles.Add(BasicStyles.DateFormat);
                }
                else if (type.Equals(Cell.CellType.Time))
                {
                    Styles.Add(BasicStyles.TimeFormat);
                }
                else
                {
                    Styles.Add(null);
                }
                currentIndex++;
            }

            public static Address GetEndAddress(int startColumn, int startRow, RangeType rangeType)
            {
                switch (rangeType)
                {
                    case RangeType.OneColumn:
                        return new Address(startColumn, startRow + 11);
                    case RangeType.OneRow:
                        return new Address(startColumn + 11, startRow);
                    case RangeType.ThreeColumnsFourRows:
                        return new Address(startColumn + 2, startRow + 3);
                    default:
                        return new Address(startColumn + 3, startRow + 2);
                }
            }

        }


    }
}
