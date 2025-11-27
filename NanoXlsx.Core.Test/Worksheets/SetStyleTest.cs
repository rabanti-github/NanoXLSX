using System;
using System.Collections.Generic;
using System.Linq;
using NanoXLSX.Styles;
using NanoXLSX.Test.Core.Utils;
using Xunit;
using FormatException = NanoXLSX.Exceptions.FormatException;

namespace NanoXLSX.Test.Core.WorksheetTest
{
    public class SetStyleTest
    {

        public enum RangeRepresentation
        {
            StringExpression,
            RangeObject
        }

        [Theory(DisplayName = "Test of the SetStyle function on an empty worksheet with a Range object or its string representation")]
        [InlineData("A1:A1", RangeRepresentation.RangeObject)]
        [InlineData("A1:A5", RangeRepresentation.RangeObject)]
        [InlineData("A1:C1", RangeRepresentation.RangeObject)]
        [InlineData("A1:C3", RangeRepresentation.RangeObject)]
        [InlineData("R17:N22", RangeRepresentation.RangeObject)]
        [InlineData("A1:A1", RangeRepresentation.StringExpression)]
        [InlineData("A1:A5", RangeRepresentation.StringExpression)]
        [InlineData("A1:C1", RangeRepresentation.StringExpression)]
        [InlineData("A1:C3", RangeRepresentation.StringExpression)]
        [InlineData("R17:N22", RangeRepresentation.StringExpression)]
        public void SetStyleTest1(string rangeString, RangeRepresentation representation)
        {
            Worksheet worksheet = new Worksheet();
            Assert.Empty(worksheet.Cells);
            Range range = new Range(rangeString);
            if (representation == RangeRepresentation.RangeObject)
            {
                worksheet.SetStyle(range, BasicStyles.BoldItalic);
            }
            else
            {
                worksheet.SetStyle(rangeString, BasicStyles.BoldItalic);
            }
            List<string> emptyCells = (from i in range.ResolveEnclosedAddresses() select i.GetAddress()).ToList();
            AssertCellRange(rangeString, BasicStyles.BoldItalic, worksheet, emptyCells, emptyCells.Count);
        }

        [Theory(DisplayName = "Test of the SetStyle function on an empty worksheet with a Range object or its string representation with null as style")]
        [InlineData("A1:A1", RangeRepresentation.RangeObject)]
        [InlineData("A1:A5", RangeRepresentation.RangeObject)]
        [InlineData("A1:C1", RangeRepresentation.RangeObject)]
        [InlineData("A1:C3", RangeRepresentation.RangeObject)]
        [InlineData("R17:N22", RangeRepresentation.RangeObject)]
        [InlineData("A1:A1", RangeRepresentation.StringExpression)]
        [InlineData("A1:A5", RangeRepresentation.StringExpression)]
        [InlineData("A1:C1", RangeRepresentation.StringExpression)]
        [InlineData("A1:C3", RangeRepresentation.StringExpression)]
        [InlineData("R17:N22", RangeRepresentation.StringExpression)]
        public void SetStyleTest1b(string rangeString, RangeRepresentation representation)
        {
            Worksheet worksheet = new Worksheet();
            Assert.Empty(worksheet.Cells);
            Range range = new Range(rangeString);
            if (representation == RangeRepresentation.RangeObject)
            {
                worksheet.SetStyle(range, null);
            }
            else
            {
                worksheet.SetStyle(rangeString, null);
            }
            Assert.Empty(worksheet.Cells); // Should not create empty cells
        }

        [Theory(DisplayName = "Test of the SetStyle function on a worksheet with existing cells and a Range object or its string representation")]
        [InlineData("A1:A1", "A1", 22, null, RangeRepresentation.RangeObject)]
        [InlineData("A1:A5", "A2", true, "A1,A3,A4,A5", RangeRepresentation.RangeObject)]
        [InlineData("A1:C1", "B1", "test", "A1,C1", RangeRepresentation.RangeObject)]
        [InlineData("A1:C3", "B2", -0.25f, "A1,A2,A3,B1,B3,C1,C2,C3", RangeRepresentation.RangeObject)]
        [InlineData("R17:T21", "R18,R19,R20,S19", 99999L, "R17,R21,S17,S18,S20,S21,T17,T18,T19,T20,T21", RangeRepresentation.RangeObject)]
        [InlineData("A1:A1", "A1", 22, null, RangeRepresentation.StringExpression)]
        [InlineData("A1:A5", "A2", true, "A1,A3,A4,A5", RangeRepresentation.StringExpression)]
        [InlineData("A1:C1", "B1", "test", "A1,C1", RangeRepresentation.StringExpression)]
        [InlineData("A1:C3", "B2", -0.25f, "A1,A2,A3,B1,B3,C1,C2,C3", RangeRepresentation.StringExpression)]
        [InlineData("R17:T21", "R18,R19,R20,S19", 99999L, "R17,R21,S17,S18,S20,S21,T17,T18,T19,T20,T21", RangeRepresentation.StringExpression)]
        public void SetStyleTest2(string rangeString, string definedCells, object cellValue, string expectedEmptyCells, RangeRepresentation representation)
        {
            Worksheet worksheet = new Worksheet();
            AddCells(worksheet, cellValue, definedCells);
            int cellCount = worksheet.Cells.Count;
            Assert.NotEqual(0, cellCount);
            Range range = new Range(rangeString);
            if (representation == RangeRepresentation.RangeObject)
            {
                worksheet.SetStyle(range, BasicStyles.Bold);
            }
            else
            {
                worksheet.SetStyle(rangeString, BasicStyles.Bold);
            }
            List<string> emptyCells = TestUtils.SplitValuesAsList(expectedEmptyCells);
            AssertCellRange(rangeString, BasicStyles.Bold, worksheet, emptyCells, emptyCells.Count + cellCount);
        }

        [Theory(DisplayName = "Test of the SetStyle function on a worksheet with existing cells and a Range object or its string representation with null as style")]
        [InlineData("A1:A1", "A1", 22, RangeRepresentation.RangeObject)]
        [InlineData("A1:A5", "A2", true, RangeRepresentation.RangeObject)]
        [InlineData("A1:C1", "B1", "test", RangeRepresentation.RangeObject)]
        [InlineData("A1:C3", "B2", -0.25f, RangeRepresentation.RangeObject)]
        [InlineData("R17:T21", "R18,R19,R20,S19", 99999L, RangeRepresentation.RangeObject)]
        [InlineData("A1:A1", "A1", 22, RangeRepresentation.StringExpression)]
        [InlineData("A1:A5", "A2", true, RangeRepresentation.StringExpression)]
        [InlineData("A1:C1", "B1", "test", RangeRepresentation.StringExpression)]
        [InlineData("A1:C3", "B2", -0.25f, RangeRepresentation.StringExpression)]
        [InlineData("R17:T21", "R18,R19,R20,S19", 99999L, RangeRepresentation.StringExpression)]
        public void SetStyleTest2b(string rangeString, string definedCells, object cellValue, RangeRepresentation representation)
        {
            Worksheet worksheet = new Worksheet();
            AddCells(worksheet, cellValue, definedCells);
            int cellCount = worksheet.Cells.Count;
            Assert.NotEqual(0, cellCount);
            Range range = new Range(rangeString);
            if (representation == RangeRepresentation.RangeObject)
            {
                worksheet.SetStyle(range, null);
            }
            else
            {
                worksheet.SetStyle(rangeString, null);
            }
            AssertRemovedStyles(worksheet, cellCount);
        }

        [Theory(DisplayName = "Test of the SetStyle function on a worksheet with existing cells that have a style defined, and a Range object")]
        [InlineData("A1:A1", "A1", 22, null, RangeRepresentation.RangeObject)]
        [InlineData("A1:A5", "A2", true, "A1,A3,A4,A5", RangeRepresentation.RangeObject)]
        [InlineData("A1:C1", "B1", "test", "A1,C1", RangeRepresentation.RangeObject)]
        [InlineData("A1:C3", "B2", -0.25f, "A1,A2,A3,B1,B3,C1,C2,C3", RangeRepresentation.RangeObject)]
        [InlineData("R17:T21", "R18,R19,R20,S19", 99999L, "R17,R21,S17,S18,S20,S21,T17,T18,T19,T20,T21", RangeRepresentation.RangeObject)]
        [InlineData("A1:A1", "A1", 22, null, RangeRepresentation.StringExpression)]
        [InlineData("A1:A5", "A2", true, "A1,A3,A4,A5", RangeRepresentation.StringExpression)]
        [InlineData("A1:C1", "B1", "test", "A1,C1", RangeRepresentation.StringExpression)]
        [InlineData("A1:C3", "B2", -0.25f, "A1,A2,A3,B1,B3,C1,C2,C3", RangeRepresentation.StringExpression)]
        [InlineData("R17:T21", "R18,R19,R20,S19", 99999L, "R17,R21,S17,S18,S20,S21,T17,T18,T19,T20,T21", RangeRepresentation.StringExpression)]
        public void SetStyleTest3(string rangeString, string definedCells, object cellValue, string expectedEmptyCells, RangeRepresentation representation)
        {
            Worksheet worksheet = new Worksheet();
            AddCells(worksheet, cellValue, definedCells, BasicStyles.Italic);
            int cellCount = worksheet.Cells.Count;
            Assert.NotEqual(0, cellCount);
            Range range = new Range(rangeString);
            if (representation == RangeRepresentation.RangeObject)
            {
                worksheet.SetStyle(range, BasicStyles.Bold);
            }
            else
            {
                worksheet.SetStyle(rangeString, BasicStyles.Bold);
            }
            List<string> emptyCells = TestUtils.SplitValuesAsList(expectedEmptyCells);
            AssertCellRange(rangeString, BasicStyles.Bold, worksheet, emptyCells, emptyCells.Count + cellCount);
        }

        [Theory(DisplayName = "Test of the SetStyle function on a worksheet with existing cells that have a style defined, and a Range object or its string representation with null as style")]
        [InlineData("A1:A1", "A1", 22, RangeRepresentation.RangeObject)]
        [InlineData("A1:A5", "A2", true, RangeRepresentation.RangeObject)]
        [InlineData("A1:C1", "B1", "test", RangeRepresentation.RangeObject)]
        [InlineData("A1:C3", "B2", -0.25f, RangeRepresentation.RangeObject)]
        [InlineData("R17:T21", "R18,R19,R20,S19", 99999L, RangeRepresentation.RangeObject)]
        [InlineData("A1:A1", "A1", 22, RangeRepresentation.StringExpression)]
        [InlineData("A1:A5", "A2", true, RangeRepresentation.StringExpression)]
        [InlineData("A1:C1", "B1", "test", RangeRepresentation.StringExpression)]
        [InlineData("A1:C3", "B2", -0.25f, RangeRepresentation.StringExpression)]
        [InlineData("R17:T21", "R18,R19,R20,S19", 99999L, RangeRepresentation.StringExpression)]
        public void SetStyleTest3b(string rangeString, string definedCells, object cellValue, RangeRepresentation representation)
        {
            Worksheet worksheet = new Worksheet();
            AddCells(worksheet, cellValue, definedCells, BasicStyles.Italic);
            int cellCount = worksheet.Cells.Count;
            Assert.NotEqual(0, cellCount);
            Range range = new Range(rangeString);
            if (representation == RangeRepresentation.RangeObject)
            {
                worksheet.SetStyle(range, null);
            }
            else
            {
                worksheet.SetStyle(rangeString, null);
            }
            AssertRemovedStyles(worksheet, cellCount);
        }

        [Theory(DisplayName = "Test of the SetStyle function on a worksheet with existing date and time cells and a Range object")]
        [InlineData(RangeRepresentation.RangeObject)]
        [InlineData(RangeRepresentation.StringExpression)]
        public void SetStyleTest4(RangeRepresentation representation)
        {
            Worksheet worksheet = new Worksheet();
            AddCells(worksheet, new DateTime(2020, 11, 10, 9, 8, 7), "B2");
            AddCells(worksheet, new TimeSpan(10, 11, 12), "B3");
            int cellCount = worksheet.Cells.Count;
            Assert.NotEqual(0, cellCount);
            Range range = new Range("A1:C3");
            if (representation == RangeRepresentation.RangeObject)
            {
                worksheet.SetStyle(range, BasicStyles.BorderFrame);
            }
            else
            {
                worksheet.SetStyle("A1:C3", BasicStyles.BorderFrame);
            }
            List<string> emptyCells = TestUtils.SplitValuesAsList("A1,A2,A3,B1,C1,C2,C3");
            AssertCellRange("A1:C3", BasicStyles.BorderFrame, worksheet, emptyCells, emptyCells.Count + cellCount);
        }

        [Theory(DisplayName = "Test of the SetStyle function on a worksheet with existing date and time cells and a Range object with null as style")]
        [InlineData(RangeRepresentation.RangeObject)]
        [InlineData(RangeRepresentation.StringExpression)]
        public void SetStyleTest4b(RangeRepresentation representation)
        {
            Worksheet worksheet = new Worksheet();
            AddCells(worksheet, new DateTime(2020, 11, 10, 9, 8, 7), "B2");
            AddCells(worksheet, new TimeSpan(10, 11, 12), "B3");
            int cellCount = worksheet.Cells.Count;
            Assert.NotEqual(0, cellCount);
            Range range = new Range("A1:C3");
            if (representation == RangeRepresentation.RangeObject)
            {
                worksheet.SetStyle(range, null);
            }
            else
            {
                worksheet.SetStyle("A1:C3", null);
            }
            AssertRemovedStyles(worksheet, cellCount);
        }

        [Theory(DisplayName = "Test of the SetStyle function on an empty worksheet with a start and end address")]
        [InlineData("A1", "A1")]
        [InlineData("A1", "A5")]
        [InlineData("A1", "C1")]
        [InlineData("A1", "C3")]
        [InlineData("R17", "N22")]
        public void SetStyleTest5(string startAddressString, string endAddressString)
        {
            Worksheet worksheet = new Worksheet();
            Assert.Empty(worksheet.Cells);
            Address startAddress = new Address(startAddressString);
            Address endAddress = new Address(endAddressString);
            worksheet.SetStyle(startAddress, endAddress, BasicStyles.BoldItalic);
            Range range = new Range(startAddress, endAddress);
            List<string> emptyCells = (from i in range.ResolveEnclosedAddresses() select i.GetAddress()).ToList();
            AssertCellRange(range.ToString(), BasicStyles.BoldItalic, worksheet, emptyCells, emptyCells.Count);
        }

        [Theory(DisplayName = "Test of the SetStyle function on an empty worksheet with a start and end address with null as style")]
        [InlineData("A1", "A1")]
        [InlineData("A1", "A5")]
        [InlineData("A1", "C1")]
        [InlineData("A1", "C3")]
        [InlineData("R17", "N22")]
        public void SetStyleTest5b(string startAddressString, string endAddressString)
        {
            Worksheet worksheet = new Worksheet();
            Assert.Empty(worksheet.Cells);
            Address startAddress = new Address(startAddressString);
            Address endAddress = new Address(endAddressString);
            worksheet.SetStyle(startAddress, endAddress, null);
            Assert.Empty(worksheet.Cells); // Should not create empty cells
        }

        [Theory(DisplayName = "Test of the SetStyle function on a worksheet with existing cells, and a start and end address")]
        [InlineData("A1", "A1", "A1", 22, null)]
        [InlineData("A1", "A5", "A2", true, "A1,A3,A4,A5")]
        [InlineData("A1", "C1", "B1", "test", "A1,C1")]
        [InlineData("A1", "C3", "B2", -0.25f, "A1,A2,A3,B1,B3,C1,C2,C3")]
        [InlineData("R17", "T21", "R18,R19,R20,S19", 99999L, "R17,R21,S17,S18,S20,S21,T17,T18,T19,T20,T21")]
        public void SetStyleTest6(string startAddressString, string endAddressString, string definedCells, object cellValue, string expectedEmptyCells)
        {
            Worksheet worksheet = new Worksheet();
            AddCells(worksheet, cellValue, definedCells);
            int cellCount = worksheet.Cells.Count;
            Assert.NotEqual(0, cellCount);
            Address startAddress = new Address(startAddressString);
            Address endAddress = new Address(endAddressString);
            worksheet.SetStyle(startAddress, endAddress, BasicStyles.Bold);
            List<string> emptyCells = TestUtils.SplitValuesAsList(expectedEmptyCells);
            Range range = new Range(startAddress, endAddress);
            AssertCellRange(range.ToString(), BasicStyles.Bold, worksheet, emptyCells, emptyCells.Count + cellCount);
        }

        [Theory(DisplayName = "Test of the SetStyle function on a worksheet with existing cells, and a start and end address with null as style")]
        [InlineData("A1", "A1", "A1", 22)]
        [InlineData("A1", "A5", "A2", true)]
        [InlineData("A1", "C1", "B1", "test")]
        [InlineData("A1", "C3", "B2", -0.25f)]
        [InlineData("R17", "T21", "R18,R19,R20,S19", 99999L)]
        public void SetStyleTest6b(string startAddressString, string endAddressString, string definedCells, object cellValue)
        {
            Worksheet worksheet = new Worksheet();
            AddCells(worksheet, cellValue, definedCells);
            int cellCount = worksheet.Cells.Count;
            Assert.NotEqual(0, cellCount);
            Address startAddress = new Address(startAddressString);
            Address endAddress = new Address(endAddressString);
            worksheet.SetStyle(startAddress, endAddress, null);
            AssertRemovedStyles(worksheet, cellCount);
        }

        [Theory(DisplayName = "Test of the SetStyle function on a worksheet with existing cells that have a style defined, and a start and end address")]
        [InlineData("A1", "A1", "A1", 22, null)]
        [InlineData("A1", "A5", "A2", true, "A1,A3,A4,A5")]
        [InlineData("A1", "C1", "B1", "test", "A1,C1")]
        [InlineData("A1", "C3", "B2", -0.25f, "A1,A2,A3,B1,B3,C1,C2,C3")]
        [InlineData("R17", "T21", "R18,R19,R20,S19", 99999L, "R17,R21,S17,S18,S20,S21,T17,T18,T19,T20,T21")]
        public void SetStyleTest7(string startAddressString, string endAddressString, string definedCells, object cellValue, string expectedEmptyCells)
        {
            Worksheet worksheet = new Worksheet();
            AddCells(worksheet, cellValue, definedCells, BasicStyles.Italic);
            int cellCount = worksheet.Cells.Count;
            Assert.NotEqual(0, cellCount);
            Address startAddress = new Address(startAddressString);
            Address endAddress = new Address(endAddressString);
            worksheet.SetStyle(startAddress, endAddress, BasicStyles.Bold);
            List<string> emptyCells = TestUtils.SplitValuesAsList(expectedEmptyCells);
            Range range = new Range(startAddress, endAddress);
            AssertCellRange(range.ToString(), BasicStyles.Bold, worksheet, emptyCells, emptyCells.Count + cellCount);
        }

        [Theory(DisplayName = "Test of the SetStyle function on a worksheet with existing cells that have a style defined, and a start and end address with null as style")]
        [InlineData("A1", "A1", "A1", 22)]
        [InlineData("A1", "A5", "A2", true)]
        [InlineData("A1", "C1", "B1", "test")]
        [InlineData("A1", "C3", "B2", -0.25f)]
        [InlineData("R17", "T21", "R18,R19,R20,S19", 99999L)]
        public void SetStyleTest7b(string startAddressString, string endAddressString, string definedCells, object cellValue)
        {
            Worksheet worksheet = new Worksheet();
            AddCells(worksheet, cellValue, definedCells, BasicStyles.Italic);
            int cellCount = worksheet.Cells.Count;
            Assert.NotEqual(0, cellCount);
            Address startAddress = new Address(startAddressString);
            Address endAddress = new Address(endAddressString);
            worksheet.SetStyle(startAddress, endAddress, null);
            AssertRemovedStyles(worksheet, cellCount);
        }

        [Fact(DisplayName = "Test of the SetStyle function with a null style (removes style) on a worksheet with existing cells, and a start and end address")]
        public void SetStyleTest8()
        {
            Worksheet worksheet = new Worksheet();
            AddCells(worksheet, "test", "B2");
            AddCells(worksheet, 123, "B3");
            int cellCount = worksheet.Cells.Count;
            Assert.NotEqual(0, cellCount);
            Address startAddress = new Address("A1");
            Address endAddress = new Address("C3");
            worksheet.SetStyle(startAddress, endAddress, null);
            List<string> emptyCells = new List<string>();
            AssertCellRange("B2:B3", null, worksheet, emptyCells, cellCount);
        }

        [Fact(DisplayName = "Test of the SetStyle function on a worksheet with existing date and time cells, and a start and end address")]
        public void SetStyleTest9()
        {
            Worksheet worksheet = new Worksheet();
            AddCells(worksheet, new DateTime(2020, 11, 10, 9, 8, 7), "B2");
            AddCells(worksheet, new TimeSpan(10, 11, 12), "B3");
            int cellCount = worksheet.Cells.Count;
            Assert.NotEqual(0, cellCount);
            Address startAddress = new Address("A1");
            Address endAddress = new Address("C3");
            worksheet.SetStyle(startAddress, endAddress, BasicStyles.BorderFrame);
            List<string> emptyCells = TestUtils.SplitValuesAsList("A1,A2,A3,B1,C1,C2,C3");
            AssertCellRange("A1:C3", BasicStyles.BorderFrame, worksheet, emptyCells, emptyCells.Count + cellCount);
        }

        [Fact(DisplayName = "Test of the SetStyle function on a worksheet with existing date and time cells, and a start and end address with null as style")]
        public void SetStyleTest9b()
        {
            Worksheet worksheet = new Worksheet();
            AddCells(worksheet, new DateTime(2020, 11, 10, 9, 8, 7), "B2");
            AddCells(worksheet, new TimeSpan(10, 11, 12), "B3");
            int cellCount = worksheet.Cells.Count;
            Assert.NotEqual(0, cellCount);
            Address startAddress = new Address("A1");
            Address endAddress = new Address("C3");
            worksheet.SetStyle(startAddress, endAddress, null);
            AssertRemovedStyles(worksheet, cellCount);
        }

        [Theory(DisplayName = "Test of the SetStyle function on an empty worksheet with a singular address or its string representation")]
        [InlineData(RangeRepresentation.RangeObject)]
        [InlineData(RangeRepresentation.StringExpression)]
        public void SetStyleTest10(RangeRepresentation representation)
        {
            Worksheet worksheet = new Worksheet();
            Assert.Empty(worksheet.Cells);
            if (representation == RangeRepresentation.RangeObject)
            {
                worksheet.SetStyle(new Address("C2"), BasicStyles.BoldItalic);
            }
            else
            {
                worksheet.SetStyle("C2", BasicStyles.BoldItalic);
            }
            Range range = new Range("C2:C2");
            List<string> emptyCells = (from i in range.ResolveEnclosedAddresses() select i.GetAddress()).ToList();
            AssertCellRange(range.ToString(), BasicStyles.BoldItalic, worksheet, emptyCells, emptyCells.Count);
        }

        [Theory(DisplayName = "Test of the SetStyle function on an empty worksheet with a singular address or its string representation with null as style")]
        [InlineData(RangeRepresentation.RangeObject)]
        [InlineData(RangeRepresentation.StringExpression)]
        public void SetStyleTest10b(RangeRepresentation representation)
        {
            Worksheet worksheet = new Worksheet();
            Assert.Empty(worksheet.Cells);
            if (representation == RangeRepresentation.RangeObject)
            {
                worksheet.SetStyle(new Address("C2"), null);
            }
            else
            {
                worksheet.SetStyle("C2", null);
            }
            Assert.Empty(worksheet.Cells);
        }

        [Theory(DisplayName = "Test of the SetStyle function on a worksheet with existing cells, and a singular address or its string representation")]
        [InlineData(RangeRepresentation.RangeObject)]
        [InlineData(RangeRepresentation.StringExpression)]
        public void SetStyleTest11(RangeRepresentation representation)
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddCell(22, "B2");
            worksheet.AddCell(false, "B3");
            worksheet.AddCell("test", "B4");
            int cellCount = worksheet.Cells.Count;
            Assert.NotEqual(0, cellCount);
            if (representation == RangeRepresentation.RangeObject)
            {
                worksheet.SetStyle(new Address("B2"), BasicStyles.Bold);
            }
            else
            {
                worksheet.SetStyle("B2", BasicStyles.Bold);
            }
            AssertCellRange("B2:B2", BasicStyles.Bold, worksheet, new List<string>(), 3);
        }

        [Theory(DisplayName = "Test of the SetStyle function on a worksheet with existing cells, and a singular address or its string representation with null as style")]
        [InlineData(RangeRepresentation.RangeObject)]
        [InlineData(RangeRepresentation.StringExpression)]
        public void SetStyleTest11b(RangeRepresentation representation)
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddCell(22, "B2");
            worksheet.AddCell(false, "B3");
            worksheet.AddCell("test", "B4");
            int cellCount = worksheet.Cells.Count;
            Assert.NotEqual(0, cellCount);
            if (representation == RangeRepresentation.RangeObject)
            {
                worksheet.SetStyle(new Address("B2"), null);
            }
            else
            {
                worksheet.SetStyle("B2", null);
            }
            AssertRemovedStyles(worksheet, cellCount);
        }

        [Theory(DisplayName = "Test of the SetStyle function on a worksheet with existing cells that have a style defined, and a singular address or its string representation")]
        [InlineData(RangeRepresentation.RangeObject)]
        [InlineData(RangeRepresentation.StringExpression)]
        public void SetStyleTest12(RangeRepresentation representation)
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddCell(22, "B2", BasicStyles.BoldItalic);
            worksheet.AddCell(true, "B3", BasicStyles.BoldItalic);
            worksheet.AddCell("test", "B4", BasicStyles.BoldItalic);
            int cellCount = worksheet.Cells.Count;
            Assert.NotEqual(0, cellCount);
            if (representation == RangeRepresentation.RangeObject)
            {
                worksheet.SetStyle(new Address("B2"), BasicStyles.Bold);
            }
            else
            {
                worksheet.SetStyle("B2", BasicStyles.Bold);
            }
            AssertCellRange("B2:B2", BasicStyles.Bold, worksheet, new List<string>(), 3);
        }

        [Theory(DisplayName = "Test of the SetStyle function on a worksheet with existing cells that have a style defined, and a singular address or its string representation with null as style")]
        [InlineData(RangeRepresentation.RangeObject)]
        [InlineData(RangeRepresentation.StringExpression)]
        public void SetStyleTest12b(RangeRepresentation representation)
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddCell(22, "B2", BasicStyles.BoldItalic);
            worksheet.AddCell(true, "B3", BasicStyles.BoldItalic);
            worksheet.AddCell("test", "B4", BasicStyles.BoldItalic);
            int cellCount = worksheet.Cells.Count;
            Assert.NotEqual(0, cellCount);
            if (representation == RangeRepresentation.RangeObject)
            {
                worksheet.SetStyle(new Address("B2"), null);
            }
            else
            {
                worksheet.SetStyle("B2", null);
            }
            Assert.Null(worksheet.Cells["B2"].CellStyle);
            Assert.True(worksheet.Cells["B3"].CellStyle.Equals(BasicStyles.BoldItalic));
            Assert.True(worksheet.Cells["B4"].CellStyle.Equals(BasicStyles.BoldItalic));
        }

        [Theory(DisplayName = "Test of the SetStyle function on a worksheet with existing date and time cells, and a singular address or its string representation")]
        [InlineData(RangeRepresentation.RangeObject)]
        [InlineData(RangeRepresentation.StringExpression)]
        public void SetStyleTest13(RangeRepresentation representation)
        {
            Worksheet worksheet = new Worksheet();
            AddCells(worksheet, new DateTime(2020, 11, 10, 9, 8, 7), "B2");
            AddCells(worksheet, new TimeSpan(10, 11, 12), "B3");
            int cellCount = worksheet.Cells.Count;
            Assert.NotEqual(0, cellCount);
            if (representation == RangeRepresentation.RangeObject)
            {
                worksheet.SetStyle(new Address("B2"), BasicStyles.BorderFrame);
                worksheet.SetStyle(new Address("B3"), BasicStyles.BorderFrame);
            }
            else
            {
                worksheet.SetStyle("B2", BasicStyles.BorderFrame);
                worksheet.SetStyle("B3", BasicStyles.BorderFrame);
            }

            AssertCellRange("B2:B3", BasicStyles.BorderFrame, worksheet, new List<string>(), cellCount); ;
        }

        [Theory(DisplayName = "Test of the SetStyle function on a worksheet with existing date and time cells, and a singular address or its string representation with null as style")]
        [InlineData(RangeRepresentation.RangeObject)]
        [InlineData(RangeRepresentation.StringExpression)]
        public void SetStyleTest13b(RangeRepresentation representation)
        {
            Worksheet worksheet = new Worksheet();
            AddCells(worksheet, new DateTime(2020, 11, 10, 9, 8, 7), "B2");
            AddCells(worksheet, new TimeSpan(10, 11, 12), "B3");
            int cellCount = worksheet.Cells.Count;
            Assert.NotEqual(0, cellCount);
            if (representation == RangeRepresentation.RangeObject)
            {
                worksheet.SetStyle(new Address("B2"), null);
                worksheet.SetStyle(new Address("B3"), null);
            }
            else
            {
                worksheet.SetStyle("B2", null);
                worksheet.SetStyle("B3", null);
            }
            AssertRemovedStyles(worksheet, cellCount);
        }

        [Fact(DisplayName = "Test of the failing SetStyle function, when no range as string was defined")]
        public void SetStyleFailTest()
        {
            Worksheet worksheet = new Worksheet();
            string range = null;
            Assert.Throws<FormatException>(() => worksheet.SetStyle(range, BasicStyles.Bold));
        }

        private void AssertCellRange(string range, Style expectedStyle, Worksheet worksheet, List<string> createdCells, int expectedSize)
        {
            Assert.Equal(expectedSize, worksheet.Cells.Count);
            Range setRange = new Range(range);
            foreach (Address address in setRange.ResolveEnclosedAddresses())
            {
                Assert.Contains(worksheet.Cells, item => item.Key.Equals(address.GetAddress(), StringComparison.Ordinal));
                if (expectedStyle == null)
                {
                    Assert.Null(worksheet.Cells[address.GetAddress()].CellStyle);
                }
                else
                {
                    Assert.True(expectedStyle.Equals(worksheet.Cells[address.GetAddress()].CellStyle));
                }
                if (createdCells != null && createdCells.Contains(address.GetAddress()))
                {
                    Assert.Equal(Cell.CellType.Empty, worksheet.Cells[address.GetAddress()].DataType);
                }
            }
        }

        private void AssertRemovedStyles(Worksheet worksheet, int expectedSize)
        {
            Assert.Equal(expectedSize, worksheet.Cells.Count);
            foreach (KeyValuePair<string, Cell> cell in worksheet.Cells)
            {
                Assert.Null(cell.Value.CellStyle);
            }
        }

        private static void AddCells(Worksheet worksheet, object sample, string addressString, Style style = null)
        {
            List<string> addresses = TestUtils.SplitValuesAsList(addressString);
            foreach (string address in addresses)
            {
                Cell cell = new Cell(sample, Cell.CellType.Default);
                worksheet.AddCell(cell, address, style);
            }
        }

    }
}
