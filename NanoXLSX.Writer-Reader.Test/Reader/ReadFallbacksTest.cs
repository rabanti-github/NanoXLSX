﻿using System.IO;
using NanoXLSX.Test.Writer_Reader.Utils;
using Xunit;
using static NanoXLSX.Styles.Border;
using static NanoXLSX.Styles.CellXf;
using static NanoXLSX.Styles.Fill;
using static NanoXLSX.Styles.Font;

namespace NanoXLSX.Test.Writer_Reader.ReaderTest
{
    public class ReadFallbacksTest
    {

        [Fact(DisplayName = "Test of the fallback behavior on unexpected border types")]
        public void ReadUnknownBorderTypeTest()
        {
            // Cell A1 contains a border style with unknown line type
            // This causes neither in Excel a crash, nor should the library crash
            Cell cell = getCell("unknown_style_enums.xlsx");
            Assert.Equal(StyleValue.none, cell.CellStyle.CurrentBorder.TopStyle);
            Assert.Equal(StyleValue.none, cell.CellStyle.CurrentBorder.BottomStyle);
            Assert.Equal(StyleValue.none, cell.CellStyle.CurrentBorder.LeftStyle);
            Assert.Equal(StyleValue.none, cell.CellStyle.CurrentBorder.RightStyle);
            Assert.Equal(StyleValue.none, cell.CellStyle.CurrentBorder.DiagonalStyle);
        }

        [Fact(DisplayName = "Test of the fallback behavior on unexpected pattern fill types")]
        public void ReadUnknownPatternFillTypeTest()
        {
            // The file contains a pattern fill definition with an unknown value
            // This causes neither in Excel a crash, nor should the library crash
            Cell cell = getCell("unknown_style_enums.xlsx");
            Assert.Equal(PatternValue.none, cell.CellStyle.CurrentFill.PatternFill);
        }

        [Fact(DisplayName = "Test of the fallback behavior on unexpected vertical align font types")]
        public void ReadUnknownFontVerticalAlignTypeTest()
        {
            // The file contains a font definition with an unknown vertical align value
            // This causes an auto-fixing action in Excel (but not a crash). The library will auto-fix this too
            Cell cell = getCell("unknown_style_enums.xlsx");
            Assert.Equal(VerticalTextAlignValue.none, cell.CellStyle.CurrentFont.VerticalAlign);
        }


        [Fact(DisplayName = "Test of the fallback behavior on unexpected horizontal align cellXF types")]
        public void ReadUnknownCellXfHorizontalAlignTypeTest()
        {
            // The file contains a CellXF definition with an unknown horizontal align value
            // This causes neither in Excel a crash, nor should the library crash
            Cell cell = getCell("unknown_style_enums.xlsx");
            Assert.Equal(HorizontalAlignValue.none, cell.CellStyle.CurrentCellXf.HorizontalAlign);
        }

        [Fact(DisplayName = "Test of the fallback behavior on unexpected vertical align cellXF types")]
        public void ReadUnknownCellXfVerticalAlignTypeTest()
        {
            // The file contains a CellXF definition with an unknown vertical align value
            // This causes neither in Excel a crash, nor should the library crash
            Cell cell = getCell("unknown_style_enums.xlsx");
            Assert.Equal(VerticalAlignValue.none, cell.CellStyle.CurrentCellXf.VerticalAlign);
        }

        [Fact(DisplayName = "Test of the fallback behavior on missing ID references in the CellXF section")]
        public void IgnoreMissingStyleRefsTest()
        {
            // The file contains 5 complex styles (and 1 default style), assigned to 5 cells. In the CellXF
            // section is for each style one particular reference ID (e.g. fontId) omitted. This should not
            // lead to a crash 
            Stream stream = TestUtils.GetResource("omitted_style_refs.xlsx");
            Workbook workbook = WorkbookReader.Load(stream);
            Assert.NotNull(workbook.Worksheets[0].Cells["A1"].CellStyle);
            Assert.NotNull(workbook.Worksheets[0].Cells["A2"].CellStyle);
            Assert.NotNull(workbook.Worksheets[0].Cells["A3"].CellStyle);
            Assert.NotNull(workbook.Worksheets[0].Cells["A4"].CellStyle);
            Assert.NotNull(workbook.Worksheets[0].Cells["A5"].CellStyle);
        }

        private static Cell getCell(string resourceName)
        {
            Stream stream = TestUtils.GetResource(resourceName);
            Workbook workbook = WorkbookReader.Load(stream);
            Cell cell = workbook.Worksheets[0].Cells["A1"];
            return cell;
        }

    }
}
