using NanoXLSX;
using NanoXLSX.Styles;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace NanoXLSX_Test.Reader
{
    public class ReadFallbacksTest
    {

        [Fact(DisplayName = "Test of the fallback behavior on unexpected border types")]
        public void ReadUnknownBorderTypeTest()
        {
            // Cell A1 contains a border style with unknown line type
            // This causes neither in Excel a crash, nor should the library crash
            Cell cell = getCell("unknown_style_enums.xlsx");
            Assert.Equal(Border.StyleValue.none, cell.CellStyle.CurrentBorder.TopStyle);
            Assert.Equal(Border.StyleValue.none, cell.CellStyle.CurrentBorder.BottomStyle);
            Assert.Equal(Border.StyleValue.none, cell.CellStyle.CurrentBorder.LeftStyle);
            Assert.Equal(Border.StyleValue.none, cell.CellStyle.CurrentBorder.RightStyle);
            Assert.Equal(Border.StyleValue.none, cell.CellStyle.CurrentBorder.DiagonalStyle);
        }

        [Fact(DisplayName = "Test of the fallback behavior on unexpected pattern fill types")]
        public void ReadUnknownPatternFillTypeTest()
        {
            // The file contains a pattern fill definition with an unknown value
            // This causes neither in Excel a crash, nor should the library crash
            Cell cell = getCell("unknown_style_enums.xlsx");
            Assert.Equal(Fill.PatternValue.none, cell.CellStyle.CurrentFill.PatternFill);
        }

        [Fact(DisplayName = "Test of the fallback behavior on unexpected vertical align font types")]
        public void ReadUnknownFontVerticalAlignTypeTest()
        {
            // The file contains a font definition with an unknown vertical align value
            // This causes an auto-fixing action in Excel (but not a crash). The library will auto-fix this too
            Cell cell = getCell("unknown_style_enums.xlsx");
            Assert.Equal(Font.VerticalAlignValue.none, cell.CellStyle.CurrentFont.VerticalAlign);
        }


        [Fact(DisplayName = "Test of the fallback behavior on unexpected horizontal align cellXF types")]
        public void ReadUnknownCellXfHorizontalAlignTypeTest()
        {
            // The file contains a CellXF definition with an unknown horizontal align value
            // This causes neither in Excel a crash, nor should the library crash
            Cell cell = getCell("unknown_style_enums.xlsx");
            Assert.Equal(CellXf.HorizontalAlignValue.none, cell.CellStyle.CurrentCellXf.HorizontalAlign);
        }

        [Fact(DisplayName = "Test of the fallback behavior on unexpected vertical align cellXF types")]
        public void ReadUnknownCellXfVerticalAlignTypeTest()
        {
            // The file contains a CellXF definition with an unknown vertical align value
            // This causes neither in Excel a crash, nor should the library crash
            Cell cell = getCell("unknown_style_enums.xlsx");
            Assert.Equal(CellXf.VerticalAlignValue.none, cell.CellStyle.CurrentCellXf.VerticalAlign);
        }

        private static Cell getCell(string resourceName)
        {
            Stream stream = TestUtils.GetResource(resourceName);
            Workbook workbook = Workbook.Load(stream);
            Cell cell = workbook.Worksheets[0].Cells["A1"];
            return cell;
        }

    }
}
