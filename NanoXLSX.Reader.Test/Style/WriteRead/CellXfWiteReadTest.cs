using NanoXLSX;
using NanoXLSX.Styles;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace NanoXLSX_Test.Styles.WriteRead
{
    public class CellXfWiteReadTest
    {
        [Theory(DisplayName = "Test of the 'ForceApplyAlignment' value when writing and reading a CellXF style")]
        [InlineData(true, "test")]
        [InlineData(false, 0.5f)]
        public void ForceApplyAlignmentCellXfTest(bool styleValue, object value)
        {
            Style style = new Style();
            style.CurrentCellXf.ForceApplyAlignment = styleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentCellXf.ForceApplyAlignment);
        }

        [Theory(DisplayName = "Test of the 'Hidden' and 'Locked' values when writing and reading a CellXF style")]
        [InlineData(false, false, "test")]
        [InlineData(false, true, 0.5f)]
        [InlineData(true, false, 22)]
        [InlineData(true, true, true)]
        public void HiddenCellXfTest(bool hiddenStyleValue, bool lockedStyleValue, object value)
        {
            Style style = new Style();
            style.CurrentCellXf.Hidden = hiddenStyleValue;
            style.CurrentCellXf.Locked = lockedStyleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(hiddenStyleValue, cell.CellStyle.CurrentCellXf.Hidden);
            Assert.Equal(lockedStyleValue, cell.CellStyle.CurrentCellXf.Locked);
        }

        [Theory(DisplayName = "Test of the 'Alignment' value when writing and reading a CellXF style")]
        [InlineData(TextBreakValue.shrinkToFit, "test")]
        [InlineData(TextBreakValue.wrapText, 0.5f)]
        [InlineData(TextBreakValue.none, true)]
        public void AlignmentCellXfTest(TextBreakValue styleValue, object value)
        {
            Style style = new Style();
            style.CurrentCellXf.Alignment = styleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentCellXf.Alignment);
        }

        [Theory(DisplayName = "Test of the 'HorizontalAlign' value when writing and reading a CellXF style")]
        [InlineData(HorizontalAlignValue.justify, "test")]
        [InlineData(HorizontalAlignValue.center, 0.5f)]
        [InlineData(HorizontalAlignValue.centerContinuous, true)]
        [InlineData(HorizontalAlignValue.distributed, 22)]
        [InlineData(HorizontalAlignValue.fill, false)]
        [InlineData(HorizontalAlignValue.general, "")]
        [InlineData(HorizontalAlignValue.left, -2.11f)]
        [InlineData(HorizontalAlignValue.right, "test")]
        [InlineData(HorizontalAlignValue.none, " ")]
        public void HorizontalAlignCellXfTest(HorizontalAlignValue styleValue, object value)
        {
            Style style = new Style();
            style.CurrentCellXf.HorizontalAlign = styleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentCellXf.HorizontalAlign);
        }

        [Theory(DisplayName = "Test of the 'VerticalAlign' value when writing and reading a CellXF style")]
        [InlineData(VerticalAlignValue.justify, "test")]
        [InlineData(VerticalAlignValue.center, 0.5f)]
        [InlineData(VerticalAlignValue.bottom, true)]
        [InlineData(VerticalAlignValue.top, 22)]
        [InlineData(VerticalAlignValue.distributed, false)]
        [InlineData(VerticalAlignValue.none, " ")]
        public void VerticalAlignCellXfTest(VerticalAlignValue styleValue, object value)
        {
            Style style = new Style();
            style.CurrentCellXf.VerticalAlign = styleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentCellXf.VerticalAlign);
        }

        [Theory(DisplayName = "Test of the 'Indent' value when writing and reading a CellXF style")]
        [InlineData(0, HorizontalAlignValue.left, 0, "test")]
        [InlineData(0, HorizontalAlignValue.right, 0, "test")]
        [InlineData(0, HorizontalAlignValue.distributed, 0, "test")]
        [InlineData(0, HorizontalAlignValue.center, 0, "test")]
        [InlineData(1, HorizontalAlignValue.left, 1, 0.5f)]
        [InlineData(1, HorizontalAlignValue.right, 1, 0.5f)]
        [InlineData(1, HorizontalAlignValue.distributed, 1, 0.5f)]
        [InlineData(1, HorizontalAlignValue.center, 0, 0.5f)]
        [InlineData(5, HorizontalAlignValue.left, 5, true)]
        [InlineData(5, HorizontalAlignValue.right, 5, true)]
        [InlineData(5, HorizontalAlignValue.distributed, 5, true)]
        [InlineData(5, HorizontalAlignValue.center, 0, true)]
        [InlineData(64, HorizontalAlignValue.left, 64, 22)]
        [InlineData(64, HorizontalAlignValue.right, 64, 22)]
        [InlineData(64, HorizontalAlignValue.distributed, 64, 22)]
        [InlineData(64, HorizontalAlignValue.center, 0, 22)]
        public void IndentCellXfTest(int styleValue, HorizontalAlignValue alignValue, int expectedIndent, object value)
        {
            Style style = new Style();
            style.CurrentCellXf.HorizontalAlign = alignValue;
            style.CurrentCellXf.Indent = styleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(expectedIndent, cell.CellStyle.CurrentCellXf.Indent);
            Assert.Equal(alignValue, cell.CellStyle.CurrentCellXf.HorizontalAlign);
        }

        [Theory(DisplayName = "Test of the 'TextRotation' value when writing and reading a CellXF style")]
        [InlineData(0, "test")]
        [InlineData(1, 0.5f)]
        [InlineData(-1, true)]
        [InlineData(45,22)]
        [InlineData(-45, -0.11f)]
        [InlineData(90, "")]
        [InlineData(-90, " ")]
        public void TextRotationCellXfTest(int styleValue, object value)
        {
            Style style = new Style();
            style.CurrentCellXf.TextRotation = styleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentCellXf.TextRotation);
        }

    }
}
