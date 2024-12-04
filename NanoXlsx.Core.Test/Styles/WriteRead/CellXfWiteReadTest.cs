using NanoXLSX;
using NanoXLSX.Shared.Enums.Styles;
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
        [InlineData(CellXfEnums.TextBreakValue.shrinkToFit, "test")]
        [InlineData(CellXfEnums.TextBreakValue.wrapText, 0.5f)]
        [InlineData(CellXfEnums.TextBreakValue.none, true)]
        public void AlignmentCellXfTest(CellXfEnums.TextBreakValue styleValue, object value)
        {
            Style style = new Style();
            style.CurrentCellXf.Alignment = styleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentCellXf.Alignment);
        }

        [Theory(DisplayName = "Test of the 'HorizontalAlign' value when writing and reading a CellXF style")]
        [InlineData(CellXfEnums.HorizontalAlignValue.justify, "test")]
        [InlineData(CellXfEnums.HorizontalAlignValue.center, 0.5f)]
        [InlineData(CellXfEnums.HorizontalAlignValue.centerContinuous, true)]
        [InlineData(CellXfEnums.HorizontalAlignValue.distributed, 22)]
        [InlineData(CellXfEnums.HorizontalAlignValue.fill, false)]
        [InlineData(CellXfEnums.HorizontalAlignValue.general, "")]
        [InlineData(CellXfEnums.HorizontalAlignValue.left, -2.11f)]
        [InlineData(CellXfEnums.HorizontalAlignValue.right, "test")]
        [InlineData(CellXfEnums.HorizontalAlignValue.none, " ")]
        public void HorizontalAlignCellXfTest(CellXfEnums.HorizontalAlignValue styleValue, object value)
        {
            Style style = new Style();
            style.CurrentCellXf.HorizontalAlign = styleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentCellXf.HorizontalAlign);
        }

        [Theory(DisplayName = "Test of the 'VerticalAlign' value when writing and reading a CellXF style")]
        [InlineData(CellXfEnums.VerticalAlignValue.justify, "test")]
        [InlineData(CellXfEnums.VerticalAlignValue.center, 0.5f)]
        [InlineData(CellXfEnums.VerticalAlignValue.bottom, true)]
        [InlineData(CellXfEnums.VerticalAlignValue.top, 22)]
        [InlineData(CellXfEnums.VerticalAlignValue.distributed, false)]
        [InlineData(CellXfEnums.VerticalAlignValue.none, " ")]
        public void VerticalAlignCellXfTest(CellXfEnums.VerticalAlignValue styleValue, object value)
        {
            Style style = new Style();
            style.CurrentCellXf.VerticalAlign = styleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentCellXf.VerticalAlign);
        }

        [Theory(DisplayName = "Test of the 'Indent' value when writing and reading a CellXF style")]
        [InlineData(0, CellXfEnums.HorizontalAlignValue.left, 0, "test")]
        [InlineData(0, CellXfEnums.HorizontalAlignValue.right, 0, "test")]
        [InlineData(0, CellXfEnums.HorizontalAlignValue.distributed, 0, "test")]
        [InlineData(0, CellXfEnums.HorizontalAlignValue.center, 0, "test")]
        [InlineData(1, CellXfEnums.HorizontalAlignValue.left, 1, 0.5f)]
        [InlineData(1, CellXfEnums.HorizontalAlignValue.right, 1, 0.5f)]
        [InlineData(1, CellXfEnums.HorizontalAlignValue.distributed, 1, 0.5f)]
        [InlineData(1, CellXfEnums.HorizontalAlignValue.center, 0, 0.5f)]
        [InlineData(5, CellXfEnums.HorizontalAlignValue.left, 5, true)]
        [InlineData(5, CellXfEnums.HorizontalAlignValue.right, 5, true)]
        [InlineData(5, CellXfEnums.HorizontalAlignValue.distributed, 5, true)]
        [InlineData(5, CellXfEnums.HorizontalAlignValue.center, 0, true)]
        [InlineData(64, CellXfEnums.HorizontalAlignValue.left, 64, 22)]
        [InlineData(64, CellXfEnums.HorizontalAlignValue.right, 64, 22)]
        [InlineData(64, CellXfEnums.HorizontalAlignValue.distributed, 64, 22)]
        [InlineData(64, CellXfEnums.HorizontalAlignValue.center, 0, 22)]
        public void IndentCellXfTest(int styleValue, CellXfEnums.HorizontalAlignValue alignValue, int expectedIndent, object value)
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
