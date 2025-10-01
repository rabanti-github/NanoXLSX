using NanoXLSX;
using NanoXLSX.Styles;
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
        [InlineData(CellXf.TextBreakValue.shrinkToFit, "test")]
        [InlineData(CellXf.TextBreakValue.wrapText, 0.5f)]
        [InlineData(CellXf.TextBreakValue.none, true)]
        public void AlignmentCellXfTest(CellXf.TextBreakValue styleValue, object value)
        {
            Style style = new Style();
            style.CurrentCellXf.Alignment = styleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentCellXf.Alignment);
        }

        [Theory(DisplayName = "Test of the 'HorizontalAlign' value when writing and reading a CellXF style")]
        [InlineData(CellXf.HorizontalAlignValue.justify, "test")]
        [InlineData(CellXf.HorizontalAlignValue.center, 0.5f)]
        [InlineData(CellXf.HorizontalAlignValue.centerContinuous, true)]
        [InlineData(CellXf.HorizontalAlignValue.distributed, 22)]
        [InlineData(CellXf.HorizontalAlignValue.fill, false)]
        [InlineData(CellXf.HorizontalAlignValue.general, "")]
        [InlineData(CellXf.HorizontalAlignValue.left, -2.11f)]
        [InlineData(CellXf.HorizontalAlignValue.right, "test")]
        [InlineData(CellXf.HorizontalAlignValue.none, " ")]
        public void HorizontalAlignCellXfTest(CellXf.HorizontalAlignValue styleValue, object value)
        {
            Style style = new Style();
            style.CurrentCellXf.HorizontalAlign = styleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentCellXf.HorizontalAlign);
        }

        [Theory(DisplayName = "Test of the 'VerticalAlign' value when writing and reading a CellXF style")]
        [InlineData(CellXf.VerticalAlignValue.justify, "test")]
        [InlineData(CellXf.VerticalAlignValue.center, 0.5f)]
        [InlineData(CellXf.VerticalAlignValue.bottom, true)]
        [InlineData(CellXf.VerticalAlignValue.top, 22)]
        [InlineData(CellXf.VerticalAlignValue.distributed, false)]
        [InlineData(CellXf.VerticalAlignValue.none, " ")]
        public void VerticalAlignCellXfTest(CellXf.VerticalAlignValue styleValue, object value)
        {
            Style style = new Style();
            style.CurrentCellXf.VerticalAlign = styleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentCellXf.VerticalAlign);
        }

        [Theory(DisplayName = "Test of the 'Indent' value when writing and reading a CellXF style")]
        [InlineData(0, CellXf.HorizontalAlignValue.left, 0, "test")]
        [InlineData(0, CellXf.HorizontalAlignValue.right, 0, "test")]
        [InlineData(0, CellXf.HorizontalAlignValue.distributed, 0, "test")]
        [InlineData(0, CellXf.HorizontalAlignValue.center, 0, "test")]
        [InlineData(1, CellXf.HorizontalAlignValue.left, 1, 0.5f)]
        [InlineData(1, CellXf.HorizontalAlignValue.right, 1, 0.5f)]
        [InlineData(1, CellXf.HorizontalAlignValue.distributed, 1, 0.5f)]
        [InlineData(1, CellXf.HorizontalAlignValue.center, 0, 0.5f)]
        [InlineData(5, CellXf.HorizontalAlignValue.left, 5, true)]
        [InlineData(5, CellXf.HorizontalAlignValue.right, 5, true)]
        [InlineData(5, CellXf.HorizontalAlignValue.distributed, 5, true)]
        [InlineData(5, CellXf.HorizontalAlignValue.center, 0, true)]
        [InlineData(64, CellXf.HorizontalAlignValue.left, 64, 22)]
        [InlineData(64, CellXf.HorizontalAlignValue.right, 64, 22)]
        [InlineData(64, CellXf.HorizontalAlignValue.distributed, 64, 22)]
        [InlineData(64, CellXf.HorizontalAlignValue.center, 0, 22)]
        public void IndentCellXfTest(int styleValue, CellXf.HorizontalAlignValue alignValue, int expectedIndent, object value)
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
        [InlineData(45, 22)]
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
