using NanoXLSX.Styles;
using NanoXLSX.Test.Writer_Reader.Utils;
using Xunit;
using static NanoXLSX.Styles.CellXf;

namespace NanoXLSX.Test.Writer_Reader.StyleTest
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
        [InlineData(TextBreakValue.ShrinkToFit, "test")]
        [InlineData(TextBreakValue.WrapText, 0.5f)]
        [InlineData(TextBreakValue.None, true)]
        public void AlignmentCellXfTest(TextBreakValue styleValue, object value)
        {
            Style style = new Style();
            style.CurrentCellXf.Alignment = styleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentCellXf.Alignment);
        }

        [Theory(DisplayName = "Test of the 'HorizontalAlign' value when writing and reading a CellXF style")]
        [InlineData(HorizontalAlignValue.Justify, "test")]
        [InlineData(HorizontalAlignValue.Center, 0.5f)]
        [InlineData(HorizontalAlignValue.CenterContinuous, true)]
        [InlineData(HorizontalAlignValue.Distributed, 22)]
        [InlineData(HorizontalAlignValue.Fill, false)]
        [InlineData(HorizontalAlignValue.General, "")]
        [InlineData(HorizontalAlignValue.Left, -2.11f)]
        [InlineData(HorizontalAlignValue.Right, "test")]
        [InlineData(HorizontalAlignValue.None, " ")]
        public void HorizontalAlignCellXfTest(HorizontalAlignValue styleValue, object value)
        {
            Style style = new Style();
            style.CurrentCellXf.HorizontalAlign = styleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentCellXf.HorizontalAlign);
        }

        [Theory(DisplayName = "Test of the 'VerticalAlign' value when writing and reading a CellXF style")]
        [InlineData(VerticalAlignValue.Justify, "test")]
        [InlineData(VerticalAlignValue.Center, 0.5f)]
        [InlineData(VerticalAlignValue.Bottom, true)]
        [InlineData(VerticalAlignValue.Top, 22)]
        [InlineData(VerticalAlignValue.Distributed, false)]
        [InlineData(VerticalAlignValue.None, " ")]
        public void VerticalAlignCellXfTest(VerticalAlignValue styleValue, object value)
        {
            Style style = new Style();
            style.CurrentCellXf.VerticalAlign = styleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentCellXf.VerticalAlign);
        }

        [Theory(DisplayName = "Test of the 'Indent' value when writing and reading a CellXF style")]
        [InlineData(0, HorizontalAlignValue.Left, 0, "test")]
        [InlineData(0, HorizontalAlignValue.Right, 0, "test")]
        [InlineData(0, HorizontalAlignValue.Distributed, 0, "test")]
        [InlineData(0, HorizontalAlignValue.Center, 0, "test")]
        [InlineData(1, HorizontalAlignValue.Left, 1, 0.5f)]
        [InlineData(1, HorizontalAlignValue.Right, 1, 0.5f)]
        [InlineData(1, HorizontalAlignValue.Distributed, 1, 0.5f)]
        [InlineData(1, HorizontalAlignValue.Center, 0, 0.5f)]
        [InlineData(5, HorizontalAlignValue.Left, 5, true)]
        [InlineData(5, HorizontalAlignValue.Right, 5, true)]
        [InlineData(5, HorizontalAlignValue.Distributed, 5, true)]
        [InlineData(5, HorizontalAlignValue.Center, 0, true)]
        [InlineData(64, HorizontalAlignValue.Left, 64, 22)]
        [InlineData(64, HorizontalAlignValue.Right, 64, 22)]
        [InlineData(64, HorizontalAlignValue.Distributed, 64, 22)]
        [InlineData(64, HorizontalAlignValue.Center, 0, 22)]
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
