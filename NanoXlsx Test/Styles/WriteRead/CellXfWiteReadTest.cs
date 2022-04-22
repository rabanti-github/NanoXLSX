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
    }
}
