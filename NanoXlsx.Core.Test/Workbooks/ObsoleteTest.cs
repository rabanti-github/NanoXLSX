using NanoXLSX;
using NanoXLSX.Shared.Exceptions;
using NanoXLSX.Shared.Exceptions;
using NanoXLSX.Styles;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace NanoXLSX_Test.Workbooks
{
    /// <summary>
    /// Note: All tests of this class are just for code coverage. The tested functions will be removed in the future
    /// </summary>
    /// Ensure that these tests are executed sequentially, since static repository methods may be called 
    [Collection(nameof(SequentialCollection))]
    public class ObsoleteTest
    {

        [Fact(DisplayName = "Test of the AddStyle function (only for code coverage)")]
        public void AddStyleTest()
        {
            Workbook workbook = new Workbook();
            workbook.AddStyle(BasicStyles.Bold);
            Assert.True(StyleRepository.Instance.Styles.ContainsKey(BasicStyles.Bold.GetHashCode()));
        }


        [Theory(DisplayName = "Test of the AddStyleComponent function (only for code coverage)")]
        [InlineData("Border")]
        [InlineData("CellXf")]
        [InlineData("Fill")]
        [InlineData("Font")]
        [InlineData("NumberFormat")]
        public void AddStyleComponentTest(string type)
        {
            Workbook workbook = new Workbook();
            AbstractStyle style = null;
            switch (type)
            {
                case "Border":
                    style = new Border();
                    break;
                case "CellXf":
                    style = new CellXf();
                    break;
                case "Fill":
                    style = new Fill() { PatternFill = NanoXLSX.Shared.Enums.Styles.FillEnums.PatternValue.gray125 };
                    break;
                case "Font":
                    style = new Font();
                    break;
                case "NumberFormat":
                    style = new NumberFormat();
                    break;
            }
            Style baseStyle = BasicStyles.DottedFill_0_125;
            workbook.AddStyleComponent(baseStyle, style);
            Assert.True(StyleRepository.Instance.Styles.ContainsKey(BasicStyles.DottedFill_0_125.GetHashCode()));
        }

        [Fact(DisplayName = "Test of the RemoveStyle function with an object (only for code coverage)")]
        public void RemoveStyleTest()
        {
            Workbook workbook = new Workbook();
            Style style = BasicStyles.Bold;
            workbook.AddStyle(style);
            workbook.RemoveStyle(style);
            workbook.RemoveStyle(style, false);
            workbook.RemoveStyle(style, false);
            Assert.True(StyleRepository.Instance.Styles.ContainsKey(BasicStyles.Bold.GetHashCode())); // This is expected
            Style style2 = null;
            Assert.Throws<StyleException>(() => workbook.RemoveStyle(style2));
        }

        [Fact(DisplayName = "Test of the RemoveStyle function with a name (only for code coverage)")]
        public void RemoveStyleTest2()
        {
            Workbook workbook = new Workbook();
            Style style = BasicStyles.Bold;
            workbook.AddStyle(style);
            workbook.RemoveStyle(style.Name);
            workbook.RemoveStyle(style.Name, true);
            workbook.RemoveStyle(style.Name, false);
            Assert.True(StyleRepository.Instance.Styles.ContainsKey(BasicStyles.Bold.GetHashCode())); // This is expected
            string styleName = null;
            Assert.Throws<StyleException>(() => workbook.RemoveStyle(styleName));
        }

    }
}
