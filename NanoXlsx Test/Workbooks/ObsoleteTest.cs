using NanoXLSX;
using NanoXLSX.Styles;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace NanoXLSX_Test.Wprkbooks
{
    /// <summary>
    /// Note: All tests of this class are just for code coverage. The tested functions will be removed in the future
    /// </summary>
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
                    style = new Fill();
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


    }
}
