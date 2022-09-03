using NanoXLSX.Exceptions;
using NanoXLSX.Styles;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace NanoXLSX_Test.Styles
{
    public class FillTest
    {

        private Fill exampleStyle;
        private Fill comparisonStyle;

        public FillTest()
        {
            exampleStyle = new Fill();
            exampleStyle.BackgroundColor = "FFAABB00";
            exampleStyle.ForegroundColor = "1188FF00";
            exampleStyle.IndexedColor = 99;
            exampleStyle.PatternFill = Fill.PatternValue.darkGray;

            comparisonStyle = new Fill();
            exampleStyle.BackgroundColor = "77CCBB00";
            exampleStyle.ForegroundColor = "DD33CC00";
            exampleStyle.IndexedColor = 32;
            exampleStyle.PatternFill = Fill.PatternValue.lightGray;
        }


        [Fact(DisplayName = "Test of the default values")]
        public void DefaultValuesTest()
        {
            Assert.Equal("FF000000", Fill.DEFAULT_COLOR);
            Assert.Equal(64, Fill.DEFAULT_INDEXED_COLOR);
            Assert.Equal(Fill.PatternValue.none, Fill.DEFAULT_PATTERN_FILL);
        }


        [Fact(DisplayName = "Test of the constructor with colors")]
        public void ConstructorTest()
        {
            Fill fill = new Fill();
            Assert.Equal(Fill.DEFAULT_INDEXED_COLOR, fill.IndexedColor);
            Assert.Equal(Fill.DEFAULT_PATTERN_FILL, fill.PatternFill);
            Assert.Equal(Fill.DEFAULT_COLOR, fill.ForegroundColor);
            Assert.Equal(Fill.DEFAULT_COLOR, fill.BackgroundColor);
        }

        [Fact(DisplayName = "Test of the constructor")]
        public void ConstructorTest2()
        {
            Fill fill = new Fill("FFAABBCC", "FF001122");
            Assert.Equal(Fill.DEFAULT_INDEXED_COLOR, fill.IndexedColor);
            Assert.Equal(Fill.PatternValue.solid, fill.PatternFill);
            Assert.Equal("FFAABBCC", fill.ForegroundColor);
            Assert.Equal("FF001122", fill.BackgroundColor);
        }


        [Theory(DisplayName = "Test of the constructor with color and fill type")]
        [InlineData("FFAABBCC", Fill.FillType.fillColor, "FFAABBCC", "FF000000")]
        [InlineData("FF112233", Fill.FillType.patternColor, "FF000000", "FF112233")]
        public void ConstructorTest3(string color, Fill.FillType fillType, string expectedForeground, string expectedBackground)
        {
            Fill fill = new Fill(color, fillType);
            Assert.Equal(Fill.DEFAULT_INDEXED_COLOR, fill.IndexedColor);
            Assert.Equal(Fill.PatternValue.solid, fill.PatternFill);
            Assert.Equal(expectedForeground, fill.ForegroundColor);
            Assert.Equal(expectedBackground, fill.BackgroundColor);
        }

        [Theory(DisplayName = "Test of the failing constructor")]
        [InlineData("", "FF000000")]
        [InlineData("FF000000", "")]
        [InlineData(null, "FF000000")]
        [InlineData("FF000000", null)]
        [InlineData("", "")]
        [InlineData(null, null)]
        [InlineData("FF00000000", "FFAABBCC")]
        [InlineData("FF000000", "FFAABBCCCC")]
        [InlineData("FF0000", "FFAABBCC")]
        [InlineData("FF000000", "FFAABB")]
        [InlineData("x", "FFAABBCC")]
        [InlineData("FF000000", "x")]
        [InlineData("x", "y")]
        public void ConstructorFailTest(string foreground, string background)
        {
            Assert.Throws<StyleException>(() => new Fill(foreground, background));
        }


        [Theory(DisplayName = "Test of the failing constructor with color and fill type")]
        [InlineData("", Fill.FillType.fillColor)]
        [InlineData(null, Fill.FillType.fillColor)]
        [InlineData("x", Fill.FillType.fillColor)]
        [InlineData("FFAABBCCDD", Fill.FillType.fillColor)]
        [InlineData("FFAABB", Fill.FillType.fillColor)]
        [InlineData("", Fill.FillType.patternColor)]
        [InlineData(null, Fill.FillType.patternColor)]
        [InlineData("x", Fill.FillType.patternColor)]
        [InlineData("FFAABBCCDD", Fill.FillType.patternColor)]
        [InlineData("FFAABB", Fill.FillType.patternColor)]
        public void ConstructorFailTest2(string color, Fill.FillType fillType)
        {
            Assert.Throws<StyleException>(() => new Fill(color, fillType));
        }



        [Theory(DisplayName = "Test of the get and set function of the BackgroundColor property")]
        [InlineData("77CCBB00")]
        [InlineData("00000000")]
        public void BackgroundColorTest(string value)
        {
            Fill fill = new Fill();
            Assert.Equal(Fill.DEFAULT_COLOR, fill.BackgroundColor);
            fill.BackgroundColor = value;
            Assert.Equal(value, fill.BackgroundColor);
        }

        [Theory(DisplayName = "Test of the failing set function of the BackgroundColor property with invalid values")]
        [InlineData("77BB00")]
        [InlineData("0002200000")]
        [InlineData("")]
        [InlineData(null)]
        [InlineData("XXXXXXXX")]
        public void BackgroundColorFailTest(string value)
        {
            Fill fill = new Fill();
            Exception ex = Assert.Throws<StyleException>(() => fill.BackgroundColor = value);
            Assert.Equal(typeof(StyleException), ex.GetType());
        }

        [Theory(DisplayName = "Test of the get and set function of the ForegroundColor property")]
        [InlineData("77CCBB00")]
        [InlineData("FFFFFFFF")]
        public void ForegroundColorTest(string value)
        {
            Fill fill = new Fill();
            Assert.Equal(Fill.DEFAULT_COLOR, fill.ForegroundColor);
            fill.ForegroundColor = value;
            Assert.Equal(value, fill.ForegroundColor);
        }

        [Theory(DisplayName = "Test of the failing set function of the ForegroundColor property with invalid values")]
        [InlineData("77BB00")]
        [InlineData("0002200000")]
        [InlineData("")]
        [InlineData(null)]
        [InlineData("XXXXXXXX")]
        public void ForegroundColorFailTest(string value)
        {
            Fill fill = new Fill();
            Exception ex = Assert.Throws<StyleException>(() => fill.ForegroundColor = value);
            Assert.Equal(typeof(StyleException), ex.GetType());
        }

        [Theory(DisplayName = "Test of the get and set function of the IndexedColor property")]
        [InlineData(0)]
        [InlineData(256)]
        [InlineData(-10)]
        public void IndexedColorTest(int value)
        {
            Fill fill = new Fill();
            Assert.Equal(Fill.DEFAULT_INDEXED_COLOR, fill.IndexedColor); // 64 is default
            fill.IndexedColor = value;
            Assert.Equal(value, fill.IndexedColor);
        }

        [Theory(DisplayName = "Test of the get and set function of the PatternFill property")]
        [InlineData(Fill.PatternValue.darkGray)]
        [InlineData(Fill.PatternValue.gray0625)]
        [InlineData(Fill.PatternValue.gray125)]
        [InlineData(Fill.PatternValue.lightGray)]
        [InlineData(Fill.PatternValue.mediumGray)]
        [InlineData(Fill.PatternValue.none)]
        [InlineData(Fill.PatternValue.solid)]
        public void PatternFillTest(Fill.PatternValue value)
        {
            Fill fill = new Fill();
            Assert.Equal(Fill.DEFAULT_PATTERN_FILL, fill.PatternFill); // default is none
            fill.PatternFill = value;
            Assert.Equal(value, fill.PatternFill);
        }

        [Theory(DisplayName = "Test of the SetColor function")]
        [InlineData("FFAABBCC", Fill.FillType.fillColor, "FFAABBCC", "FF000000")]
        [InlineData("FF112233", Fill.FillType.patternColor, "FF000000", "FF112233")]
        public void SetColorTest(string color, Fill.FillType fillType, string expectedForeground, string expectedBackground)
        {
            Fill fill = new Fill();
            Assert.Equal(Fill.DEFAULT_COLOR, fill.ForegroundColor);
            Assert.Equal(Fill.DEFAULT_COLOR, fill.BackgroundColor);
            Assert.Equal(Fill.PatternValue.none, fill.PatternFill);
            fill.SetColor(color, fillType);
            Assert.Equal(Fill.DEFAULT_INDEXED_COLOR, fill.IndexedColor);
            Assert.Equal(Fill.PatternValue.solid, fill.PatternFill);
            Assert.Equal(expectedForeground, fill.ForegroundColor);
            Assert.Equal(expectedBackground, fill.BackgroundColor);
        }

        [Theory(DisplayName = "Test of the ValidateColor function")]
        [InlineData("", false, false, false)]
        [InlineData(null, false, false, false)]
        [InlineData("", true, false, false)]
        [InlineData(null, true, false, false)]
        [InlineData("", false, true, true)]
        [InlineData(null, false, true, true)]
        [InlineData("", true, true, true)]
        [InlineData(null, true, true, true)]
        [InlineData("FFAABBCC", false, false, false)]
        [InlineData("FFAABBCC", true, false, true)]
        [InlineData("FFAABBCC", false, true, false)]
        [InlineData("FFAABBCC", true, true, true)]
        [InlineData("FFAABB", false, false, true)]
        [InlineData("FFAABB", true, false, false)]
        [InlineData("FFAA", true, false, false)]
        [InlineData("FFAA", false, false, false)]
        [InlineData("FFAA", true, true, false)]
        [InlineData("FFAACCDDDD", true, false, false)]
        [InlineData("FFAACCDDDD", false, false, false)]
        [InlineData("FFAACCDDDD", true, true, false)]
        public void ValidateColorTest(string color, bool useAlpha, bool allowEmpty, bool expectedValid)
        {
            if (expectedValid)
            {
                // Should not throw
                Fill.ValidateColor(color, useAlpha, allowEmpty);
            }
            else
            {
                Assert.Throws<StyleException>(() => Fill.ValidateColor(color, useAlpha, allowEmpty));
            }
            
        }

        [Fact(DisplayName = "Test of the CopyFill function")]
        public void CopyFillTest()
        {
            Fill copy = exampleStyle.CopyFill();
            Assert.Equal(exampleStyle.GetHashCode(), copy.GetHashCode());
        }

        [Fact(DisplayName = "Test of the Equals method")]
        public void EqualsTest()
        {
            Fill style2 = (Fill)exampleStyle.Copy();
            Assert.True(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of BackgroundColor)")]
        public void EqualsTest2a()
        {
            Fill style2 = (Fill)exampleStyle.Copy();
            style2.BackgroundColor = "66880000";
            Assert.False(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of ForegroundColor)")]
        public void EqualsTest2b()
        {
            Fill style2 = (Fill)exampleStyle.Copy();
            style2.ForegroundColor = "AA330000";
            Assert.False(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of IndexedColor)")]
        public void EqualsTest2c()
        {
            Fill style2 = (Fill)exampleStyle.Copy();
            style2.IndexedColor = 78;
            Assert.False(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of PatternFill)")]
        public void EqualsTest2d()
        {
            Fill style2 = (Fill)exampleStyle.Copy();
            style2.PatternFill = Fill.PatternValue.solid;
            Assert.False(exampleStyle.Equals(style2));
        }

        [Theory(DisplayName = "Test of the Equals method (inequality on null or different objects)")]
        [InlineData(null)]
        [InlineData("text")]
        [InlineData(true)]
        public void EqualsTest3(object obj)
        {
            Assert.False(exampleStyle.Equals(obj));
        }

        [Theory(DisplayName = "Test of the Equals method when the origin object is null or not of the same type")]
        [InlineData(null)]
        [InlineData(true)]
        [InlineData("origin")]
        public void EqualsTest5(object origin)
        {
            Fill copy = (Fill)exampleStyle.Copy();
            Assert.False(copy.Equals(origin));
        }

        [Fact(DisplayName = "Test of the GetHashCode method (equality of two identical objects)")]
        public void GetHashCodeTest()
        {
            Fill copy = (Fill)exampleStyle.Copy();
            copy.InternalID = 99;  // Should not influence
            Assert.Equal(exampleStyle.GetHashCode(), copy.GetHashCode());
            Assert.Equal(exampleStyle.GetHashCode(), copy.GetHashCode()); // For code coverage
        }

        [Fact(DisplayName = "Test of the GetHashCode method (inequality of two different objects)")]
        public void GetHashCodeTest2()
        {
            Fill copy = (Fill)exampleStyle.Copy();
            copy.BackgroundColor = "778800FF";
            Assert.NotEqual(exampleStyle.GetHashCode(), copy.GetHashCode());
            Assert.NotEqual(exampleStyle.GetHashCode(), copy.GetHashCode()); // For code coverage
        }

        [Fact(DisplayName = "Test of the CompareTo method")]
        public void CompareToTest()
        {
            Fill fill = new Fill();
            Fill other = new Fill();
            fill.InternalID = null;
            other.InternalID = null;
            Assert.Equal(-1, fill.CompareTo(other));
            fill.InternalID = 5;
            Assert.Equal(1, fill.CompareTo(other));
            Assert.Equal(1, fill.CompareTo(null));
            other.InternalID = 5;
            Assert.Equal(0, fill.CompareTo(other));
            other.InternalID = 4;
            Assert.Equal(1, fill.CompareTo(other));
            other.InternalID = 6;
            Assert.Equal(-1, fill.CompareTo(other));
        }

        // For code coverage
        [Fact(DisplayName = "Test of the ToString function")]
        public void ToStringTest()
        {
            Fill fill = new Fill();
            string s1 = fill.ToString();
            fill.ForegroundColor = "FFAABBCC";
            Assert.NotEqual(s1, fill.ToString()); // An explicit value comparison is probably not sensible
        }

    }
}
