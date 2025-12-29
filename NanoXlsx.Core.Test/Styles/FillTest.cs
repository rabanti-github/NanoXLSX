using System;
using NanoXLSX.Colors;
using NanoXLSX.Exceptions;
using NanoXLSX.Styles;
using NanoXLSX.Utils;
using Xunit;
using static NanoXLSX.Styles.Fill;

namespace NanoXLSX.Test.Core.StyleTest
{
    public class FillTest
    {

        private Fill exampleStyle;
        private Fill comparisonStyle;

        public FillTest()
        {
            exampleStyle = new Fill
            {
                BackgroundColor = "FFAABB00",
                ForegroundColor = Color.CreateIndexed(IndexedColor.Value.BrightGreen),
                PatternFill = PatternValue.DarkGray
            };

            comparisonStyle = new Fill();
            exampleStyle.BackgroundColor = "77CCBB00";
            exampleStyle.ForegroundColor = Color.CreateIndexed(IndexedColor.Value.Blue4);
            exampleStyle.PatternFill = PatternValue.LightGray;
        }


        [Fact(DisplayName = "Test of the default values")]
        public void DefaultValuesTest()
        {
            Assert.Equal("FF000000", Fill.DefaultColor);
            Assert.Equal(IndexedColor.DefaultIndexedColor, Fill.DefaultIndexedColor.IndexedColor);
            Assert.Equal(PatternValue.None, Fill.DefaultPatternFill);
        }


        [Fact(DisplayName = "Test of the constructor with colors")]
        public void ConstructorTest()
        {
            Fill fill = new Fill();
            Assert.Equal(Fill.DefaultPatternFill, fill.PatternFill);
            Assert.Equal(Fill.DefaultColor, fill.ForegroundColor);
            Assert.Equal(Fill.DefaultColor, fill.BackgroundColor);
        }

        [Fact(DisplayName = "Test of the constructor")]
        public void ConstructorTest2()
        {
            Fill fill = new Fill("FFAABBCC", "FF001122");
            Assert.Equal(PatternValue.Solid, fill.PatternFill);
            Assert.Equal("FFAABBCC", fill.ForegroundColor);
            Assert.Equal("FF001122", fill.BackgroundColor);
        }


        [Theory(DisplayName = "Test of the constructor with color and fill type")]
        [InlineData("FFAABBCC", FillType.FillColor, "FFAABBCC", "FF000000")]
        [InlineData("FF112233", FillType.PatternColor, "FF000000", "FF112233")]
        public void ConstructorTest3(string color, FillType fillType, string expectedForeground, string expectedBackground)
        {
            Fill fill = new Fill(color, fillType);
            Assert.Equal(PatternValue.Solid, fill.PatternFill);
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
        [InlineData("FF0000", "FAABBCC")]
        [InlineData("FF0000", "#FFAABBCC")]
        [InlineData("F000000", "FFAABB")]
        [InlineData("#FF000000", "FFAABB")]
        [InlineData("x", "FFAABBCC")]
        [InlineData("FF000000", "x")]
        [InlineData("x", "y")]
        public void ConstructorFailTest(string foreground, string background)
        {
            Assert.Throws<StyleException>(() => new Fill(foreground, background));
        }


        [Theory(DisplayName = "Test of the failing constructor with color and fill type")]
        [InlineData("", FillType.FillColor)]
        [InlineData(null, FillType.FillColor)]
        [InlineData("x", FillType.FillColor)]
        [InlineData("FFAABBCCDD", FillType.FillColor)]
        [InlineData("FAABB", FillType.FillColor)]
        [InlineData("#FFAABB", FillType.FillColor)]
        [InlineData("", FillType.PatternColor)]
        [InlineData(null, FillType.PatternColor)]
        [InlineData("x", FillType.PatternColor)]
        [InlineData("FFAABBCCDD", FillType.PatternColor)]
        [InlineData("FAABB", FillType.PatternColor)]
        [InlineData("#FFAABB", FillType.PatternColor)]
        public void ConstructorFailTest2(string color, FillType fillType)
        {
            Assert.Throws<StyleException>(() => new Fill(color, fillType));
        }



        [Theory(DisplayName = "Test of the get and set function of the BackgroundColor property")]
        [InlineData("77CCBB00")]
        [InlineData("00000000")]
        public void BackgroundColorTest(string value)
        {
            Fill fill = new Fill();
            Assert.Equal(Fill.DefaultColor, fill.BackgroundColor);
            fill.BackgroundColor = value;
            Assert.Equal(value, fill.BackgroundColor);
        }

        [Theory(DisplayName = "Test of the failing set function of the BackgroundColor property with invalid values")]
        [InlineData("7BB00")]
        [InlineData("#77BB00")]
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
            Assert.Equal(Fill.DefaultColor, fill.ForegroundColor);
            fill.ForegroundColor = value;
            Assert.Equal(value, fill.ForegroundColor);
        }

        [Theory(DisplayName = "Test of the failing set function of the ForegroundColor property with invalid values")]
        [InlineData("7BB00")]
        [InlineData("#77BB00")]
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
        [InlineData(IndexedColor.Value.Black0)]
        [InlineData(IndexedColor.Value.White1)]
        [InlineData(IndexedColor.Value.Red2)]
        [InlineData(IndexedColor.Value.BrightGreen3)]
        [InlineData(IndexedColor.Value.Blue4)]
        [InlineData(IndexedColor.Value.Yellow5)]
        [InlineData(IndexedColor.Value.Magenta6)]
        [InlineData(IndexedColor.Value.Cyan7)]
        [InlineData(IndexedColor.Value.Black)]
        [InlineData(IndexedColor.Value.White)]
        [InlineData(IndexedColor.Value.Red)]
        [InlineData(IndexedColor.Value.BrightGreen)]
        [InlineData(IndexedColor.Value.Blue)]
        [InlineData(IndexedColor.Value.Yellow)]
        [InlineData(IndexedColor.Value.Magenta)]
        [InlineData(IndexedColor.Value.Cyan)]
        [InlineData(IndexedColor.Value.DarkRed)]
        [InlineData(IndexedColor.Value.DarkGreen)]
        [InlineData(IndexedColor.Value.DarkBlue)]
        [InlineData(IndexedColor.Value.Olive)]
        [InlineData(IndexedColor.Value.Purple)]
        [InlineData(IndexedColor.Value.Teal)]
        [InlineData(IndexedColor.Value.LightGray)]
        [InlineData(IndexedColor.Value.Gray)]
        [InlineData(IndexedColor.Value.LightCornflowerBlue)]
        [InlineData(IndexedColor.Value.DarkRose)]
        [InlineData(IndexedColor.Value.LightYellow)]
        [InlineData(IndexedColor.Value.LightCyan)]
        [InlineData(IndexedColor.Value.DarkPurple)]
        [InlineData(IndexedColor.Value.Salmon)]
        [InlineData(IndexedColor.Value.MediumBlue)]
        [InlineData(IndexedColor.Value.LightLavender)]
        [InlineData(IndexedColor.Value.Navy)]
        [InlineData(IndexedColor.Value.StrongMagenta)]
        [InlineData(IndexedColor.Value.StrongYellow)]
        [InlineData(IndexedColor.Value.StrongCyan)]
        [InlineData(IndexedColor.Value.DarkViolet)]
        [InlineData(IndexedColor.Value.DarkMaroon)]
        [InlineData(IndexedColor.Value.DarkTeal)]
        [InlineData(IndexedColor.Value.PureBlue)]
        [InlineData(IndexedColor.Value.SkyBlue)]
        [InlineData(IndexedColor.Value.PaleCyan)]
        [InlineData(IndexedColor.Value.LightMint)]
        [InlineData(IndexedColor.Value.PastelYellow)]
        [InlineData(IndexedColor.Value.LightSkyBlue)]
        [InlineData(IndexedColor.Value.Rose)]
        [InlineData(IndexedColor.Value.Lavender)]
        [InlineData(IndexedColor.Value.Peach)]
        [InlineData(IndexedColor.Value.RoyalBlue)]
        [InlineData(IndexedColor.Value.Turquoise)]
        [InlineData(IndexedColor.Value.LightOlive)]
        [InlineData(IndexedColor.Value.Gold)]
        [InlineData(IndexedColor.Value.Orange)]
        [InlineData(IndexedColor.Value.DarkOrange)]
        [InlineData(IndexedColor.Value.BlueGray)]
        [InlineData(IndexedColor.Value.MediumGray)]
        [InlineData(IndexedColor.Value.DarkSlateBlue)]
        [InlineData(IndexedColor.Value.SeaGreen)]
        [InlineData(IndexedColor.Value.VeryDarkGreen)]
        [InlineData(IndexedColor.Value.DarkOlive)]
        [InlineData(IndexedColor.Value.Brown)]
        [InlineData(IndexedColor.Value.DarkRoseDuplicate)]
        [InlineData(IndexedColor.Value.Indigo)]
        [InlineData(IndexedColor.Value.VeryDarkGray)]
        [InlineData(IndexedColor.Value.SystemForeground)]
        [InlineData(IndexedColor.Value.SystemBackground)]
        public void IndexedColorTest(IndexedColor.Value value)
        {
            Fill fill = new Fill();
            Assert.Null(fill.ForegroundColor.IndexedColor);
            Assert.Equal(Color.ColorType.Rgb, fill.ForegroundColor.Type); // RGB is default
            fill.ForegroundColor = Color.CreateIndexed(value);
            Assert.Equal(value, fill.ForegroundColor.IndexedColor);
            Assert.Equal(Color.ColorType.Indexed, fill.ForegroundColor.Type);

            fill = new Fill();
            Assert.Null(fill.BackgroundColor.IndexedColor);
            fill.BackgroundColor = Color.CreateIndexed(value);
            Assert.Equal(value, fill.BackgroundColor.IndexedColor);
            Assert.Equal(Color.ColorType.Indexed, fill.BackgroundColor.Type);
        }

        [Theory(DisplayName = "Test of the get and set function of the PatternFill property")]
        [InlineData(PatternValue.DarkGray)]
        [InlineData(PatternValue.Gray0625)]
        [InlineData(PatternValue.Gray125)]
        [InlineData(PatternValue.LightGray)]
        [InlineData(PatternValue.MediumGray)]
        [InlineData(PatternValue.None)]
        [InlineData(PatternValue.Solid)]
        public void PatternFillTest(PatternValue value)
        {
            Fill fill = new Fill();
            Assert.Equal(Fill.DefaultPatternFill, fill.PatternFill); // default is none
            fill.PatternFill = value;
            Assert.Equal(value, fill.PatternFill);
        }

        [Theory(DisplayName = "Test of the SetColor function")]
        [InlineData("FFAABBCC", FillType.FillColor, "FFAABBCC", "FF000000")]
        [InlineData("FF112233", FillType.PatternColor, "FF000000", "FF112233")]
        public void SetColorTest(string color, FillType fillType, string expectedForeground, string expectedBackground)
        {
            Fill fill = new Fill();
            Assert.Equal(Fill.DefaultColor, fill.ForegroundColor);
            Assert.Equal(Fill.DefaultColor, fill.BackgroundColor);
            Assert.Equal(PatternValue.None, fill.PatternFill);
            fill.SetColor(color, fillType);
           Assert.Equal(Color.ColorType.Rgb, fill.ForegroundColor.Type);
            Assert.Equal(Color.ColorType.Rgb, fill.BackgroundColor.Type);
            Assert.Equal(PatternValue.Solid, fill.PatternFill);
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
                Validators.ValidateColor(color, useAlpha, allowEmpty);
            }
            else
            {
                Assert.Throws<StyleException>(() => Validators.ValidateColor(color, useAlpha, allowEmpty));
            }

        }

        [Fact(DisplayName = "Test of the CopyFill function")]
        public void CopyFillTest()
        {
            Fill copy = exampleStyle.CopyFill();
            Assert.Equal(exampleStyle.GetHashCode(), copy.GetHashCode());
        }

        [Theory(DisplayName = "Test of the implicit operator (create ARGB color by string)")]
        [InlineData("FF000000")]
        [InlineData("AC000000")]
        [InlineData("FF0000FF")]
        [InlineData("FFFFFFFF")]
        [InlineData("FF123456")]
        public void ImplicitOperatorTest(string value)
        {
            Fill fill = value;
            Fill expectedFill = new Fill(value, FillType.FillColor);
            Assert.Equal(fill.ForegroundColor, expectedFill.ForegroundColor);
            Assert.Equal(fill.BackgroundColor, expectedFill.BackgroundColor);
        }

        [Theory(DisplayName = "Test of the implicit operator (create indexed color by int)")]
        [InlineData(10, IndexedColor.Value.Red)]
        [InlineData(0, IndexedColor.Value.Black0)]
        [InlineData(8, IndexedColor.Value.Black)]
        [InlineData(64, IndexedColor.Value.SystemForeground)]
        [InlineData(65, IndexedColor.Value.SystemBackground)]
        public void ImplicitOperatorTest3(int givenValue, IndexedColor.Value expectedValue)
        {
            Fill fill = givenValue;
            Fill expectedFill = new Fill();
            expectedFill.ForegroundColor = Color.CreateIndexed(expectedValue);
            expectedFill.PatternFill = PatternValue.Solid;
            Assert.Equal(fill.ForegroundColor, expectedFill.ForegroundColor);
            Assert.Equal(fill.BackgroundColor, expectedFill.BackgroundColor);
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
            style2.ForegroundColor = Color.CreateIndexed(IndexedColor.Value.Olive);
            Assert.False(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of PatternFill)")]
        public void EqualsTest2d()
        {
            Fill style2 = (Fill)exampleStyle.Copy();
            style2.PatternFill = PatternValue.Solid;
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
