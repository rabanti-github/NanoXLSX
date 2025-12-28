using NanoXLSX.Colors;
using NanoXLSX.Exceptions;
using Xunit;

namespace NanoXLSX.Core.Test.Colors
{
    public class IndexedColorTest
    {
        [Theory(DisplayName = "Test of the getter and setter of the ColorValue property on valid values")]
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
        public void ColorValueTest(IndexedColor.Value value)
        {
            var color = new IndexedColor();
            Assert.Equal(IndexedColor.DefaultIndexedColor, color.ColorValue); // Default is 64 (SystemForeground)
            color.ColorValue = value;
            Assert.Equal(value, color.ColorValue);
        }

        [Theory(DisplayName = "Test of the getter of the StringValue property")]
        [InlineData(IndexedColor.Value.Black0, "0")]
        [InlineData(IndexedColor.Value.White1, "1")]
        [InlineData(IndexedColor.Value.Red2, "2")]
        [InlineData(IndexedColor.Value.BrightGreen3, "3")]
        [InlineData(IndexedColor.Value.Blue4, "4")]
        [InlineData(IndexedColor.Value.Yellow5, "5")]
        [InlineData(IndexedColor.Value.Magenta6, "6")]
        [InlineData(IndexedColor.Value.Cyan7, "7")]
        [InlineData(IndexedColor.Value.Black, "8")]
        [InlineData(IndexedColor.Value.White, "9")]
        [InlineData(IndexedColor.Value.Red, "10")]
        [InlineData(IndexedColor.Value.BrightGreen, "11")]
        [InlineData(IndexedColor.Value.Blue, "12")]
        [InlineData(IndexedColor.Value.Yellow, "13")]
        [InlineData(IndexedColor.Value.Magenta, "14")]
        [InlineData(IndexedColor.Value.Cyan, "15")]
        [InlineData(IndexedColor.Value.DarkRed, "16")]
        [InlineData(IndexedColor.Value.DarkGreen, "17")]
        [InlineData(IndexedColor.Value.DarkBlue, "18")]
        [InlineData(IndexedColor.Value.Olive, "19")]
        [InlineData(IndexedColor.Value.Purple, "20")]
        [InlineData(IndexedColor.Value.Teal, "21")]
        [InlineData(IndexedColor.Value.LightGray, "22")]
        [InlineData(IndexedColor.Value.Gray, "23")]
        [InlineData(IndexedColor.Value.LightCornflowerBlue, "24")]
        [InlineData(IndexedColor.Value.DarkRose, "25")]
        [InlineData(IndexedColor.Value.LightYellow, "26")]
        [InlineData(IndexedColor.Value.LightCyan, "27")]
        [InlineData(IndexedColor.Value.DarkPurple, "28")]
        [InlineData(IndexedColor.Value.Salmon, "29")]
        [InlineData(IndexedColor.Value.MediumBlue, "30")]
        [InlineData(IndexedColor.Value.LightLavender, "31")]
        [InlineData(IndexedColor.Value.Navy, "32")]
        [InlineData(IndexedColor.Value.StrongMagenta, "33")]
        [InlineData(IndexedColor.Value.StrongYellow, "34")]
        [InlineData(IndexedColor.Value.StrongCyan, "35")]
        [InlineData(IndexedColor.Value.DarkViolet, "36")]
        [InlineData(IndexedColor.Value.DarkMaroon, "37")]
        [InlineData(IndexedColor.Value.DarkTeal, "38")]
        [InlineData(IndexedColor.Value.PureBlue, "39")]
        [InlineData(IndexedColor.Value.SkyBlue, "40")]
        [InlineData(IndexedColor.Value.PaleCyan, "41")]
        [InlineData(IndexedColor.Value.LightMint, "42")]
        [InlineData(IndexedColor.Value.PastelYellow, "43")]
        [InlineData(IndexedColor.Value.LightSkyBlue, "44")]
        [InlineData(IndexedColor.Value.Rose, "45")]
        [InlineData(IndexedColor.Value.Lavender, "46")]
        [InlineData(IndexedColor.Value.Peach, "47")]
        [InlineData(IndexedColor.Value.RoyalBlue, "48")]
        [InlineData(IndexedColor.Value.Turquoise, "49")]
        [InlineData(IndexedColor.Value.LightOlive, "50")]
        [InlineData(IndexedColor.Value.Gold, "51")]
        [InlineData(IndexedColor.Value.Orange, "52")]
        [InlineData(IndexedColor.Value.DarkOrange, "53")]
        [InlineData(IndexedColor.Value.BlueGray, "54")]
        [InlineData(IndexedColor.Value.MediumGray, "55")]
        [InlineData(IndexedColor.Value.DarkSlateBlue, "56")]
        [InlineData(IndexedColor.Value.SeaGreen, "57")]
        [InlineData(IndexedColor.Value.VeryDarkGreen, "58")]
        [InlineData(IndexedColor.Value.DarkOlive, "59")]
        [InlineData(IndexedColor.Value.Brown, "60")]
        [InlineData(IndexedColor.Value.DarkRoseDuplicate, "61")]
        [InlineData(IndexedColor.Value.Indigo, "62")]
        [InlineData(IndexedColor.Value.VeryDarkGray, "63")]
        [InlineData(IndexedColor.Value.SystemForeground, "64")]
        [InlineData(IndexedColor.Value.SystemBackground, "65")]
        public void StringValueTest(IndexedColor.Value givenValue, string expectedValue)
        {
            var color = new IndexedColor();
            Assert.Equal(IndexedColor.DefaultIndexedColor, color.ColorValue); // Default
            color.ColorValue = givenValue;
            Assert.Equal(expectedValue, color.StringValue);
        }

        [Fact(DisplayName = "Test of the default Constructor")]
        public void ConstructorTest()
        {
            var color = new IndexedColor();
            Assert.Equal(IndexedColor.DefaultIndexedColor, color.ColorValue); // Default
        }

        [Theory(DisplayName = "Test of the Constructor with an enum value")]
        [InlineData(IndexedColor.Value.Black0)]
        [InlineData(IndexedColor.Value.White1)]
        [InlineData(IndexedColor.Value.Red2)]
        [InlineData(IndexedColor.Value.BrightGreen3)]
        [InlineData(IndexedColor.Value.SystemBackground)]
        [InlineData(IndexedColor.Value.SystemForeground)]
        public void ConstructorTest2(IndexedColor.Value value)
        {
            var color = new IndexedColor(value);
            Assert.Equal(value, color.ColorValue);
        }

        [Theory(DisplayName = "Test of the Constructor with an index")]
        [InlineData(0, IndexedColor.Value.Black0)]
        [InlineData(1, IndexedColor.Value.White1)]
        [InlineData(2, IndexedColor.Value.Red2)]
        [InlineData(3, IndexedColor.Value.BrightGreen3)]
        [InlineData(4, IndexedColor.Value.Blue4)]
        [InlineData(5, IndexedColor.Value.Yellow5)]
        [InlineData(6, IndexedColor.Value.Magenta6)]
        [InlineData(7, IndexedColor.Value.Cyan7)]
        [InlineData(8, IndexedColor.Value.Black)]
        [InlineData(9, IndexedColor.Value.White)]
        [InlineData(10, IndexedColor.Value.Red)]
        [InlineData(11, IndexedColor.Value.BrightGreen)]
        [InlineData(12, IndexedColor.Value.Blue)]
        [InlineData(13, IndexedColor.Value.Yellow)]
        [InlineData(14, IndexedColor.Value.Magenta)]
        [InlineData(15, IndexedColor.Value.Cyan)]
        [InlineData(16, IndexedColor.Value.DarkRed)]
        [InlineData(17, IndexedColor.Value.DarkGreen)]
        [InlineData(18, IndexedColor.Value.DarkBlue)]
        [InlineData(19, IndexedColor.Value.Olive)]
        [InlineData(20, IndexedColor.Value.Purple)]
        [InlineData(21, IndexedColor.Value.Teal)]
        [InlineData(22, IndexedColor.Value.LightGray)]
        [InlineData(23, IndexedColor.Value.Gray)]
        [InlineData(24, IndexedColor.Value.LightCornflowerBlue)]
        [InlineData(25, IndexedColor.Value.DarkRose)]
        [InlineData(26, IndexedColor.Value.LightYellow)]
        [InlineData(27, IndexedColor.Value.LightCyan)]
        [InlineData(28, IndexedColor.Value.DarkPurple)]
        [InlineData(29, IndexedColor.Value.Salmon)]
        [InlineData(30, IndexedColor.Value.MediumBlue)]
        [InlineData(31, IndexedColor.Value.LightLavender)]
        [InlineData(32, IndexedColor.Value.Navy)]
        [InlineData(33, IndexedColor.Value.StrongMagenta)]
        [InlineData(34, IndexedColor.Value.StrongYellow)]
        [InlineData(35, IndexedColor.Value.StrongCyan)]
        [InlineData(36, IndexedColor.Value.DarkViolet)]
        [InlineData(37, IndexedColor.Value.DarkMaroon)]
        [InlineData(38, IndexedColor.Value.DarkTeal)]
        [InlineData(39, IndexedColor.Value.PureBlue)]
        [InlineData(40, IndexedColor.Value.SkyBlue)]
        [InlineData(41, IndexedColor.Value.PaleCyan)]
        [InlineData(42, IndexedColor.Value.LightMint)]
        [InlineData(43, IndexedColor.Value.PastelYellow)]
        [InlineData(44, IndexedColor.Value.LightSkyBlue)]
        [InlineData(45, IndexedColor.Value.Rose)]
        [InlineData(46, IndexedColor.Value.Lavender)]
        [InlineData(47, IndexedColor.Value.Peach)]
        [InlineData(48, IndexedColor.Value.RoyalBlue)]
        [InlineData(49, IndexedColor.Value.Turquoise)]
        [InlineData(50, IndexedColor.Value.LightOlive)]
        [InlineData(51, IndexedColor.Value.Gold)]
        [InlineData(52, IndexedColor.Value.Orange)]
        [InlineData(53, IndexedColor.Value.DarkOrange)]
        [InlineData(54, IndexedColor.Value.BlueGray)]
        [InlineData(55, IndexedColor.Value.MediumGray)]
        [InlineData(56, IndexedColor.Value.DarkSlateBlue)]
        [InlineData(57, IndexedColor.Value.SeaGreen)]
        [InlineData(58, IndexedColor.Value.VeryDarkGreen)]
        [InlineData(59, IndexedColor.Value.DarkOlive)]
        [InlineData(60, IndexedColor.Value.Brown)]
        [InlineData(61, IndexedColor.Value.DarkRoseDuplicate)]
        [InlineData(62, IndexedColor.Value.Indigo)]
        [InlineData(63, IndexedColor.Value.VeryDarkGray)]
        [InlineData(64, IndexedColor.Value.SystemForeground)]
        [InlineData(65, IndexedColor.Value.SystemBackground)]
        public void ConstructorTest3(int index, IndexedColor.Value expectedValue)
        {
            var color = new IndexedColor(index);
            Assert.Equal(expectedValue, color.ColorValue);
        }

        [Theory(DisplayName = "Test of the failing Constructor on invalid values")]
        [InlineData(-1)]
        [InlineData(66)]
        [InlineData(255)]
        [InlineData(-100)]
        public void ConstructorFailTest(int value)
        {
            Assert.Throws<StyleException>(() => { var color = new IndexedColor(value); });
        }

        [Theory(DisplayName = "Test of the GetSrgbColor method")]
        [InlineData(IndexedColor.Value.Black0, "FF000000")]
        [InlineData(IndexedColor.Value.White1, "FFFFFFFF")]
        [InlineData(IndexedColor.Value.Red2, "FFFF0000")]
        [InlineData(IndexedColor.Value.BrightGreen3, "FF00FF00")]
        [InlineData(IndexedColor.Value.Blue4, "FF0000FF")]
        [InlineData(IndexedColor.Value.Yellow5, "FFFFFF00")]
        [InlineData(IndexedColor.Value.Magenta6, "FFFF00FF")]
        [InlineData(IndexedColor.Value.Cyan7, "FF00FFFF")]
        [InlineData(IndexedColor.Value.Black, "FF000000")]
        [InlineData(IndexedColor.Value.White, "FFFFFFFF")]
        [InlineData(IndexedColor.Value.Red, "FFFF0000")]
        [InlineData(IndexedColor.Value.BrightGreen, "FF00FF00")]
        [InlineData(IndexedColor.Value.Blue, "FF0000FF")]
        [InlineData(IndexedColor.Value.PureBlue, "FF0000FF")]
        [InlineData(IndexedColor.Value.Yellow, "FFFFFF00")]
        [InlineData(IndexedColor.Value.StrongYellow, "FFFFFF00")]
        [InlineData(IndexedColor.Value.Magenta, "FFFF00FF")]
        [InlineData(IndexedColor.Value.StrongMagenta, "FFFF00FF")]
        [InlineData(IndexedColor.Value.Cyan, "FF00FFFF")]
        [InlineData(IndexedColor.Value.StrongCyan, "FF00FFFF")]
        [InlineData(IndexedColor.Value.DarkRed, "FF800000")]
        [InlineData(IndexedColor.Value.DarkMaroon, "FF800000")]
        [InlineData(IndexedColor.Value.DarkGreen, "FF008000")]
        [InlineData(IndexedColor.Value.DarkBlue, "FF000080")]
        [InlineData(IndexedColor.Value.Navy, "FF000080")]
        [InlineData(IndexedColor.Value.Olive, "FF808000")]
        [InlineData(IndexedColor.Value.Purple, "FF800080")]
        [InlineData(IndexedColor.Value.DarkViolet, "FF800080")]
        [InlineData(IndexedColor.Value.Teal, "FF008080")]
        [InlineData(IndexedColor.Value.DarkTeal, "FF008080")]
        [InlineData(IndexedColor.Value.LightGray, "FFC0C0C0")]
        [InlineData(IndexedColor.Value.Gray, "FF808080")]
        [InlineData(IndexedColor.Value.LightCornflowerBlue, "FF9999FF")]
        [InlineData(IndexedColor.Value.DarkRose, "FF993366")]
        [InlineData(IndexedColor.Value.DarkRoseDuplicate, "FF993366")]
        [InlineData(IndexedColor.Value.LightYellow, "FFFFFFCC")]
        [InlineData(IndexedColor.Value.LightCyan, "FFCCFFFF")]
        [InlineData(IndexedColor.Value.PaleCyan, "FFCCFFFF")]
        [InlineData(IndexedColor.Value.DarkPurple, "FF660066")]
        [InlineData(IndexedColor.Value.Salmon, "FFFF8080")]
        [InlineData(IndexedColor.Value.MediumBlue, "FF0066CC")]
        [InlineData(IndexedColor.Value.LightLavender, "FFCCCCFF")]
        [InlineData(IndexedColor.Value.SkyBlue, "FF00CCFF")]
        [InlineData(IndexedColor.Value.LightMint, "FFCCFFCC")]
        [InlineData(IndexedColor.Value.PastelYellow, "FFFFFF99")]
        [InlineData(IndexedColor.Value.LightSkyBlue, "FF99CCFF")]
        [InlineData(IndexedColor.Value.Rose, "FFFF99CC")]
        [InlineData(IndexedColor.Value.Lavender, "FFCC99FF")]
        [InlineData(IndexedColor.Value.Peach, "FFFFCC99")]
        [InlineData(IndexedColor.Value.RoyalBlue, "FF3366FF")]
        [InlineData(IndexedColor.Value.Turquoise, "FF33CCCC")]
        [InlineData(IndexedColor.Value.LightOlive, "FF99CC00")]
        [InlineData(IndexedColor.Value.Gold, "FFFFCC00")]
        [InlineData(IndexedColor.Value.Orange, "FFFF9900")]
        [InlineData(IndexedColor.Value.DarkOrange, "FFFF6600")]
        [InlineData(IndexedColor.Value.BlueGray, "FF666699")]
        [InlineData(IndexedColor.Value.MediumGray, "FF969696")]
        [InlineData(IndexedColor.Value.DarkSlateBlue, "FF003366")]
        [InlineData(IndexedColor.Value.SeaGreen, "FF339966")]
        [InlineData(IndexedColor.Value.VeryDarkGreen, "FF003300")]
        [InlineData(IndexedColor.Value.DarkOlive, "FF333300")]
        [InlineData(IndexedColor.Value.Brown, "FF993300")]
        [InlineData(IndexedColor.Value.Indigo, "FF333399")]
        [InlineData(IndexedColor.Value.VeryDarkGray, "FF333333")]
        [InlineData(IndexedColor.Value.SystemBackground, IndexedColor.DefaultSystemBackgroundColorArgb)]
        [InlineData(IndexedColor.Value.SystemForeground, IndexedColor.DefaultSystemForegroundColorArgb)]
        public void GetSrgbTest(IndexedColor.Value givenValue, string expectedArgbValue)
        {
            var color = new IndexedColor(givenValue);
            var rgb = color.GetSrgbColor();
            Assert.Equal(expectedArgbValue, rgb.ColorValue);
        }

        [Fact(DisplayName = "Test of the Equals method (multiple cases)")]
        public void EqualsTest()
        {
            var color1 = new IndexedColor(IndexedColor.Value.Red);
            var color2 = new IndexedColor(IndexedColor.Value.Red);
            Assert.True(color1.Equals(color2)); // Same value


            var color3 = new IndexedColor();
            var color4 = new IndexedColor();
            Assert.True(color3.Equals(color4)); // Default value

            var color5 = new IndexedColor(23);
            var color6 = new IndexedColor(23);
            Assert.True(color5.Equals(color6)); // Same index value

            var color1b = new IndexedColor(10); // equivalent to Red
            Assert.True(color1.Equals(color1b)); // Different construction, same value

            var colorDefault1 = new IndexedColor(IndexedColor.Value.SystemForeground);
            Assert.True(color3.Equals(colorDefault1)); // Default value by property or constructor
        }

        [Fact(DisplayName = "Test of the Equals method on inequality (multiple cases)")]
        public void EqualsTest2()
        {
            var color1 = new IndexedColor(IndexedColor.Value.BrightGreen);
            var color2 = new IndexedColor(IndexedColor.Value.Blue);
            Assert.False(color1.Equals(color2));

            var obj = new object();
            Assert.False(color1.Equals(obj));

            IndexedColor color3 = null;
            Assert.False(color1.Equals(color3));

            var color4 = new IndexedColor(55);
            Assert.False(color1.Equals(color4));
        }

        [Fact(DisplayName = "Test of the GetHashCode method (multiple cases)")]
        public void GetHashCodeTest()
        {
            var color1 = new IndexedColor(IndexedColor.Value.Yellow);
            var color2 = new IndexedColor(IndexedColor.Value.Yellow);
            Assert.Equal(color1.GetHashCode(), color2.GetHashCode());

            var color3 = new IndexedColor();
            var color4 = new IndexedColor();
            Assert.Equal(color3.GetHashCode(), color4.GetHashCode());

            var color5 = new IndexedColor(30);
            var color6 = new IndexedColor(30);
            Assert.Equal(color5.GetHashCode(), color6.GetHashCode());
        }

        [Fact(DisplayName = "Test of the GetHashCode method on inequality (multiple cases)")]
        public void GetHashCodeTest2()
        {
            var color1 = new IndexedColor(IndexedColor.Value.Magenta);
            var color2 = new IndexedColor(IndexedColor.Value.Cyan);
            Assert.NotEqual(color1.GetHashCode(), color2.GetHashCode());

            var color3 = new IndexedColor(12);
            var color4 = new IndexedColor(45);
            Assert.NotEqual(color3.GetHashCode(), color4.GetHashCode());

            var color5 = new IndexedColor();
            var color6 = new IndexedColor(IndexedColor.Value.SystemBackground);
            Assert.NotEqual(color5.GetHashCode(), color6.GetHashCode());
        }

        [Fact(DisplayName = "Test of the implicit operator from IndexedColor.Value to IndexedColor")]
        public void ImplicitOperatorTest()
        {
            IndexedColor color1 = IndexedColor.Value.Red;
            Assert.Equal(IndexedColor.Value.Red, color1.ColorValue);
            IndexedColor color2 = IndexedColor.Value.SystemBackground;
            Assert.Equal(IndexedColor.Value.SystemBackground, color2.ColorValue);
        }
    }
}
