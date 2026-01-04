using NanoXLSX.Colors;
using NanoXLSX.Exceptions;
using NanoXLSX.Themes;
using Xunit;

namespace NanoXLSX.Core.Test.Colors
{
    public class ColorTest
    {
        [Fact(DisplayName = "Test of the CreateNone function")]
        public void CreateNoneTest()
        {
            Color color = Color.CreateNone();
            Assert.Equal(Color.ColorType.None, color.Type);
            Assert.False(color.IsDefined);
            Assert.Null(color.Value);
        }

        [Fact(DisplayName = "Test of the CreateAuto function")]
        public void CreateAutoTest()
        {
            Color color = Color.CreateAuto();
            Assert.Equal(Color.ColorType.Auto, color.Type);
            Assert.True(color.Auto);
            Assert.True(color.IsDefined);
            Assert.NotNull(color.Value);
        }

        [Theory(DisplayName = "Test of the CreateRgb function")]
        [InlineData("000000", "FF000000")]
        [InlineData("FFFFFF", "FFFFFFFF")]
        [InlineData("AABBCC", "FFAABBCC")]
        [InlineData("FF000000", "FF000000")]
        [InlineData("FFFFFFFF", "FFFFFFFF")]
        [InlineData("FFAABBCC", "FFAABBCC")]
        public void CreateRgbFromStringTest(string givenRgb, string expectedRgb)
        {
            Color color = Color.CreateRgb(givenRgb);
            Assert.Equal(Color.ColorType.Rgb, color.Type);
            Assert.Equal(expectedRgb, color.GetArgbValue(), ignoreCase: true);
        }

        [Theory(DisplayName = "Test of the CreateRgb function, using a SrgbColor instance")]
        [InlineData("000000", "FF000000")]
        [InlineData("FFFFFF", "FFFFFFFF")]
        [InlineData("AABBCC", "FFAABBCC")]
        [InlineData("FF000000", "FF000000")]
        [InlineData("FFFFFFFF", "FFFFFFFF")]
        [InlineData("FFAABBCC", "FFAABBCC")]
        public void CreateRgbFromStringTest2(string givenRgb, string expectedRgb)
        {
            SrgbColor color = new SrgbColor(givenRgb);
            Color c = Color.CreateRgb(color);
            Assert.Equal(Color.ColorType.Rgb, c.Type);
            Assert.Equal(expectedRgb, c.GetArgbValue(), ignoreCase: true);
        }

        [Theory(DisplayName = "Test of the failing CreateRgb function")]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("XYZ")]
        [InlineData("FFAABBCCDD")]
        [InlineData("FFAAB")]
        public void CreateRgbFromStringFailureTest(string rgb)
        {
            Assert.Throws<StyleException>(() => Color.CreateRgb(rgb));
        }

        [Theory(DisplayName = "Test of the CreateIndexed function")]
        [InlineData(0)]
        [InlineData(8)]
        [InlineData(64)]
        public void CreateIndexedTest(int index)
        {
            Color color = Color.CreateIndexed(index);
            Assert.Equal(Color.ColorType.Indexed, color.Type);
            Assert.NotNull(color.IndexedColor);
            Assert.NotNull(color.GetArgbValue());
            Assert.Equal(index, (int)color.IndexedColor.ColorValue);
        }

        [Theory(DisplayName = "Test of the CreateIndexed function, using a IndexedColor instance")]
        [InlineData(0)]
        [InlineData(8)]
        [InlineData(64)]
        public void CreateIndexedTest2(int index)
        {
            IndexedColor color = new IndexedColor(index);
            Color color2 = Color.CreateIndexed(color);
            Assert.Equal(Color.ColorType.Indexed, color2.Type);
            Assert.NotNull(color2.IndexedColor);
            Assert.NotNull(color2.GetArgbValue());
            Assert.Equal(index, (int)color2.IndexedColor.ColorValue);
        }

        [Theory(DisplayName = "Test of the CreateIndexed function, using a IndexedColor enum value")]
        [InlineData(IndexedColor.Value.Black)]
        [InlineData(IndexedColor.Value.StrongMagenta)]
        [InlineData(IndexedColor.Value.SystemBackground)]
        [InlineData(IndexedColor.Value.SystemForeground)]
        public void CreateIndexedTest3(IndexedColor.Value value)
        {
            IndexedColor color = new IndexedColor(value);
            Color color2 = Color.CreateIndexed(color);
            Assert.Equal(Color.ColorType.Indexed, color2.Type);
            Assert.NotNull(color2.IndexedColor);
            Assert.NotNull(color2.GetArgbValue());
            Assert.Equal(value, color2.IndexedColor.ColorValue);
        }

        [Theory(DisplayName = "Test of the failing CreateIndexed function")]
        [InlineData(-1)]
        [InlineData(66)]
        public void CreateIndexedFailureTest(int index)
        {
            Assert.Throws<StyleException>(() => Color.CreateIndexed(index));
        }

        [Fact(DisplayName = "Test of the failing CreateIndexed function when passing null")]
        public void CreateIndexedFailureTest2()
        {
            Assert.Throws<StyleException>(() => Color.CreateIndexed(null));
        }

        [Theory(DisplayName = "Test of the CreateTheme function")]
        [InlineData(Theme.ColorSchemeElement.Accent1)]
        [InlineData(Theme.ColorSchemeElement.Dark1)]
        [InlineData(Theme.ColorSchemeElement.FollowedHyperlink)]
        [InlineData(Theme.ColorSchemeElement.Light1)]
        public void CreateThemeTest(Theme.ColorSchemeElement value)
        {
            Color color = Color.CreateTheme(value, 0.25);
            Assert.Equal(Color.ColorType.Theme, color.Type);
            Assert.Equal(0.25, color.Tint);
            Assert.Null(color.GetArgbValue());
            Assert.Equal(value, color.ThemeColor.ColorValue);
        }

        [Theory(DisplayName = "Test of the CreateTheme function, using a ThemeColor instance")]
        [InlineData(Theme.ColorSchemeElement.Accent1)]
        [InlineData(Theme.ColorSchemeElement.Dark1)]
        [InlineData(Theme.ColorSchemeElement.FollowedHyperlink)]
        [InlineData(Theme.ColorSchemeElement.Light1)]
        public void CreateThemeTest2(Theme.ColorSchemeElement value)
        {
            ThemeColor color = new ThemeColor(value);
            Color color2 = Color.CreateTheme(color, -0.25);
            Assert.Equal(Color.ColorType.Theme, color2.Type);
            Assert.Equal(-0.25, color2.Tint);
            Assert.Null(color2.GetArgbValue());
            Assert.Equal(value, color2.ThemeColor.ColorValue);
        }

        [Fact(DisplayName = "Test of the failing CreateTheme function")]
        public void CreateThemeFailureTest()
        {
            Assert.Throws<StyleException>(() => Color.CreateTheme(null));
        }

        [Theory(DisplayName = "Test of the CreateSystem function")]
        [InlineData(SystemColor.Value.ActiveBorder)]
        [InlineData(SystemColor.Value.ButtonText)]
        [InlineData(SystemColor.Value.Highlight)]
        [InlineData(SystemColor.Value.Window)]
        public void CreateSystemTest(SystemColor.Value value)
        {
            Color color = Color.CreateSystem(value);
            Assert.Equal(Color.ColorType.System, color.Type);
            Assert.Null(color.Tint);
            Assert.Null(color.GetArgbValue());
            Assert.Equal(value, color.SystemColor.ColorValue);
        }

        [Theory(DisplayName = "Test of the CreateSystem function, using a SystemColor instance")]
        [InlineData(SystemColor.Value.ActiveBorder)]
        [InlineData(SystemColor.Value.ButtonText)]
        [InlineData(SystemColor.Value.Highlight)]
        [InlineData(SystemColor.Value.Window)]
        public void CreateSystemTest2(SystemColor.Value value)
        {
            SystemColor color = new SystemColor(value);
            Color color2 = Color.CreateSystem(color);
            Assert.Equal(Color.ColorType.System, color2.Type);
            Assert.Null(color2.Tint);
            Assert.Null(color2.GetArgbValue());
            Assert.Equal(value, color2.SystemColor.ColorValue);
        }

        [Fact(DisplayName = "Test of the failing CreateSystem function")]
        public void CreateSystemFailureTest()
        {
            Assert.Throws<StyleException>(() => Color.CreateSystem(null));
        }

        [Theory(DisplayName = "Test of the implicit operator function, when using a string")]
        [InlineData("000000", "FF000000")]
        [InlineData("FFFFFF", "FFFFFFFF")]
        [InlineData("123456", "FF123456")]
        [InlineData("FF000000", "FF000000")]
        [InlineData("FFFFFFFF", "FFFFFFFF")]
        [InlineData("FF234567", "FF234567")]
        public void ImplicitRgbConversionTest(string givrnRgb, string expectedRgb)
        {
            Color color = givrnRgb;
            Assert.Equal(Color.ColorType.Rgb, color.Type);
            Assert.Equal(expectedRgb, color.GetArgbValue(), ignoreCase: true);
        }

        [Theory(DisplayName = "Test of the implicit operator function, when using a value of IndexedColor.Value")]
        [InlineData(IndexedColor.Value.Black)]
        [InlineData(IndexedColor.Value.Black0)]
        [InlineData(IndexedColor.Value.Cyan)]
        [InlineData(IndexedColor.Value.SystemBackground)]
        [InlineData(IndexedColor.Value.SystemForeground)]
        public void ImplicitIndexedConversionTest(IndexedColor.Value index)
        {
            Color color = index;
            Assert.Equal(Color.ColorType.Indexed, color.Type);
            Assert.NotNull(color.GetArgbValue());
            Assert.Equal(index, color.IndexedColor.ColorValue);
        }


        [Theory(DisplayName = "Test of the implicit operator function, when using an int")]
        [InlineData(5)]
        [InlineData(0)]
        [InlineData(22)]
        [InlineData(65)]
        public void ImplicitIndexedConversionTest2(int index)
        {
            Color color = index;
            Assert.Equal(Color.ColorType.Indexed, color.Type);
            Assert.NotNull(color.GetArgbValue());
            Assert.Equal(index, (int)color.IndexedColor.ColorValue);
        }

        [Theory(DisplayName = "Test of the failing implicit operator function, when using a string")]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("XYZ")]
        [InlineData("FFAABBCCDD")]
        [InlineData("FFAAB")]
        public void FailingImplicitRgbConversionTest(string value)
        {
            Assert.Throws<StyleException>(() => { Color c = value; });
        }

        [Theory(DisplayName = "Test of the failing implicit operator function, when using an int")]
        [InlineData(-10)]
        [InlineData(100)]
        public void FailingImplicitIndexedConversionTest(int index)
        {
            Assert.Throws<StyleException>(() => { Color c = index; });
        }

        [Fact(DisplayName = "Test of the Value property on None")]
        public void ValueNoneTest()
        {
            Color color = Color.CreateNone();
            Assert.Null(color.Value);
        }

        [Fact(DisplayName = "Test of the Value property on Auto")]
        public void ValueAutoTest()
        {
            Color color = Color.CreateAuto();
            Assert.True(color.Value is AutoColor);
        }

        [Theory(DisplayName = "Test of the Value property on sRGB")]
        [InlineData("000000", "FF000000")]
        [InlineData("FFFFFF", "FFFFFFFF")]
        [InlineData("123456", "FF123456")]
        [InlineData("FF000000", "FF000000")]
        [InlineData("FFFFFFFF", "FFFFFFFF")]
        [InlineData("FF234567", "FF234567")]
        public void ValueSrgbTest(string givenRgbValue, string expectedRgbValue)
        {
            Color color = Color.CreateRgb(givenRgbValue);
            Assert.True(color.Value is SrgbColor);
            Assert.Equal(expectedRgbValue, color.RgbColor.ColorValue);
        }

        [Theory(DisplayName = "Test of the Value property on indexed colors")]
        [InlineData(IndexedColor.Value.Black)]
        [InlineData(IndexedColor.Value.StrongMagenta)]
        [InlineData(IndexedColor.Value.SystemBackground)]
        [InlineData(IndexedColor.Value.SystemForeground)]
        public void ValueIndexedTest(IndexedColor.Value indexedValue)
        {
            Color color = Color.CreateIndexed(indexedValue);
            Assert.True(color.Value is IndexedColor);
            Assert.Equal(indexedValue, color.IndexedColor.ColorValue);
        }

        [Theory(DisplayName = "Test of the Value property on system colors")]
        [InlineData(SystemColor.Value.ActiveBorder)]
        [InlineData(SystemColor.Value.ButtonText)]
        [InlineData(SystemColor.Value.Highlight)]
        [InlineData(SystemColor.Value.Window)]
        public void ValueSystemTest(SystemColor.Value systemColor)
        {
            Color color = Color.CreateSystem(systemColor);
            Assert.True(color.Value is SystemColor);
            Assert.Equal(systemColor, color.SystemColor.ColorValue);
        }

        [Theory(DisplayName = "Test of the Value property on theme colors")]
        [InlineData(Theme.ColorSchemeElement.Accent1)]
        [InlineData(Theme.ColorSchemeElement.Dark1)]
        [InlineData(Theme.ColorSchemeElement.FollowedHyperlink)]
        [InlineData(Theme.ColorSchemeElement.Light1)]
        public void ValueThemeTest(Theme.ColorSchemeElement themeElement)
        {
            Color color = Color.CreateTheme(themeElement);
            Assert.True(color.Value is ThemeColor);
            Assert.Equal(themeElement, ((ThemeColor)color.Value).ColorValue);
        }

        [Theory(DisplayName = "Test of the GetArgbValue function on a sRGB color")]
        [InlineData("000000", "FF000000")]
        [InlineData("FFFFFF", "FFFFFFFF")]
        [InlineData("123456", "FF123456")]
        [InlineData("FF000000", "FF000000")]
        [InlineData("FFFFFFFF", "FFFFFFFF")]
        [InlineData("FF234567", "FF234567")]
        public void GetArgbValueSRgbTest(string givenRgb, string expectedRgb)
        {
            Color color = Color.CreateRgb(givenRgb);
            Assert.Equal(expectedRgb, color.GetArgbValue());
        }

        [Theory(DisplayName = "Test of the GetArgbValue function on a sRGB color")]
        [InlineData(IndexedColor.Value.Black0, "FF000000")]
        [InlineData(IndexedColor.Value.Black, "FF000000")]
        [InlineData(IndexedColor.Value.White, "FFFFFFFF")]
        [InlineData(IndexedColor.Value.StrongCyan, "FF00FFFF")]
        [InlineData(IndexedColor.Value.DarkMaroon, "FF800000")]
        [InlineData(IndexedColor.Value.Lavender, "FFCC99FF")]
        public void GetArgbValueIndexedTest(IndexedColor.Value givenIndex, string expectedRgb)
        {
            Color color = Color.CreateIndexed(givenIndex);
            Assert.Equal(expectedRgb, color.GetArgbValue());
        }


        [Fact(DisplayName = "Test of the GetArgbValue function on a theme color")]
        public void GetArgbValueReturnsNullForThemeTest()
        {
            Color color = Color.CreateTheme(Theme.ColorSchemeElement.Dark1);
            Assert.Null(color.GetArgbValue());
        }

        [Fact(DisplayName = "Test of the GetArgbValue function on a system color")]
        public void GetArgbValueReturnsNullForSystemTest()
        {
            Color color = Color.CreateSystem(SystemColor.Value.ActiveBorder);
            Assert.Null(color.GetArgbValue());
        }

        [Fact(DisplayName = "Test of the GetArgbValue function on a auto color")]
        public void GetArgbValueReturnsNullForAutoTest()
        {
            Color color = Color.CreateAuto();
            Assert.Null(color.GetArgbValue());
        }


        [Fact(DisplayName = "Test of the Equals method on equality")]
        public void EqualsSameRgbValueTest()
        {
            Color a = Color.CreateRgb("FFABCDEF");
            Color b = Color.CreateRgb("FFABCDEF");
            Assert.Equal(a, b);
            Assert.True(a.Equals(b));
        }

        [Fact(DisplayName = "Test of the Equals method on inequality")]
        public void EqualsDifferentRgbValueTest()
        {
            Color a = Color.CreateRgb("FFABCDEF");
            Color b = Color.CreateRgb("FFABCDEE");
            Assert.NotEqual(a, b);
        }

        [Fact(DisplayName = "Test of the Equals method on inequality on different types")]
        public void EqualsDifferentTypeTest()
        {
            Color a = Color.CreateRgb("FF000000");
            Color b = Color.CreateIndexed(0);
            Assert.NotEqual(a, b);
        }

        [Fact(DisplayName = "Test of the GetHasCode method on equality")]
        public void GetHashCodeEqualObjectsTest()
        {
            Color a = Color.CreateRgb("FF112233");
            Color b = Color.CreateRgb("FF112233");
            Assert.Equal(a.GetHashCode(), b.GetHashCode());
        }

        [Fact(DisplayName = "Test of the GetHasCode method on inequality")]
        public void GetHashCodeDifferentObjectsTest()
        {
            Color a = Color.CreateRgb("FF112233");
            Color b = Color.CreateRgb("FF332211");
            Assert.NotEqual(a.GetHashCode(), b.GetHashCode());
        }


        [Fact(DisplayName = "Test of the CompareTo method on null values")]
        public void CompareToNullTest()
        {
            Color color = Color.CreateRgb("FF000000");
            Assert.True(color.CompareTo(null) > 0);
        }

        [Fact(DisplayName = "Test of the CompareTo method on different types")]
        public void CompareToWrongTypeTest()
        {
            Color color = Color.CreateRgb("FF000000");
            Assert.Throws<StyleException>(() => color.CompareTo("not a color"));
        }

        [Fact(DisplayName = "Test of the CompareTo method on two none color types")]
        public void CompareNoneColorTypeTest()
        {
            Color a = Color.CreateNone();
            Color b = Color.CreateNone();
            Assert.Equal(0, a.CompareTo(b));
        }

        [Fact(DisplayName = "Test of the CompareTo method on two auto color types")]
        public void CompareAutoColorTypeTest()
        {
            Color a = Color.CreateAuto();
            Color b = Color.CreateAuto();
            Assert.Equal(0, a.CompareTo(b));
        }

        [Theory(DisplayName = "Test of the CompareTo method on identical RGB/ARGB values")]
        [InlineData("000000")]
        [InlineData("FFFFFF")]
        [InlineData("AABBCC")]
        [InlineData("FF000000")]
        [InlineData("FFFFFFFF")]
        [InlineData("FFAABBCC")]
        public void CompareToSameRgbTest(string rgbValue)
        {
            Color a = Color.CreateRgb(rgbValue);
            Color b = Color.CreateRgb(rgbValue);
            Assert.Equal(0, a.CompareTo(b));
        }

        [Fact(DisplayName = "Test of the CompareTo method on different sRGB values")]
        public void CompareToRgbOrderingTest()
        {
            Color a = Color.CreateRgb("FF000000");
            Color b = Color.CreateRgb("FFFFFFFF");
            Assert.True(a.CompareTo(b) < 0);
        }

        [Fact(DisplayName = "Test of the CompareTo method on different color values if sRGB and indexes are compared")]
        public void CompareToDifferentTypeOrderingTest()
        {
            Color rgb = Color.CreateRgb("FF000000");
            Color indexed = Color.CreateIndexed(0);
            Assert.NotEqual(0, rgb.CompareTo(indexed));
        }

        [Fact(DisplayName = "Test of the CompareTo method on different tint values")]
        public void CompareToThemeTintTest()
        {
            Color a = Color.CreateTheme(Theme.ColorSchemeElement.Accent1, 0.1);
            Color b = Color.CreateTheme(Theme.ColorSchemeElement.Accent1, 0.2);
            Assert.True(a.CompareTo(b) < 0);
        }

        [Fact(DisplayName = "Test of the CompareTo method on colors with different theme slots")]
        public void CompareToThemeDifferentThemeSlots()
        {
            Color c1 = Color.CreateTheme(Theme.ColorSchemeElement.Dark1);
            Color c2 = Color.CreateTheme(Theme.ColorSchemeElement.Accent1);
            int result = c1.CompareTo(c2);
            Assert.True(result < 0);
        }

        [Fact(DisplayName = "Test of the CompareTo method on colors with same slot but different tint")]
        public void CompareToThemeSameSlotDifferentTint()
        {
            Color c1 = Color.CreateTheme(Theme.ColorSchemeElement.Accent1, tint: -0.2);
            Color c2 = Color.CreateTheme(Theme.ColorSchemeElement.Accent1, tint: 0.2);
            int result = c1.CompareTo(c2);
            Assert.True(result < 0);
        }

        [Fact(DisplayName = "Test of the CompareTo method on System colors")]
        public void CompareToSystemColors()
        {
            Color c1 = Color.CreateSystem(new SystemColor(SystemColor.Value.AppWorkspace));
            Color c2 = Color.CreateSystem(new SystemColor(SystemColor.Value.Menu));
            int result = c1.CompareTo(c2);
            Assert.NotEqual(0, result);
        }

        [Fact(DisplayName = "Test of the CompareTo method on a defensive fallback path")]
        public void CompareToDefensiveFallback()
        {
            Color c1 = Color.CreateRgb("FF0000");
            Color c2 = Color.CreateRgb("00FF00");
            typeof(Color)
                .GetProperty(nameof(Color.Type))
                .SetValue(c1, (Color.ColorType)999);
            typeof(Color)
                .GetProperty(nameof(Color.Type))
                .SetValue(c2, (Color.ColorType)999);
            int result = c1.CompareTo(c2);
            Assert.Equal(0, result);
        }

        [Fact(DisplayName = "Test of the CompareTo method on indexed colors uses numeric index")]
        public void CompareToIndexedNumericComparison()
        {
            Color c1 = Color.CreateIndexed(IndexedColor.Value.Black);
            Color c2 = Color.CreateIndexed(IndexedColor.Value.White);
            int result = c1.CompareTo(c2);
            Assert.True(result < 0); // Both are invalid types - corner case
        }

        [Theory(DisplayName = "Test of the ToStringFunction (for code coverage)")]
        [InlineData(Color.ColorType.Rgb)]
        [InlineData(Color.ColorType.Indexed)]
        [InlineData(Color.ColorType.Theme)]
        [InlineData(Color.ColorType.System)]
        [InlineData(Color.ColorType.Auto)]
        [InlineData(Color.ColorType.None)]
        public void ToStringTest(Color.ColorType type)
        {
            Color color;
            string expectedToken;
            switch (type)
            {
                case Color.ColorType.Rgb:
                    color = Color.CreateRgb("FFAABB");
                    expectedToken = "FFAABB";
                    break;
                case Color.ColorType.Indexed:
                    color = Color.CreateIndexed(IndexedColor.Value.Rose);
                    expectedToken = ((int)IndexedColor.Value.Rose).ToString();
                    break;
                case Color.ColorType.Theme:
                    color = Color.CreateTheme(Theme.ColorSchemeElement.Accent5);
                    expectedToken = ((int)Theme.ColorSchemeElement.Accent5).ToString();
                    break;
                case Color.ColorType.System:
                    color = Color.CreateSystem(SystemColor.Value.Background);
                    expectedToken = "Background";
                    break;
                case Color.ColorType.Auto:
                    color = Color.CreateAuto();
                    expectedToken = "Auto";
                    break;
                default:
                    color = Color.CreateNone();
                    expectedToken = "Undefined";
                    break;
            }
            string given = color.ToString().ToLower();
            Assert.Contains(expectedToken.ToLower(), given);
        }

    }
}
