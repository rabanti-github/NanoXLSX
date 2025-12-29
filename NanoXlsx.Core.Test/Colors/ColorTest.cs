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
            var c = Color.CreateNone();

            Assert.Equal(Color.ColorType.None, c.Type);
            Assert.False(c.IsDefined);
            Assert.Null(c.Value);
        }

        [Fact(DisplayName = "Test of the CreateAuto function")]
        public void CreateAutoTest()
        {
            var c = Color.CreateAuto();

            Assert.Equal(Color.ColorType.Auto, c.Type);
            Assert.True(c.Auto);
            Assert.True(c.IsDefined);
            Assert.NotNull(c.Value);
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
            var c = Color.CreateRgb(givenRgb);

            Assert.Equal(Color.ColorType.Rgb, c.Type);
            Assert.Equal(expectedRgb, c.GetArgbValue(), ignoreCase: true);
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
            var c = Color.CreateRgb(color);

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
            var c = Color.CreateIndexed(index);

            Assert.Equal(Color.ColorType.Indexed, c.Type);
            Assert.NotNull(c.IndexedColor);
            Assert.NotNull(c.GetArgbValue());
            Assert.Equal(index, (int)c.IndexedColor.ColorValue);
        }

        [Theory(DisplayName = "Test of the CreateIndexed function, using a IndexedColor instance")]
        [InlineData(0)]
        [InlineData(8)]
        [InlineData(64)]
        public void CreateIndexedTest2(int index)
        {
            IndexedColor color = new IndexedColor(index);
            var c = Color.CreateIndexed(color);

            Assert.Equal(Color.ColorType.Indexed, c.Type);
            Assert.NotNull(c.IndexedColor);
            Assert.NotNull(c.GetArgbValue());
            Assert.Equal(index, (int)c.IndexedColor.ColorValue);
        }

        [Theory(DisplayName = "Test of the CreateIndexed function, using a IndexedColor enum value")]
        [InlineData(IndexedColor.Value.Black)]
        [InlineData(IndexedColor.Value.StrongMagenta)]
        [InlineData(IndexedColor.Value.SystemBackground)]
        [InlineData(IndexedColor.Value.SystemForeground)]
        public void CreateIndexedTest3(IndexedColor.Value value)
        {
            IndexedColor color = new IndexedColor(value);
            var c = Color.CreateIndexed(color);

            Assert.Equal(Color.ColorType.Indexed, c.Type);
            Assert.NotNull(c.IndexedColor);
            Assert.NotNull(c.GetArgbValue());
            Assert.Equal(value, c.IndexedColor.ColorValue);
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
            var c = Color.CreateTheme(value, 0.25);

            Assert.Equal(Color.ColorType.Theme, c.Type);
            Assert.Equal(0.25, c.Tint);
            Assert.Null(c.GetArgbValue());
            Assert.Equal(value, c.ThemeColor.ColorValue);
        }

        [Theory(DisplayName = "Test of the CreateTheme function, using a ThemeColor instance")]
        [InlineData(Theme.ColorSchemeElement.Accent1)]
        [InlineData(Theme.ColorSchemeElement.Dark1)]
        [InlineData(Theme.ColorSchemeElement.FollowedHyperlink)]
        [InlineData(Theme.ColorSchemeElement.Light1)]
        public void CreateThemeTest2(Theme.ColorSchemeElement value)
        {
            ThemeColor color = new ThemeColor(value);
            var c = Color.CreateTheme(color, -0.25);

            Assert.Equal(Color.ColorType.Theme, c.Type);
            Assert.Equal(-0.25, c.Tint);
            Assert.Null(c.GetArgbValue());
            Assert.Equal(value, c.ThemeColor.ColorValue);
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
            var c = Color.CreateSystem(value);

            Assert.Equal(Color.ColorType.System, c.Type);
            Assert.Null(c.Tint);
            Assert.Null(c.GetArgbValue());
            Assert.Equal(value, c.SystemColor.ColorValue);
        }

        [Theory(DisplayName = "Test of the CreateSystem function, using a SystemColor instance")]
        [InlineData(SystemColor.Value.ActiveBorder)]
        [InlineData(SystemColor.Value.ButtonText)]
        [InlineData(SystemColor.Value.Highlight)]
        [InlineData(SystemColor.Value.Window)]
        public void CreateSystemTest2(SystemColor.Value value)
        {
            SystemColor color = new SystemColor(value);
            var c = Color.CreateSystem(color);

            Assert.Equal(Color.ColorType.System, c.Type);
            Assert.Null(c.Tint);
            Assert.Null(c.GetArgbValue());
            Assert.Equal(value, c.SystemColor.ColorValue);
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
            Color c = givrnRgb;

            Assert.Equal(Color.ColorType.Rgb, c.Type);
            Assert.Equal(expectedRgb, c.GetArgbValue(), ignoreCase: true);
        }

        [Theory(DisplayName = "Test of the implicit operator function, when using an int")]
        [InlineData(5)]
        [InlineData(0)]
        [InlineData(22)]
        [InlineData(65)]
        public void ImplicitIndexedConversionTest(int index)
        {
            Color c = index;

            Assert.Equal(Color.ColorType.Indexed, c.Type);
            Assert.NotNull(c.GetArgbValue());
            Assert.Equal(index, (int)c.IndexedColor.ColorValue);
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
            var c = Color.CreateRgb(givenRgb);
            Assert.Equal(expectedRgb, c.GetArgbValue());
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
            var c = Color.CreateIndexed(givenIndex);
            Assert.Equal(expectedRgb, c.GetArgbValue());
        }


        [Fact(DisplayName = "Test of the GetArgbValue function on a theme color")]
        public void GetArgbValueReturnsNullForThemeTest()
        {
            var c = Color.CreateTheme(Theme.ColorSchemeElement.Dark1);

            Assert.Null(c.GetArgbValue());
        }

        [Fact(DisplayName = "Test of the GetArgbValue function on a system color")]
        public void GetArgbValueReturnsNullForSystemTest()
        {
            var c = Color.CreateSystem(SystemColor.Value.ActiveBorder);

            Assert.Null(c.GetArgbValue());
        }

        [Fact(DisplayName = "Test of the GetArgbValue function on a auto color")]
        public void GetArgbValueReturnsNullForAutoTest()
        {
            var c = Color.CreateAuto();

            Assert.Null(c.GetArgbValue());
        }


        [Fact(DisplayName = "Test of the Equals method on equality")]
        public void EqualsSameRgbValueTest()
        {
            var a = Color.CreateRgb("FFABCDEF");
            var b = Color.CreateRgb("FFABCDEF");

            Assert.Equal(a, b);
            Assert.True(a.Equals(b));
        }

        [Fact(DisplayName = "Test of the Equals method on inequality")]
        public void EqualsDifferentRgbValueTest()
        {
            var a = Color.CreateRgb("FFABCDEF");
            var b = Color.CreateRgb("FFABCDEE");

            Assert.NotEqual(a, b);
        }

        [Fact(DisplayName = "Test of the Equals method on inequality on different types")]
        public void EqualsDifferentTypeTest()
        {
            var a = Color.CreateRgb("FF000000");
            var b = Color.CreateIndexed(0);

            Assert.NotEqual(a, b);
        }

        [Fact(DisplayName = "Test of the GetHasCode method on equality")]
        public void GetHashCodeEqualObjectsTest()
        {
            var a = Color.CreateRgb("FF112233");
            var b = Color.CreateRgb("FF112233");

            Assert.Equal(a.GetHashCode(), b.GetHashCode());
        }

        [Fact(DisplayName = "Test of the GetHasCode method on inequality")]
        public void GetHashCodeDifferentObjectsTest()
        {
            var a = Color.CreateRgb("FF112233");
            var b = Color.CreateRgb("FF332211");

            Assert.NotEqual(a.GetHashCode(), b.GetHashCode());
        }


        [Fact(DisplayName = "Test of the CompareTo method on null values")]
        public void CompareToNullTest()
        {
            var c = Color.CreateRgb("FF000000");

            Assert.True(c.CompareTo(null) > 0);
        }

        [Fact(DisplayName = "Test of the CompareTo method on different types")]
        public void CompareToWrongTypeTest()
        {
            var c = Color.CreateRgb("FF000000");

            Assert.Throws<StyleException>(() => c.CompareTo("not a color"));
        }

        [Fact(DisplayName = "Test of the CompareTo method on two none color types")]
        public void CompareNoneColorTypeTest()
        {
            var a = Color.CreateNone();
            var b = Color.CreateNone();
            Assert.Equal(0, a.CompareTo(b));
        }

        [Fact(DisplayName = "Test of the CompareTo method on two auto color types")]
        public void CompareAutoColorTypeTest()
        {
            var a = Color.CreateAuto();
            var b = Color.CreateAuto();
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
            var a = Color.CreateRgb(rgbValue);
            var b = Color.CreateRgb(rgbValue);

            Assert.Equal(0, a.CompareTo(b));
        }

        [Fact(DisplayName = "Test of the CompareTo method on different sRGB values")]
        public void CompareToRgbOrderingTest()
        {
            var a = Color.CreateRgb("FF000000");
            var b = Color.CreateRgb("FFFFFFFF");

            Assert.True(a.CompareTo(b) < 0);
        }

        [Fact(DisplayName = "Test of the CompareTo method on different color values if sRGB and indexes are compared")]
        public void CompareToDifferentTypeOrderingTest()
        {
            var rgb = Color.CreateRgb("FF000000");
            var indexed = Color.CreateIndexed(0);

            Assert.NotEqual(0, rgb.CompareTo(indexed));
        }

        [Fact(DisplayName = "Test of the CompareTo method on different tint values")]
        public void CompareToThemeTintTest()
        {
            var a = Color.CreateTheme(Theme.ColorSchemeElement.Accent1, 0.1);
            var b = Color.CreateTheme(Theme.ColorSchemeElement.Accent1, 0.2);

            Assert.True(a.CompareTo(b) < 0);
        }

        [Fact(DisplayName = "Test of the CompareTo method on colors with different theme slots")]
        public void CompareToThemeDifferentThemeSlots()
        {
            var c1 = Color.CreateTheme(Theme.ColorSchemeElement.Dark1);
            var c2 = Color.CreateTheme(Theme.ColorSchemeElement.Accent1);

            int result = c1.CompareTo(c2);

            Assert.True(result < 0);
        }

        [Fact(DisplayName = "Test of the CompareTo method on colors with same slot but different tint")]
        public void CompareToThemeSameSlotDifferentTint()
        {
            var c1 = Color.CreateTheme(Theme.ColorSchemeElement.Accent1, tint: -0.2);
            var c2 = Color.CreateTheme(Theme.ColorSchemeElement.Accent1, tint: 0.2);

            int result = c1.CompareTo(c2);

            Assert.True(result < 0);
        }

        [Fact(DisplayName = "Test of the CompareTo method on System colors")]
        public void CompareToSystemColors()
        {
            var c1 = Color.CreateSystem(new SystemColor(SystemColor.Value.AppWorkspace));
            var c2 = Color.CreateSystem(new SystemColor(SystemColor.Value.Menu));

            int result = c1.CompareTo(c2);

            Assert.NotEqual(0, result);
        }

        [Fact(DisplayName = "Test of the CompareTo method on a defensive fallback path")]
        public void CompareToDefensiveFallback()
        {
            var c1 = Color.CreateRgb("FF0000");
            var c2 = Color.CreateRgb("00FF00");

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
            var c1 = Color.CreateIndexed(IndexedColor.Value.Black); // e.g. 8
            var c2 = Color.CreateIndexed(IndexedColor.Value.White); // e.g. 9

            int result = c1.CompareTo(c2);

            Assert.True(result < 0); // Both are invalid types - corner case
        }

    }
}
