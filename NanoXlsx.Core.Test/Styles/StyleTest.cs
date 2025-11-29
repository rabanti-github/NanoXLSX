using NanoXLSX.Exceptions;
using NanoXLSX.Styles;
using Xunit;
using static NanoXLSX.Styles.Border;
using static NanoXLSX.Styles.CellXf;
using static NanoXLSX.Styles.Font;
using static NanoXLSX.Styles.NumberFormat;

namespace NanoXLSX.Test.Core.StyleTest
{
    public class StyleTest
    {
        [Fact(DisplayName = "Test of the get and set function of the CurrentBorder property")]
        public void CurrentBorderTest()
        {
            Style style = new Style();
            Border border = new Border();
            Assert.NotNull(style.CurrentBorder);
            Assert.Equal(border.GetHashCode(), style.CurrentBorder.GetHashCode());
            style.CurrentBorder = border;
            border.BottomColor = "FFAABBCC";
            Assert.Equal("FFAABBCC", style.CurrentBorder.BottomColor);
        }

        [Fact(DisplayName = "Test of the get and set function of the CurrentCellXf property")]
        public void CurrentCellXfTest()
        {
            Style style = new Style();
            CellXf cellXf = new CellXf();
            Assert.NotNull(style.CurrentCellXf);
            Assert.Equal(cellXf.GetHashCode(), style.CurrentCellXf.GetHashCode());
            style.CurrentCellXf = cellXf;
            cellXf.Indent = 5;
            Assert.Equal(5, style.CurrentCellXf.Indent);
        }

        [Fact(DisplayName = "Test of the get and set function of the CurrentFill property")]
        public void CurrentFillTest()
        {
            Style style = new Style();
            Fill fill = new Fill();
            Assert.NotNull(style.CurrentFill);
            Assert.Equal(fill.GetHashCode(), style.CurrentFill.GetHashCode());
            style.CurrentFill = fill;
            fill.BackgroundColor = "AACCBBDD";
            Assert.Equal("AACCBBDD", style.CurrentFill.BackgroundColor);
        }

        [Fact(DisplayName = "Test of the get and set function of the CurrentFont property")]
        public void CurrentFontTest()
        {
            Style style = new Style();
            Font font = new Font();
            Assert.NotNull(style.CurrentFont);
            Assert.Equal(font.GetHashCode(), style.CurrentFont.GetHashCode());
            style.CurrentFont = font;
            font.Name = "Sans Serif";
            Assert.Equal("Sans Serif", style.CurrentFont.Name);
        }

        [Fact(DisplayName = "Test of the get and set function of the CurrentNumberFormat property")]
        public void CurrentNumberFormatTest()
        {
            Style style = new Style();
            NumberFormat numberFormat = new NumberFormat();
            Assert.NotNull(style.CurrentFill);
            Assert.Equal(numberFormat.GetHashCode(), style.CurrentNumberFormat.GetHashCode());
            style.CurrentNumberFormat = numberFormat;
            numberFormat.Number = FormatNumber.Format15;
            Assert.Equal(FormatNumber.Format15, style.CurrentNumberFormat.Number);
        }

        [Fact(DisplayName = "Test of the get and set function of the Name property")]
        public void NameTest()
        {
            Style style = new Style();
            Assert.Equal(style.GetHashCode().ToString(), style.Name);
            style.Name = "Test";
            Assert.Equal("Test", style.Name);
        }

        [Fact(DisplayName = "Test of the get function of the IsInternalStyle property")]
        public void IsInternalStyleTest()
        {
            Style style = new Style();
            Assert.False(style.IsInternalStyle);
            Style internalStyle = new Style("test", 0, true);
            Assert.True(internalStyle.IsInternalStyle);
        }

        [Fact(DisplayName = "Test of the get and set function of the InternalID property")]
        public void InternalIDTest()
        {
            Style style = new Style();
            Assert.Null(style.InternalID);
            style.InternalID = 962;
            Assert.Equal(962, style.InternalID);
        }

        [Fact(DisplayName = "Test of the default constructor")]
        public void ConstructorTest()
        {
            Style style = new Style();
            Assert.NotNull(style.CurrentBorder);
            Assert.NotNull(style.CurrentCellXf);
            Assert.NotNull(style.CurrentFill);
            Assert.NotNull(style.CurrentFont);
            Assert.NotNull(style.CurrentNumberFormat);
            Assert.NotNull(style.Name);
            Assert.Null(style.InternalID);
        }

        [Fact(DisplayName = "Test of the constructor with a name")]
        public void ConstructorTest2()
        {
            Style style = new Style("test1");
            Assert.NotNull(style.CurrentBorder);
            Assert.NotNull(style.CurrentCellXf);
            Assert.NotNull(style.CurrentFill);
            Assert.NotNull(style.CurrentFont);
            Assert.NotNull(style.CurrentNumberFormat);
            Assert.Equal("test1", style.Name);
            Assert.Null(style.InternalID);
        }

        [Theory(DisplayName = "Test of the constructor for internal styles")]
        [InlineData("test", 0, false)]
        [InlineData("test2", 777, false)]
        [InlineData("test3", -17, true)]
        public void ConstructorTest3(string name, int forceOrder, bool isInternal)
        {
            Style style = new Style(name, forceOrder, isInternal);
            Assert.NotNull(style.CurrentBorder);
            Assert.NotNull(style.CurrentCellXf);
            Assert.NotNull(style.CurrentFill);
            Assert.NotNull(style.CurrentFont);
            Assert.NotNull(style.CurrentNumberFormat);
            Assert.Equal(name, style.Name);
            Assert.Equal(isInternal, style.IsInternalStyle);
            Assert.Equal(forceOrder, style.InternalID);
        }

        [Fact(DisplayName = "Test of the Append function on a Border object")]
        public void AppendTest()
        {
            Style style = new Style();
            Border border = new Border();
            Assert.Equal(border.GetHashCode(), style.CurrentBorder.GetHashCode());
            Border modified = new Border
            {
                BottomColor = "FFAABBCC",
                BottomStyle = StyleValue.DashDotDot
            };
            style.Append(modified);
            Assert.Equal(modified.GetHashCode(), style.CurrentBorder.GetHashCode());
        }

        [Fact(DisplayName = "Test of the Append function on a Font object")]
        public void AppendTest2()
        {
            Style style = new Style();
            Font font = new Font();
            Assert.Equal(font.GetHashCode(), style.CurrentFont.GetHashCode());
            Font modified = new Font
            {
                Bold = true,
                Family = FontFamilyValue.Modern
            };
            style.Append(modified);
            Assert.Equal(modified.GetHashCode(), style.CurrentFont.GetHashCode());
        }

        [Fact(DisplayName = "Test of the Append function on a Fill object")]
        public void AppendTest3()
        {
            Style style = new Style();
            Fill fill = new Fill();
            Assert.Equal(fill.GetHashCode(), style.CurrentFill.GetHashCode());
            Fill modified = new Fill
            {
                BackgroundColor = "FFAABBCC",
                ForegroundColor = "FF112233"
            };
            style.Append(modified);
            Assert.Equal(modified.GetHashCode(), style.CurrentFill.GetHashCode());
        }

        [Fact(DisplayName = "Test of the Append function on a CellXf object")]
        public void AppendTest4()
        {
            Style style = new Style();
            CellXf cellXf = new CellXf();
            Assert.Equal(cellXf.GetHashCode(), style.CurrentCellXf.GetHashCode());
            CellXf modified = new CellXf
            {
                HorizontalAlign = HorizontalAlignValue.Distributed,
                TextRotation = 35
            };
            style.Append(modified);
            Assert.Equal(modified.GetHashCode(), style.CurrentCellXf.GetHashCode());
        }

        [Fact(DisplayName = "Test of the Append function on a NumberFormat object")]
        public void AppendTest5()
        {
            Style style = new Style();
            NumberFormat numberFormat = new NumberFormat();
            Assert.Equal(numberFormat.GetHashCode(), style.CurrentNumberFormat.GetHashCode());
            NumberFormat modified = new NumberFormat
            {
                Number = FormatNumber.Format11
            };
            style.Append(modified);
            Assert.Equal(modified.GetHashCode(), style.CurrentNumberFormat.GetHashCode());
        }

        [Fact(DisplayName = "Test of the Append function on a combination of all components")]
        public void AppendTest6()
        {
            Style style = new Style();
            style.CurrentFont.Size = 18f;
            style.CurrentCellXf.Alignment = TextBreakValue.ShrinkToFit;
            style.CurrentBorder.BottomColor = "FFAA3344";
            style.CurrentFill.BackgroundColor = "FF55AACC";
            style.CurrentNumberFormat.CustomFormatID = 190;
            Font font = new Font
            {
                Name = "Arial"
            };
            CellXf cellXf = new CellXf
            {
                HorizontalAlign = HorizontalAlignValue.Justify
            };
            Border border = new Border
            {
                TopColor = "FF55BB11"
            };
            Fill fill = new Fill
            {
                ForegroundColor = "FFDDDDDD"
            };
            NumberFormat numberFormat = new NumberFormat
            {
                CustomFormatCode = "##--##"
            };

            style.Append(font);
            style.Append(cellXf);
            style.Append(border);
            style.Append(fill);
            style.Append(numberFormat);
            Assert.Equal(18f, style.CurrentFont.Size);
            Assert.Equal("Arial", style.CurrentFont.Name);
            Assert.Equal(TextBreakValue.ShrinkToFit, style.CurrentCellXf.Alignment);
            Assert.Equal(HorizontalAlignValue.Justify, style.CurrentCellXf.HorizontalAlign);
            Assert.Equal("FFAA3344", style.CurrentBorder.BottomColor);
            Assert.Equal("FF55BB11", style.CurrentBorder.TopColor);
            Assert.Equal("FF55AACC", style.CurrentFill.BackgroundColor);
            Assert.Equal("FFDDDDDD", style.CurrentFill.ForegroundColor);
            Assert.Equal(190, style.CurrentNumberFormat.CustomFormatID);
            Assert.Equal("##--##", style.CurrentNumberFormat.CustomFormatCode);
        }

        [Fact(DisplayName = "Test of the Append function on a full other style object")]
        public void AppendTest7()
        {
            Style style = new Style();
            style.CurrentFont.Size = 18f;
            style.CurrentCellXf.Alignment = TextBreakValue.ShrinkToFit;
            style.CurrentBorder.BottomColor = "FFAA3344";
            style.CurrentFill.BackgroundColor = "FF55AACC";
            style.CurrentNumberFormat.CustomFormatID = 190;

            Style style2 = new Style();
            style2.CurrentFont.Name = "Arial";
            style2.CurrentCellXf.HorizontalAlign = HorizontalAlignValue.Justify;
            style2.CurrentBorder.TopColor = "FF55BB11";
            style2.CurrentFill.ForegroundColor = "FFDDDDDD";
            style2.CurrentNumberFormat.CustomFormatCode = "##--##";

            style.Append(style2);
            Assert.Equal(18f, style.CurrentFont.Size);
            Assert.Equal("Arial", style.CurrentFont.Name);
            Assert.Equal(TextBreakValue.ShrinkToFit, style.CurrentCellXf.Alignment);
            Assert.Equal(HorizontalAlignValue.Justify, style.CurrentCellXf.HorizontalAlign);
            Assert.Equal("FFAA3344", style.CurrentBorder.BottomColor);
            Assert.Equal("FF55BB11", style.CurrentBorder.TopColor);
            Assert.Equal("FF55AACC", style.CurrentFill.BackgroundColor);
            Assert.Equal("FFDDDDDD", style.CurrentFill.ForegroundColor);
            Assert.Equal(190, style.CurrentNumberFormat.CustomFormatID);
            Assert.Equal("##--##", style.CurrentNumberFormat.CustomFormatCode);
        }

        [Fact(DisplayName = "Test of the Append function on a null style component")]
        public void AppendTest8()
        {
            Style style = new Style();
            style.CurrentBorder.BottomColor = "FFAA6677";
            int hashCode = style.GetHashCode();
            style.Append(null);
            Assert.Equal(hashCode, style.GetHashCode());
        }

        [Fact(DisplayName = "Test of the failing Append function on a invalid style component (null instance)")]
        public void AppendFailTest()
        {
            Style style = new Style();
            Style style2 = new Style();
            style.CurrentBorder = null;
            Assert.Throws<StyleException>(() => style2.Append(style));
            style2 = new Style();
            style.CurrentCellXf = null;
            Assert.Throws<StyleException>(() => style2.Append(style));
            style2 = new Style();
            style.CurrentFill = null;
            Assert.Throws<StyleException>(() => style2.Append(style));
            style2 = new Style();
            style.CurrentFont = null;
            Assert.Throws<StyleException>(() => style2.Append(style));
            style2 = new Style();
            style.CurrentNumberFormat = null;
            Assert.Throws<StyleException>(() => style2.Append(style));
        }


        [Fact(DisplayName = "Test of the failing GetHashCode function on a invalid style component (null instance)")]
        public void GetHashCodeFailTest()
        {
            Style style = new Style
            {
                CurrentBorder = null
            };
            Assert.Throws<StyleException>(() => style.GetHashCode());
            style = new Style
            {
                CurrentCellXf = null
            };
            Assert.Throws<StyleException>(() => style.GetHashCode());
            style = new Style
            {
                CurrentFill = null
            };
            Assert.Throws<StyleException>(() => style.GetHashCode());
            style = new Style
            {
                CurrentFont = null
            };
            Assert.Throws<StyleException>(() => style.GetHashCode());
            style = new Style
            {
                CurrentNumberFormat = null
            };
            Assert.Throws<StyleException>(() => style.GetHashCode());
        }

        [Fact(DisplayName = "Test of the failing Copy function on a invalid style component (null instance)")]
        public void CopyFailTest()
        {
            Style style = new Style
            {
                CurrentBorder = null
            };
            Assert.Throws<StyleException>(() => style.Copy());
            style = new Style
            {
                CurrentCellXf = null
            };
            Assert.Throws<StyleException>(() => style.Copy());
            style = new Style
            {
                CurrentFill = null
            };
            Assert.Throws<StyleException>(() => style.Copy());
            style = new Style
            {
                CurrentFont = null
            };
            Assert.Throws<StyleException>(() => style.Copy());
            style = new Style
            {
                CurrentNumberFormat = null
            };
            Assert.Throws<StyleException>(() => style.Copy());
        }

        // For code coverage
        [Fact(DisplayName = "Test of the ToString function")]
        public void ToStringTest()
        {
            Style style = new Style();
            string s1 = style.ToString();
            style.Name = "Test1";
            string s2 = style.ToString();
            style.Name = null;
            string s3 = style.ToString();
            string hashCode = style.GetHashCode().ToString();
            Assert.NotEqual(s1, s2);
            Assert.Contains("Test1", s2);
            Assert.Contains(hashCode, s3);
        }

    }
}
