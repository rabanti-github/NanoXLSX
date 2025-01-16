using System;
using NanoXLSX;
using NanoXLSX.Styles;
using NanoXLSX.Test.Writer_Reader.Utils;
using Xunit;
using static NanoXLSX.Styles.NumberFormat;

namespace NanoXLSX.Test.Writer_Reader.StyleTest
{
    public class NumberFormatWriteReadTest
    {

        [Theory(DisplayName = "Test of the 'customFormatID' value when writing and reading a NumberFormat style")]
        [InlineData(164, "test")]
        [InlineData(165, 0.5f)]
        [InlineData(200, 22)]
        [InlineData(2000, true)]
        public void CustomFormatIDFormatTest(int styleValue, object value)
        {
            Style style = new Style();
            style.CurrentNumberFormat.CustomFormatID = styleValue;
            style.CurrentNumberFormat.Number = FormatNumber.custom; // Mandatory
            style.CurrentNumberFormat.CustomFormatCode = "#.##"; // Mandatory
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentNumberFormat.CustomFormatID);
        }

        [Theory(DisplayName = "Test of the failing save attempt of 'customFormatID' value when writing and reading a NumberFormat style with missing CustomFormatCode")]
        [InlineData(164, "test")]
        [InlineData(165, 0.5f)]
        [InlineData(200, 22)]
        [InlineData(2000, true)]
        public void CustomFormatIDFormatFailTest(int styleValue, object value)
        {
            Style style = new Style();
            style.CurrentNumberFormat.CustomFormatID = styleValue;
            style.CurrentNumberFormat.Number = FormatNumber.custom; // Mandatory
            Assert.ThrowsAny<Exception>(() => TestUtils.SaveAndReadStyledCell(value, style, "A1"));
        }

        [Theory(DisplayName = "Test of the 'customFormatCode' value when writing and reading a NumberFormat style")]
        [InlineData("#", "test")]
        [InlineData("\\", 0.5f)]
        [InlineData("\\\\", "")]
        [InlineData(" \\.\\ ", false)]
        [InlineData(" ", 22)]
        [InlineData("ABCDE", true)]
        public void CustomFormatCodeNumberFormatTest(string styleValue, object value)
        {
            Style style = new Style();
            style.CurrentNumberFormat.CustomFormatCode = styleValue;
            style.CurrentNumberFormat.Number = FormatNumber.custom; // Mandatory
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentNumberFormat.CustomFormatCode);
            Assert.Equal(FormatNumber.custom, cell.CellStyle.CurrentNumberFormat.Number);
            Assert.True(cell.CellStyle.CurrentNumberFormat.IsCustomFormat);
        }

        [Theory(DisplayName = "Test of the 'formatNumber' value when writing and reading a NumberFormat style")]
        [InlineData(FormatNumber.format_1, "test")]
        [InlineData(FormatNumber.format_2, 0.5f)]
        [InlineData(FormatNumber.format_3, 22)]
        [InlineData(FormatNumber.format_4, true)]
        [InlineData(FormatNumber.format_5, "")]
        [InlineData(FormatNumber.format_6, -1)]
        [InlineData(FormatNumber.format_7, -22.222f)]
        [InlineData(FormatNumber.format_8, false)]
        [InlineData(FormatNumber.format_9, 0)]
        [InlineData(FormatNumber.format_10, "Æ")]
        [InlineData(FormatNumber.format_11, "test")]
        [InlineData(FormatNumber.format_12, 0.5f)]
        [InlineData(FormatNumber.format_13, 22)]
        [InlineData(FormatNumber.format_14, true)]
        [InlineData(FormatNumber.format_15, "")]
        [InlineData(FormatNumber.format_16, -1)]
        [InlineData(FormatNumber.format_17, -22.222f)]
        [InlineData(FormatNumber.format_18, false)]
        [InlineData(FormatNumber.format_19, "noDate")]
        [InlineData(FormatNumber.format_20, "Æ")]
        [InlineData(FormatNumber.format_21, "test")]
        [InlineData(FormatNumber.format_22, "noDate")]
        [InlineData(FormatNumber.format_37, 22)]
        [InlineData(FormatNumber.format_38, true)]
        [InlineData(FormatNumber.format_39, "")]
        [InlineData(FormatNumber.format_40, -1)]
        [InlineData(FormatNumber.format_45, -22.222f)]
        [InlineData(FormatNumber.format_46, false)]
        [InlineData(FormatNumber.format_47, "noDate")]
        [InlineData(FormatNumber.format_48, "Æ")]
        [InlineData(FormatNumber.format_49, "test")]
        public void NumberFormatTest(FormatNumber styleValue, object value)
        {
            Style style = new Style();
            style.CurrentNumberFormat.Number = styleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentNumberFormat.Number);
        }

        [Fact(DisplayName = "Test of the 'formatNumber' value when writing and reading a custom NumberFormat style")]
        public void NumberFormatTest1b()
        {
            Style style = new Style();
            style.CurrentNumberFormat.Number = FormatNumber.custom;
            style.CurrentNumberFormat.CustomFormatCode = "#.##";
            Cell cell = TestUtils.SaveAndReadStyledCell(0.5f, style, "A1");
            Assert.Equal(FormatNumber.custom, cell.CellStyle.CurrentNumberFormat.Number);
        }

        [Theory(DisplayName = "Test of the 'formatNumber' value with date formats when writing and reading a NumberFormat style")]
        [InlineData(FormatNumber.format_14, 1000, "26.09.1902")]
        [InlineData(FormatNumber.format_15, 1000, "26.09.1902")]
        [InlineData(FormatNumber.format_16, 1000, "26.09.1902")]
        [InlineData(FormatNumber.format_17, 1000, "26.09.1902")]
        [InlineData(FormatNumber.format_22, 1000, "26.09.1902")]
        public void NumberNumberFormatTest2(FormatNumber styleValue, int value, string expected)
        {
            DateTime expectedValue = DateTime.ParseExact(expected, "dd.MM.yyyy", System.Globalization.CultureInfo.InvariantCulture);
            Style style = new Style();
            style.CurrentNumberFormat.Number = styleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, expectedValue, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentNumberFormat.Number);
            Assert.Equal(expectedValue, cell.Value);
        }

        [Theory(DisplayName = "Test of the 'formatNumber' value with time formats when writing and reading a NumberFormat style")]
        [InlineData(FormatNumber.format_19, 0.5, "12:00:00")]
        [InlineData(FormatNumber.format_20, 0.5, "12:00:00")]
        [InlineData(FormatNumber.format_21, 0.5, "12:00:00")]
        [InlineData(FormatNumber.format_45, 0.5, "12:00:00")]
        [InlineData(FormatNumber.format_46, 0.5, "12:00:00")]
        [InlineData(FormatNumber.format_47, 0.5, "12:00:00")]
        public void NumberNumberFormatTest3(FormatNumber styleValue, float value, string expected)
        {
            TimeSpan expectedValue = TimeSpan.ParseExact(expected, "hh\\:mm\\:ss", System.Globalization.CultureInfo.InvariantCulture);
            Style style = new Style();
            style.CurrentNumberFormat.Number = styleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, expectedValue, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentNumberFormat.Number);
            Assert.Equal(expectedValue, cell.Value);
        }

    }
}
