using System;
using NanoXLSX.Styles;
using NanoXLSX.Test.Writer_Reader.Utils;
using Xunit;
using static NanoXLSX.Styles.NumberFormat;

namespace NanoXLSX.Test.Writer_Reader.Styles
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
            var style = new Style();
            style.CurrentNumberFormat.CustomFormatID = styleValue;
            style.CurrentNumberFormat.Number = FormatNumber.Custom; // Mandatory
            style.CurrentNumberFormat.CustomFormatCode = "#.##"; // Mandatory
            var cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentNumberFormat.CustomFormatID);
        }

        [Theory(DisplayName = "Test of the failing save attempt of 'customFormatID' value when writing and reading a NumberFormat style with missing CustomFormatCode")]
        [InlineData(164, "test")]
        [InlineData(165, 0.5f)]
        [InlineData(200, 22)]
        [InlineData(2000, true)]
        public void CustomFormatIDFormatFailTest(int styleValue, object value)
        {
            var style = new Style();
            style.CurrentNumberFormat.CustomFormatID = styleValue;
            style.CurrentNumberFormat.Number = FormatNumber.Custom; // Mandatory
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
            var style = new Style();
            style.CurrentNumberFormat.CustomFormatCode = styleValue;
            style.CurrentNumberFormat.Number = FormatNumber.Custom; // Mandatory
            var cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentNumberFormat.CustomFormatCode);
            Assert.Equal(FormatNumber.Custom, cell.CellStyle.CurrentNumberFormat.Number);
            Assert.True(cell.CellStyle.CurrentNumberFormat.IsCustomFormat);
        }

        [Theory(DisplayName = "Test of the 'formatNumber' value when writing and reading a NumberFormat style")]
        [InlineData(FormatNumber.Format1, "test")]
        [InlineData(FormatNumber.Format2, 0.5f)]
        [InlineData(FormatNumber.Format3, 22)]
        [InlineData(FormatNumber.Format4, true)]
        [InlineData(FormatNumber.Format5, "")]
        [InlineData(FormatNumber.Format6, -1)]
        [InlineData(FormatNumber.Format7, -22.222f)]
        [InlineData(FormatNumber.Format8, false)]
        [InlineData(FormatNumber.Format9, 0)]
        [InlineData(FormatNumber.Format10, "Æ")]
        [InlineData(FormatNumber.Format11, "test")]
        [InlineData(FormatNumber.Format12, 0.5f)]
        [InlineData(FormatNumber.Format13, 22)]
        [InlineData(FormatNumber.Format14, true)]
        [InlineData(FormatNumber.Format15, "")]
        [InlineData(FormatNumber.Format16, -1)]
        [InlineData(FormatNumber.Format17, -22.222f)]
        [InlineData(FormatNumber.Format18, false)]
        [InlineData(FormatNumber.Format19, "noDate")]
        [InlineData(FormatNumber.Format20, "Æ")]
        [InlineData(FormatNumber.Format21, "test")]
        [InlineData(FormatNumber.Format22, "noDate")]
        [InlineData(FormatNumber.Format37, 22)]
        [InlineData(FormatNumber.Format38, true)]
        [InlineData(FormatNumber.Format39, "")]
        [InlineData(FormatNumber.Format40, -1)]
        [InlineData(FormatNumber.Format45, -22.222f)]
        [InlineData(FormatNumber.Format46, false)]
        [InlineData(FormatNumber.Format47, "noDate")]
        [InlineData(FormatNumber.Format48, "Æ")]
        [InlineData(FormatNumber.Format49, "test")]
        public void NumberFormatTest(FormatNumber styleValue, object value)
        {
            var style = new Style();
            style.CurrentNumberFormat.Number = styleValue;
            var cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentNumberFormat.Number);
        }

        [Fact(DisplayName = "Test of the 'formatNumber' value when writing and reading a custom NumberFormat style")]
        public void NumberFormatTest1b()
        {
            var style = new Style();
            style.CurrentNumberFormat.Number = FormatNumber.Custom;
            style.CurrentNumberFormat.CustomFormatCode = "#.##";
            var cell = TestUtils.SaveAndReadStyledCell(0.5f, style, "A1");
            Assert.Equal(FormatNumber.Custom, cell.CellStyle.CurrentNumberFormat.Number);
        }

        [Theory(DisplayName = "Test of the 'formatNumber' value with date formats when writing and reading a NumberFormat style")]
        [InlineData(FormatNumber.Format14, 1000, "26.09.1902")]
        [InlineData(FormatNumber.Format15, 1000, "26.09.1902")]
        [InlineData(FormatNumber.Format16, 1000, "26.09.1902")]
        [InlineData(FormatNumber.Format17, 1000, "26.09.1902")]
        [InlineData(FormatNumber.Format22, 1000, "26.09.1902")]
        public void NumberNumberFormatTest2(FormatNumber styleValue, int value, string expected)
        {
            var expectedValue = DateTime.ParseExact(expected, "dd.MM.yyyy", System.Globalization.CultureInfo.InvariantCulture);
            var style = new Style();
            style.CurrentNumberFormat.Number = styleValue;
            var cell = TestUtils.SaveAndReadStyledCell(value, expectedValue, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentNumberFormat.Number);
            Assert.Equal(expectedValue, cell.Value);
        }

        [Theory(DisplayName = "Test of the 'formatNumber' value with time formats when writing and reading a NumberFormat style")]
        [InlineData(FormatNumber.Format19, 0.5, "12:00:00")]
        [InlineData(FormatNumber.Format20, 0.5, "12:00:00")]
        [InlineData(FormatNumber.Format21, 0.5, "12:00:00")]
        [InlineData(FormatNumber.Format45, 0.5, "12:00:00")]
        [InlineData(FormatNumber.Format46, 0.5, "12:00:00")]
        [InlineData(FormatNumber.Format47, 0.5, "12:00:00")]
        public void NumberNumberFormatTest3(FormatNumber styleValue, float value, string expected)
        {
            var expectedValue = TimeSpan.ParseExact(expected, "hh\\:mm\\:ss", System.Globalization.CultureInfo.InvariantCulture);
            var style = new Style();
            style.CurrentNumberFormat.Number = styleValue;
            var cell = TestUtils.SaveAndReadStyledCell(value, expectedValue, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentNumberFormat.Number);
            Assert.Equal(expectedValue, cell.Value);
        }

    }
}
