using System;
using NanoXLSX.Exceptions;
using NanoXLSX.Styles;
using NanoXLSX.Test.Core.Utils;
using Xunit;
using static NanoXLSX.Styles.NumberFormat;
using FormatException = NanoXLSX.Exceptions.FormatException;

namespace NanoXLSX.Test.Core.StyleTest
{
    // Ensure that these tests are executed sequentially, since static repository methods may be called 
    [Collection(nameof(SequentialCollection))]
    public class NumberFormatTest
    {

        private readonly NumberFormat exampleStyle;

        public NumberFormatTest()
        {
            exampleStyle = new NumberFormat();
            exampleStyle.CustomFormatCode = "#.###";
            exampleStyle.Number = FormatNumber.format_10;
            exampleStyle.CustomFormatID = 170;
        }

        [Theory(DisplayName = "Test of the get and set function of the FormatNumber property")]
        [InlineData(FormatNumber.none)]
        [InlineData(FormatNumber.format_1)]
        [InlineData(FormatNumber.format_2)]
        [InlineData(FormatNumber.format_3)]
        [InlineData(FormatNumber.format_4)]
        [InlineData(FormatNumber.format_5)]
        [InlineData(FormatNumber.format_6)]
        [InlineData(FormatNumber.format_7)]
        [InlineData(FormatNumber.format_8)]
        [InlineData(FormatNumber.format_9)]
        [InlineData(FormatNumber.format_10)]
        [InlineData(FormatNumber.format_11)]
        [InlineData(FormatNumber.format_12)]
        [InlineData(FormatNumber.format_13)]
        [InlineData(FormatNumber.format_14)]
        [InlineData(FormatNumber.format_15)]
        [InlineData(FormatNumber.format_16)]
        [InlineData(FormatNumber.format_17)]
        [InlineData(FormatNumber.format_18)]
        [InlineData(FormatNumber.format_19)]
        [InlineData(FormatNumber.format_20)]
        [InlineData(FormatNumber.format_21)]
        [InlineData(FormatNumber.format_22)]
        [InlineData(FormatNumber.format_37)]
        [InlineData(FormatNumber.format_38)]
        [InlineData(FormatNumber.format_39)]
        [InlineData(FormatNumber.format_40)]
        [InlineData(FormatNumber.format_45)]
        [InlineData(FormatNumber.format_46)]
        [InlineData(FormatNumber.format_47)]
        [InlineData(FormatNumber.format_48)]
        [InlineData(FormatNumber.format_49)]
        [InlineData(FormatNumber.custom)]
        public void FormatNumberTest(FormatNumber number)
        {
            NumberFormat numberFormat = new NumberFormat();
            Assert.Equal(NumberFormat.DEFAULT_NUMBER, numberFormat.Number); // default is none
            numberFormat.Number = number;
            Assert.Equal(number, numberFormat.Number);
        }

        [Theory(DisplayName = "Test of the get and set function of the CustomFormatCode property")]
        [InlineData("//")]
        [InlineData("#.###")]
        public void CustomFormatCodeTest(string value)
        {
            NumberFormat numberFormat = new NumberFormat();
            Assert.Equal(string.Empty, numberFormat.CustomFormatCode);
            numberFormat.CustomFormatCode = value;
            Assert.Equal(value, numberFormat.CustomFormatCode);
        }

        [Theory(DisplayName = "Test of the failing set function of the CustomFormatCode property on invalid values")]
        [InlineData("")]
        [InlineData(null)]
        public void CustomFormatCodeFailTest(string value)
        {
            NumberFormat numberFormat = new NumberFormat();
            Exception ex = Assert.Throws<FormatException>(() => numberFormat.CustomFormatCode = value);
            Assert.Equal(typeof(FormatException), ex.GetType());
        }

        [Theory(DisplayName = "Test of the get and set function of the CustomFormatID property")]
        [InlineData(164)]
        [InlineData(200)]
        public void CustomFormatIDTest(int value)
        {
            NumberFormat numberFormat = new NumberFormat();
            Assert.Equal(164, numberFormat.CustomFormatID);
            numberFormat.CustomFormatID = value;
            Assert.Equal(value, numberFormat.CustomFormatID);
        }

        [Theory(DisplayName = "Test of the failing set function of the CustomFormatID property (invalid values)")]
        [InlineData(163)]
        [InlineData(0)]
        [InlineData(-100)]
        public void CustomFormatIDFailTest(int value)
        {
            NumberFormat numberFormat = new NumberFormat();
            Exception ex = Assert.Throws<StyleException>(() => numberFormat.CustomFormatID = value);
            Assert.Equal(typeof(StyleException), ex.GetType());
        }

        [Theory(DisplayName = "Test of the get function of the IsCustomFormat property")]
        [InlineData(FormatNumber.none, false)]
        [InlineData(FormatNumber.format_10, false)]
        [InlineData(FormatNumber.custom, true)]
        public void IsCustomFormatTest(FormatNumber number, bool expectedResult)
        {
            NumberFormat numberFormat = new NumberFormat();
            Assert.False(numberFormat.IsCustomFormat);
            numberFormat.Number = number;
            Assert.Equal(expectedResult, numberFormat.IsCustomFormat);
        }

        [Theory(DisplayName = "Test of the IsDateFormat method")]
        [InlineData(FormatNumber.none, false)]
        [InlineData(FormatNumber.format_1, false)]
        [InlineData(FormatNumber.format_2, false)]
        [InlineData(FormatNumber.format_3, false)]
        [InlineData(FormatNumber.format_4, false)]
        [InlineData(FormatNumber.format_5, false)]
        [InlineData(FormatNumber.format_6, false)]
        [InlineData(FormatNumber.format_7, false)]
        [InlineData(FormatNumber.format_8, false)]
        [InlineData(FormatNumber.format_9, false)]
        [InlineData(FormatNumber.format_10, false)]
        [InlineData(FormatNumber.format_11, false)]
        [InlineData(FormatNumber.format_12, false)]
        [InlineData(FormatNumber.format_13, false)]
        [InlineData(FormatNumber.format_14, true)]
        [InlineData(FormatNumber.format_15, true)]
        [InlineData(FormatNumber.format_16, true)]
        [InlineData(FormatNumber.format_17, true)]
        [InlineData(FormatNumber.format_18, false)]
        [InlineData(FormatNumber.format_19, false)]
        [InlineData(FormatNumber.format_20, false)]
        [InlineData(FormatNumber.format_21, false)]
        [InlineData(FormatNumber.format_22, true)]
        [InlineData(FormatNumber.format_37, false)]
        [InlineData(FormatNumber.format_38, false)]
        [InlineData(FormatNumber.format_39, false)]
        [InlineData(FormatNumber.format_40, false)]
        [InlineData(FormatNumber.format_45, false)]
        [InlineData(FormatNumber.format_46, false)]
        [InlineData(FormatNumber.format_47, false)]
        [InlineData(FormatNumber.format_48, false)]
        [InlineData(FormatNumber.format_49, false)]
        [InlineData(FormatNumber.custom, false)]
        public void IsDateFormatTest(FormatNumber number, bool expectedDate)
        {
            Assert.Equal(expectedDate, NumberFormat.IsDateFormat(number));
        }

        [Theory(DisplayName = "Test of the IsTimeFormat method")]
        [InlineData(FormatNumber.none, false)]
        [InlineData(FormatNumber.format_1, false)]
        [InlineData(FormatNumber.format_2, false)]
        [InlineData(FormatNumber.format_3, false)]
        [InlineData(FormatNumber.format_4, false)]
        [InlineData(FormatNumber.format_5, false)]
        [InlineData(FormatNumber.format_6, false)]
        [InlineData(FormatNumber.format_7, false)]
        [InlineData(FormatNumber.format_8, false)]
        [InlineData(FormatNumber.format_9, false)]
        [InlineData(FormatNumber.format_10, false)]
        [InlineData(FormatNumber.format_11, false)]
        [InlineData(FormatNumber.format_12, false)]
        [InlineData(FormatNumber.format_13, false)]
        [InlineData(FormatNumber.format_14, false)]
        [InlineData(FormatNumber.format_15, false)]
        [InlineData(FormatNumber.format_16, false)]
        [InlineData(FormatNumber.format_17, false)]
        [InlineData(FormatNumber.format_18, true)]
        [InlineData(FormatNumber.format_19, true)]
        [InlineData(FormatNumber.format_20, true)]
        [InlineData(FormatNumber.format_21, true)]
        [InlineData(FormatNumber.format_22, false)]
        [InlineData(FormatNumber.format_37, false)]
        [InlineData(FormatNumber.format_38, false)]
        [InlineData(FormatNumber.format_39, false)]
        [InlineData(FormatNumber.format_40, false)]
        [InlineData(FormatNumber.format_45, true)]
        [InlineData(FormatNumber.format_46, true)]
        [InlineData(FormatNumber.format_47, true)]
        [InlineData(FormatNumber.format_48, false)]
        [InlineData(FormatNumber.format_49, false)]
        [InlineData(FormatNumber.custom, false)]
        public void IsTimeFormatTest(FormatNumber number, bool expectedTime)
        {
            Assert.Equal(expectedTime, NumberFormat.IsTimeFormat(number));
        }

        [Theory(DisplayName = "Test of the TryParseFormatNumber method")]
        [InlineData(0, FormatRange.defined_format, FormatNumber.none)]
        [InlineData(-1, FormatRange.invalid, FormatNumber.none)]
        [InlineData(22, FormatRange.defined_format, FormatNumber.format_22)]
        [InlineData(23, FormatRange.undefined, FormatNumber.none)]
        [InlineData(163, FormatRange.undefined, FormatNumber.none)]
        [InlineData(164, FormatRange.defined_format, FormatNumber.custom)]
        [InlineData(165, FormatRange.custom_format, FormatNumber.custom)]
        [InlineData(700, FormatRange.custom_format, FormatNumber.custom)]
        public void TryParseFormatNumberTest(int givenNumber, FormatRange expectedRange, FormatNumber expectedFormatNumber)
        {
            FormatNumber number;
            FormatRange range = NumberFormat.TryParseFormatNumber(givenNumber, out number);
            Assert.Equal(expectedRange, range);
            Assert.Equal(expectedFormatNumber, number);
        }

        [Fact(DisplayName = "Test of the Equals method")]
        public void EqualsTest()
        {
            NumberFormat style2 = (NumberFormat)exampleStyle.Copy();
            Assert.True(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of Number)")]
        public void EqualsTest2()
        {
            NumberFormat style2 = (NumberFormat)exampleStyle.Copy();
            style2.Number = FormatNumber.format_15;
            Assert.False(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of CustomFormatCode)")]
        public void EqualsTest2b()
        {
            NumberFormat style2 = (NumberFormat)exampleStyle.Copy();
            style2.CustomFormatCode = "hh-mm-ss";
            Assert.False(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of CustomFormatID)")]
        public void EqualsTest2c()
        {
            NumberFormat style2 = (NumberFormat)exampleStyle.Copy();
            style2.CustomFormatID = 180;
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
            NumberFormat copy = (NumberFormat)exampleStyle.Copy();
            Assert.False(copy.Equals(origin));
        }

        [Fact(DisplayName = "Test of the GetHashCode method (equality of two identical objects)")]
        public void GetHashCodeTest()
        {
            NumberFormat copy = (NumberFormat)exampleStyle.Copy();
            copy.InternalID = 99;  // Should not influence
            Assert.Equal(exampleStyle.GetHashCode(), copy.GetHashCode());
        }

        [Fact(DisplayName = "Test of the GetHashCode method (inequality of two different objects)")]
        public void GetHashCodeTest2()
        {
            NumberFormat copy = (NumberFormat)exampleStyle.Copy();
            copy.Number = FormatNumber.format_14;
            Assert.NotEqual(exampleStyle.GetHashCode(), copy.GetHashCode());
        }

        [Fact(DisplayName = "Test of the constant of the default custom format start number")]
        public void DefaultFontNameTest()
        {
            Assert.Equal(164, NumberFormat.CUSTOMFORMAT_START_NUMBER); // Expected 164
        }

        [Fact(DisplayName = "Test of the CompareTo method")]
        public void CompareToTest()
        {
            NumberFormat numberFormat = new NumberFormat();
            NumberFormat other = new NumberFormat();
            numberFormat.InternalID = null;
            other.InternalID = null;
            Assert.Equal(-1, numberFormat.CompareTo(other));
            numberFormat.InternalID = 5;
            Assert.Equal(1, numberFormat.CompareTo(other));
            Assert.Equal(1, numberFormat.CompareTo(null));
            other.InternalID = 5;
            Assert.Equal(0, numberFormat.CompareTo(other));
            other.InternalID = 4;
            Assert.Equal(1, numberFormat.CompareTo(other));
            other.InternalID = 6;
            Assert.Equal(-1, numberFormat.CompareTo(other));
        }

        // For code coverage
        [Fact(DisplayName = "Test of the ToString function")]
        public void ToStringTest()
        {
            NumberFormat numberFormat = new NumberFormat();
            string s1 = numberFormat.ToString();
            numberFormat.Number = FormatNumber.format_11;
            Assert.NotEqual(s1, numberFormat.ToString()); // An explicit value comparison is probably not sensible
        }

    }
}
