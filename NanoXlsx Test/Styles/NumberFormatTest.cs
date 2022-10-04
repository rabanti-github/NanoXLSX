using NanoXLSX.Exceptions;
using NanoXLSX.Styles;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;
using FormatException = NanoXLSX.Exceptions.FormatException;

namespace NanoXLSX_Test.Styles
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
            exampleStyle.Number = NumberFormat.FormatNumber.format_10;
            exampleStyle.CustomFormatID = 170;
        }

        [Theory(DisplayName = "Test of the get and set function of the FormatNumber property")]
        [InlineData(NumberFormat.FormatNumber.none)]
        [InlineData(NumberFormat.FormatNumber.format_1)]
        [InlineData(NumberFormat.FormatNumber.format_2)]
        [InlineData(NumberFormat.FormatNumber.format_3)]
        [InlineData(NumberFormat.FormatNumber.format_4)]
        [InlineData(NumberFormat.FormatNumber.format_5)]
        [InlineData(NumberFormat.FormatNumber.format_6)]
        [InlineData(NumberFormat.FormatNumber.format_7)]
        [InlineData(NumberFormat.FormatNumber.format_8)]
        [InlineData(NumberFormat.FormatNumber.format_9)]
        [InlineData(NumberFormat.FormatNumber.format_10)]
        [InlineData(NumberFormat.FormatNumber.format_11)]
        [InlineData(NumberFormat.FormatNumber.format_12)]
        [InlineData(NumberFormat.FormatNumber.format_13)]
        [InlineData(NumberFormat.FormatNumber.format_14)]
        [InlineData(NumberFormat.FormatNumber.format_15)]
        [InlineData(NumberFormat.FormatNumber.format_16)]
        [InlineData(NumberFormat.FormatNumber.format_17)]
        [InlineData(NumberFormat.FormatNumber.format_18)]
        [InlineData(NumberFormat.FormatNumber.format_19)]
        [InlineData(NumberFormat.FormatNumber.format_20)]
        [InlineData(NumberFormat.FormatNumber.format_21)]
        [InlineData(NumberFormat.FormatNumber.format_22)]
        [InlineData(NumberFormat.FormatNumber.format_37)]
        [InlineData(NumberFormat.FormatNumber.format_38)]
        [InlineData(NumberFormat.FormatNumber.format_39)]
        [InlineData(NumberFormat.FormatNumber.format_40)]
        [InlineData(NumberFormat.FormatNumber.format_45)]
        [InlineData(NumberFormat.FormatNumber.format_46)]
        [InlineData(NumberFormat.FormatNumber.format_47)]
        [InlineData(NumberFormat.FormatNumber.format_48)]
        [InlineData(NumberFormat.FormatNumber.format_49)]
        [InlineData(NumberFormat.FormatNumber.custom)]
        public void FormatNumberTest(NumberFormat.FormatNumber number)
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
        [InlineData(NumberFormat.FormatNumber.none, false)]
        [InlineData(NumberFormat.FormatNumber.format_10, false)]
        [InlineData(NumberFormat.FormatNumber.custom, true)]
        public void IsCustomFormatTest(NumberFormat.FormatNumber number, bool expectedResult)
        {
            NumberFormat numberFormat = new NumberFormat();
            Assert.False(numberFormat.IsCustomFormat);
            numberFormat.Number = number;
            Assert.Equal(expectedResult, numberFormat.IsCustomFormat);
        }

        [Theory(DisplayName = "Test of the IsDateFormat method")]
        [InlineData(NumberFormat.FormatNumber.none, false)]
        [InlineData(NumberFormat.FormatNumber.format_1, false)]
        [InlineData(NumberFormat.FormatNumber.format_2, false)]
        [InlineData(NumberFormat.FormatNumber.format_3, false)]
        [InlineData(NumberFormat.FormatNumber.format_4, false)]
        [InlineData(NumberFormat.FormatNumber.format_5, false)]
        [InlineData(NumberFormat.FormatNumber.format_6, false)]
        [InlineData(NumberFormat.FormatNumber.format_7, false)]
        [InlineData(NumberFormat.FormatNumber.format_8, false)]
        [InlineData(NumberFormat.FormatNumber.format_9, false)]
        [InlineData(NumberFormat.FormatNumber.format_10, false)]
        [InlineData(NumberFormat.FormatNumber.format_11, false)]
        [InlineData(NumberFormat.FormatNumber.format_12, false)]
        [InlineData(NumberFormat.FormatNumber.format_13, false)]
        [InlineData(NumberFormat.FormatNumber.format_14, true)]
        [InlineData(NumberFormat.FormatNumber.format_15, true)]
        [InlineData(NumberFormat.FormatNumber.format_16, true)]
        [InlineData(NumberFormat.FormatNumber.format_17, true)]
        [InlineData(NumberFormat.FormatNumber.format_18, false)]
        [InlineData(NumberFormat.FormatNumber.format_19, false)]
        [InlineData(NumberFormat.FormatNumber.format_20, false)]
        [InlineData(NumberFormat.FormatNumber.format_21, false)]
        [InlineData(NumberFormat.FormatNumber.format_22, true)]
        [InlineData(NumberFormat.FormatNumber.format_37, false)]
        [InlineData(NumberFormat.FormatNumber.format_38, false)]
        [InlineData(NumberFormat.FormatNumber.format_39, false)]
        [InlineData(NumberFormat.FormatNumber.format_40, false)]
        [InlineData(NumberFormat.FormatNumber.format_45, false)]
        [InlineData(NumberFormat.FormatNumber.format_46, false)]
        [InlineData(NumberFormat.FormatNumber.format_47, false)]
        [InlineData(NumberFormat.FormatNumber.format_48, false)]
        [InlineData(NumberFormat.FormatNumber.format_49, false)]
        [InlineData(NumberFormat.FormatNumber.custom, false)]
        public void IsDateFormatTest(NumberFormat.FormatNumber number, bool expectedDate)
        {
            Assert.Equal(expectedDate, NumberFormat.IsDateFormat(number));
        }

        [Theory(DisplayName = "Test of the IsTimeFormat method")]
        [InlineData(NumberFormat.FormatNumber.none, false)]
        [InlineData(NumberFormat.FormatNumber.format_1, false)]
        [InlineData(NumberFormat.FormatNumber.format_2, false)]
        [InlineData(NumberFormat.FormatNumber.format_3, false)]
        [InlineData(NumberFormat.FormatNumber.format_4, false)]
        [InlineData(NumberFormat.FormatNumber.format_5, false)]
        [InlineData(NumberFormat.FormatNumber.format_6, false)]
        [InlineData(NumberFormat.FormatNumber.format_7, false)]
        [InlineData(NumberFormat.FormatNumber.format_8, false)]
        [InlineData(NumberFormat.FormatNumber.format_9, false)]
        [InlineData(NumberFormat.FormatNumber.format_10, false)]
        [InlineData(NumberFormat.FormatNumber.format_11, false)]
        [InlineData(NumberFormat.FormatNumber.format_12, false)]
        [InlineData(NumberFormat.FormatNumber.format_13, false)]
        [InlineData(NumberFormat.FormatNumber.format_14, false)]
        [InlineData(NumberFormat.FormatNumber.format_15, false)]
        [InlineData(NumberFormat.FormatNumber.format_16, false)]
        [InlineData(NumberFormat.FormatNumber.format_17, false)]
        [InlineData(NumberFormat.FormatNumber.format_18, true)]
        [InlineData(NumberFormat.FormatNumber.format_19, true)]
        [InlineData(NumberFormat.FormatNumber.format_20, true)]
        [InlineData(NumberFormat.FormatNumber.format_21, true)]
        [InlineData(NumberFormat.FormatNumber.format_22, false)]
        [InlineData(NumberFormat.FormatNumber.format_37, false)]
        [InlineData(NumberFormat.FormatNumber.format_38, false)]
        [InlineData(NumberFormat.FormatNumber.format_39, false)]
        [InlineData(NumberFormat.FormatNumber.format_40, false)]
        [InlineData(NumberFormat.FormatNumber.format_45, true)]
        [InlineData(NumberFormat.FormatNumber.format_46, true)]
        [InlineData(NumberFormat.FormatNumber.format_47, true)]
        [InlineData(NumberFormat.FormatNumber.format_48, false)]
        [InlineData(NumberFormat.FormatNumber.format_49, false)]
        [InlineData(NumberFormat.FormatNumber.custom, false)]
        public void IsTimeFormatTest(NumberFormat.FormatNumber number, bool expectedTime)
        {
            Assert.Equal(expectedTime, NumberFormat.IsTimeFormat(number));
        }

        [Theory(DisplayName = "Test of the TryParseFormatNumber method")]
        [InlineData(0, NumberFormat.FormatRange.defined_format, NumberFormat.FormatNumber.none)]
        [InlineData(-1, NumberFormat.FormatRange.invalid, NumberFormat.FormatNumber.none)]
        [InlineData(22, NumberFormat.FormatRange.defined_format, NumberFormat.FormatNumber.format_22)]
        [InlineData(23, NumberFormat.FormatRange.undefined, NumberFormat.FormatNumber.none)]
        [InlineData(163, NumberFormat.FormatRange.undefined, NumberFormat.FormatNumber.none)]
        [InlineData(164, NumberFormat.FormatRange.defined_format, NumberFormat.FormatNumber.custom)]
        [InlineData(165, NumberFormat.FormatRange.custom_format, NumberFormat.FormatNumber.custom)]
        [InlineData(700, NumberFormat.FormatRange.custom_format, NumberFormat.FormatNumber.custom)]
        public void TryParseFormatNumberTest(int givenNumber, NumberFormat.FormatRange expectedRange, NumberFormat.FormatNumber expectedFormatNumber)
        {
            NumberFormat.FormatNumber number;
            NumberFormat.FormatRange range = NumberFormat.TryParseFormatNumber(givenNumber, out number);
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
            style2.Number = NumberFormat.FormatNumber.format_15;
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
            copy.Number = NumberFormat.FormatNumber.format_14;
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
            numberFormat.Number = NumberFormat.FormatNumber.format_11;
            Assert.NotEqual(s1, numberFormat.ToString()); // An explicit value comparison is probably not sensible
        }

        private static object SequentialCollection()
        {
            throw new NotImplementedException();
        }
    }
}
