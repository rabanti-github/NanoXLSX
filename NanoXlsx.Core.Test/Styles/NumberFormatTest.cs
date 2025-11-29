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
            exampleStyle = new NumberFormat
            {
                CustomFormatCode = "#.###",
                Number = FormatNumber.Format10,
                CustomFormatID = 170
            };
        }

        [Theory(DisplayName = "Test of the get and set function of the FormatNumber property")]
        [InlineData(FormatNumber.None)]
        [InlineData(FormatNumber.Format1)]
        [InlineData(FormatNumber.Format2)]
        [InlineData(FormatNumber.Format3)]
        [InlineData(FormatNumber.Format4)]
        [InlineData(FormatNumber.Format5)]
        [InlineData(FormatNumber.Format6)]
        [InlineData(FormatNumber.Format7)]
        [InlineData(FormatNumber.Format8)]
        [InlineData(FormatNumber.Format9)]
        [InlineData(FormatNumber.Format10)]
        [InlineData(FormatNumber.Format11)]
        [InlineData(FormatNumber.Format12)]
        [InlineData(FormatNumber.Format13)]
        [InlineData(FormatNumber.Format14)]
        [InlineData(FormatNumber.Format15)]
        [InlineData(FormatNumber.Format16)]
        [InlineData(FormatNumber.Format17)]
        [InlineData(FormatNumber.Format18)]
        [InlineData(FormatNumber.Format19)]
        [InlineData(FormatNumber.Format20)]
        [InlineData(FormatNumber.Format21)]
        [InlineData(FormatNumber.Format22)]
        [InlineData(FormatNumber.Format37)]
        [InlineData(FormatNumber.Format38)]
        [InlineData(FormatNumber.Format39)]
        [InlineData(FormatNumber.Format40)]
        [InlineData(FormatNumber.Format45)]
        [InlineData(FormatNumber.Format46)]
        [InlineData(FormatNumber.Format47)]
        [InlineData(FormatNumber.Format48)]
        [InlineData(FormatNumber.Format49)]
        [InlineData(FormatNumber.Custom)]
        public void FormatNumberTest(FormatNumber number)
        {
            NumberFormat numberFormat = new NumberFormat();
            Assert.Equal(NumberFormat.DefaultNumber, numberFormat.Number); // default is none
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
        [InlineData(FormatNumber.None, false)]
        [InlineData(FormatNumber.Format10, false)]
        [InlineData(FormatNumber.Custom, true)]
        public void IsCustomFormatTest(FormatNumber number, bool expectedResult)
        {
            NumberFormat numberFormat = new NumberFormat();
            Assert.False(numberFormat.IsCustomFormat);
            numberFormat.Number = number;
            Assert.Equal(expectedResult, numberFormat.IsCustomFormat);
        }

        [Theory(DisplayName = "Test of the IsDateFormat method")]
        [InlineData(FormatNumber.None, false)]
        [InlineData(FormatNumber.Format1, false)]
        [InlineData(FormatNumber.Format2, false)]
        [InlineData(FormatNumber.Format3, false)]
        [InlineData(FormatNumber.Format4, false)]
        [InlineData(FormatNumber.Format5, false)]
        [InlineData(FormatNumber.Format6, false)]
        [InlineData(FormatNumber.Format7, false)]
        [InlineData(FormatNumber.Format8, false)]
        [InlineData(FormatNumber.Format9, false)]
        [InlineData(FormatNumber.Format10, false)]
        [InlineData(FormatNumber.Format11, false)]
        [InlineData(FormatNumber.Format12, false)]
        [InlineData(FormatNumber.Format13, false)]
        [InlineData(FormatNumber.Format14, true)]
        [InlineData(FormatNumber.Format15, true)]
        [InlineData(FormatNumber.Format16, true)]
        [InlineData(FormatNumber.Format17, true)]
        [InlineData(FormatNumber.Format18, false)]
        [InlineData(FormatNumber.Format19, false)]
        [InlineData(FormatNumber.Format20, false)]
        [InlineData(FormatNumber.Format21, false)]
        [InlineData(FormatNumber.Format22, true)]
        [InlineData(FormatNumber.Format37, false)]
        [InlineData(FormatNumber.Format38, false)]
        [InlineData(FormatNumber.Format39, false)]
        [InlineData(FormatNumber.Format40, false)]
        [InlineData(FormatNumber.Format45, false)]
        [InlineData(FormatNumber.Format46, false)]
        [InlineData(FormatNumber.Format47, false)]
        [InlineData(FormatNumber.Format48, false)]
        [InlineData(FormatNumber.Format49, false)]
        [InlineData(FormatNumber.Custom, false)]
        public void IsDateFormatTest(FormatNumber number, bool expectedDate)
        {
            Assert.Equal(expectedDate, NumberFormat.IsDateFormat(number));
        }

        [Theory(DisplayName = "Test of the IsTimeFormat method")]
        [InlineData(FormatNumber.None, false)]
        [InlineData(FormatNumber.Format1, false)]
        [InlineData(FormatNumber.Format2, false)]
        [InlineData(FormatNumber.Format3, false)]
        [InlineData(FormatNumber.Format4, false)]
        [InlineData(FormatNumber.Format5, false)]
        [InlineData(FormatNumber.Format6, false)]
        [InlineData(FormatNumber.Format7, false)]
        [InlineData(FormatNumber.Format8, false)]
        [InlineData(FormatNumber.Format9, false)]
        [InlineData(FormatNumber.Format10, false)]
        [InlineData(FormatNumber.Format11, false)]
        [InlineData(FormatNumber.Format12, false)]
        [InlineData(FormatNumber.Format13, false)]
        [InlineData(FormatNumber.Format14, false)]
        [InlineData(FormatNumber.Format15, false)]
        [InlineData(FormatNumber.Format16, false)]
        [InlineData(FormatNumber.Format17, false)]
        [InlineData(FormatNumber.Format18, true)]
        [InlineData(FormatNumber.Format19, true)]
        [InlineData(FormatNumber.Format20, true)]
        [InlineData(FormatNumber.Format21, true)]
        [InlineData(FormatNumber.Format22, false)]
        [InlineData(FormatNumber.Format37, false)]
        [InlineData(FormatNumber.Format38, false)]
        [InlineData(FormatNumber.Format39, false)]
        [InlineData(FormatNumber.Format40, false)]
        [InlineData(FormatNumber.Format45, true)]
        [InlineData(FormatNumber.Format46, true)]
        [InlineData(FormatNumber.Format47, true)]
        [InlineData(FormatNumber.Format48, false)]
        [InlineData(FormatNumber.Format49, false)]
        [InlineData(FormatNumber.Custom, false)]
        public void IsTimeFormatTest(FormatNumber number, bool expectedTime)
        {
            Assert.Equal(expectedTime, NumberFormat.IsTimeFormat(number));
        }

        [Theory(DisplayName = "Test of the TryParseFormatNumber method")]
        [InlineData(0, FormatRange.DefinedFormat, FormatNumber.None)]
        [InlineData(-1, FormatRange.Invalid, FormatNumber.None)]
        [InlineData(22, FormatRange.DefinedFormat, FormatNumber.Format22)]
        [InlineData(23, FormatRange.Undefined, FormatNumber.None)]
        [InlineData(163, FormatRange.Undefined, FormatNumber.None)]
        [InlineData(164, FormatRange.DefinedFormat, FormatNumber.Custom)]
        [InlineData(165, FormatRange.CustomFormat, FormatNumber.Custom)]
        [InlineData(700, FormatRange.CustomFormat, FormatNumber.Custom)]
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
            style2.Number = FormatNumber.Format15;
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
            copy.Number = FormatNumber.Format14;
            Assert.NotEqual(exampleStyle.GetHashCode(), copy.GetHashCode());
        }

        [Fact(DisplayName = "Test of the constant of the default custom format start number")]
        public void DefaultFontNameTest()
        {
            Assert.Equal(164, NumberFormat.CustomFormatStartNumber); // Expected 164
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
            numberFormat.Number = FormatNumber.Format11;
            Assert.NotEqual(s1, numberFormat.ToString()); // An explicit value comparison is probably not sensible
        }

    }
}
