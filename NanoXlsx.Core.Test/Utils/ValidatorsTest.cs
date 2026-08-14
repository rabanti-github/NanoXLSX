using NanoXLSX.Exceptions;
using NanoXLSX.Utils;
using Xunit;

namespace NanoXLSX.Test.Core.UtilsTest
{
    public class ValidatorsTest
    {
        [Theory(DisplayName = "Test of the successful Validator function ValidateColor")]
        [InlineData("000000", false, true)]
        [InlineData("000000", false, false)]
        [InlineData("00AACC", false, true)]
        [InlineData("00AACC", false, false)]
        [InlineData("FFFFFF", false, true)]
        [InlineData("FFFFFF", false, false)]
        [InlineData("00000000", true, true)]
        [InlineData("00000000", true, false)]
        [InlineData("FF000000", true, true)]
        [InlineData("FF000000", true, false)]
        [InlineData("00AACC00", true, true)]
        [InlineData("00AACC00", true, false)]
        [InlineData("", true, true)]
        [InlineData("", false, true)]
        [InlineData(null, true, true)]
        [InlineData(null, false, true)]
        public void ValidateColorTest(string givenHexCode, bool givenUseAlpha, bool givenAllowEmpty)
        {
            Validators.ValidateColor(givenHexCode, givenUseAlpha, givenAllowEmpty);
            Assert.True(true);
        }

        [Theory(DisplayName = "Test of the failing Validator function ValidateColor")]
        [InlineData("000000", true, true)]
        [InlineData("FFFFFF", true, true)]
        [InlineData("0ACFD", true, true)]
        [InlineData("00000000", false, true)]
        [InlineData("00FFFFFF", false, true)]
        [InlineData("000ACFD", false, true)]
        [InlineData("FF000000", false, true)]
        [InlineData("FFFFFFFF", false, true)]
        [InlineData("FF0ACFD", false, true)]
        [InlineData("AA", false, true)]
        [InlineData("CCC", false, true)]
        [InlineData("DDDD", false, true)]
        [InlineData("001122", true, true)]
        [InlineData("X", false, true)]
        [InlineData("AAX022", false, true)]
        [InlineData(" ", false, true)]
        [InlineData("0 0000", false, true)]
        [InlineData("", false, false)]
        [InlineData(null, false, false)]

        public void ValidateColorFailTest(string givenHexCode, bool givenUseAlpha, bool givenAllowEmpty)
        {
            Assert.Throws<Exceptions.StyleException>(() => Validators.ValidateColor(givenHexCode, givenUseAlpha, givenAllowEmpty));
        }

        [Theory(DisplayName = "Test of the successful Validator function ValidateGenericColor")]
        [InlineData("000000", true)]
        [InlineData("000000", false)]
        [InlineData("00AACC", true)]
        [InlineData("00AACC", false)]
        [InlineData("FFFFFF", true)]
        [InlineData("FFFFFF", false)]
        [InlineData("00000000", true)]
        [InlineData("00000000", false)]
        [InlineData("FF000000", true)]
        [InlineData("FF000000", false)]
        [InlineData("00AACC00", true)]
        [InlineData("00AACC00", false)]
        [InlineData("", true)]
        [InlineData(null, true)]
        public void ValidateGenericColorTest(string givenHexCode, bool givenAllowEmpty)
        {
            Validators.ValidateGenericColor(givenHexCode, givenAllowEmpty);
            Assert.True(true);
        }

        [Theory(DisplayName = "Test of the failing Validator function ValidateGenericColor")]
        [InlineData("0ACFD", true)]
        [InlineData("000ACFD", true)]
        [InlineData("FF0ACFD", true)]
        [InlineData("AA", true)]
        [InlineData("CCC", true)]
        [InlineData("DDDD", true)]
        [InlineData("X", true)]
        [InlineData("AAX022", true)]
        [InlineData(" ", true)]
        [InlineData("0 0000", true)]
        [InlineData("", false)]
        [InlineData(null, false)]

        public void ValidateGenericColorFailTest(string givenHexCode, bool givenAllowEmpty)
        {
            Assert.Throws<Exceptions.StyleException>(() => Validators.ValidateGenericColor(givenHexCode, givenAllowEmpty));
        }

        [Theory(DisplayName = "Test of the successful Validator function ValidateCellAddressExpression")]
        [InlineData("A1", Cell.AddressScope.Any)]
        [InlineData("$XFD$1048576", Cell.AddressScope.Any)]
        [InlineData("A1:B2", Cell.AddressScope.Any)]
        [InlineData("A1", Cell.AddressScope.SingleAddress)]
        [InlineData("$A$1", Cell.AddressScope.SingleAddress)]
        [InlineData("A1:B2", Cell.AddressScope.Range)]
        [InlineData("A1:A1", Cell.AddressScope.Range)]
        [InlineData(null, Cell.AddressScope.Invalid)]
        [InlineData("", Cell.AddressScope.Invalid)]
        [InlineData(" ", Cell.AddressScope.Invalid)]
        [InlineData("A1:B2:C3", Cell.AddressScope.Invalid)]
        [InlineData("$A$$1", Cell.AddressScope.Invalid)]
        [InlineData("世界1", Cell.AddressScope.Invalid)]
        [InlineData("A0", Cell.AddressScope.Invalid)]
        public void ValidateCellAddressExpressionTest(string givenExpression, Cell.AddressScope givenScope)
        {
            Validators.ValidateCellAddressExpression(givenExpression, givenScope);
            Assert.True(true);
        }

        [Theory(DisplayName = "Test of the failing Validator function ValidateCellAddressExpression")]
        [InlineData(null, Cell.AddressScope.Any)]
        [InlineData("", Cell.AddressScope.Any)]
        [InlineData(" ", Cell.AddressScope.Any)]
        [InlineData("世界1", Cell.AddressScope.Any)]
        [InlineData("A0", Cell.AddressScope.Any)]
        [InlineData("XFE1", Cell.AddressScope.Any)]
        [InlineData("A1048577", Cell.AddressScope.Any)]
        [InlineData("A1:B2:C3", Cell.AddressScope.Any)]
        [InlineData(null, Cell.AddressScope.SingleAddress)]
        [InlineData("A1:B2", Cell.AddressScope.SingleAddress)]
        [InlineData("A1:A1", Cell.AddressScope.SingleAddress)]
        [InlineData(null, Cell.AddressScope.Range)]
        [InlineData("A1", Cell.AddressScope.Range)]
        [InlineData("A1:B2:C3", Cell.AddressScope.Range)]
        [InlineData("$A$$1", Cell.AddressScope.SingleAddress)]
        public void ValidateCellAddressExpressionFailTest(string givenExpression, Cell.AddressScope givenScope)
        {
            FormatException exception = Assert.Throws<FormatException>(
                () => Validators.ValidateCellAddressExpression(givenExpression, givenScope));
            Assert.NotNull(exception.InnerException);
        }

        [Theory(DisplayName = "Test of the inverted Validator function ValidateCellAddressExpression")]
        [InlineData("A1")]
        [InlineData("$XFD$1048576")]
        [InlineData("A1:B2")]
        public void ValidateCellAddressExpressionInvalidScopeFailTest(string givenExpression)
        {
            FormatException exception = Assert.Throws<FormatException>(
                () => Validators.ValidateCellAddressExpression(givenExpression, Cell.AddressScope.Invalid));
            Assert.Null(exception.InnerException);
            // TODO fix this term if the naming changes in ValidateCellAddressExpression
            Assert.Equal("The passed expression is valid cell address or range, but the validation was explicitly inverted", exception.Message);
        }

        [Theory(DisplayName = "Test of the ValidateWorksheetName function")]
        [InlineData("1", true)]
        [InlineData("test", true)]
        [InlineData("test-test", true)]
        [InlineData("$$$", true)]
        [InlineData("a b", true)]
        [InlineData("a\tb", true)]
        [InlineData("-------------------------------", true)]
        [InlineData("", false)]
        [InlineData(null, false)]
        [InlineData("a[b", false)]
        [InlineData("a]b", false)]
        [InlineData("a*b", false)]
        [InlineData("a?b", false)]
        [InlineData("a/b", false)]
        [InlineData("a\\b", false)]
        [InlineData("--------------------------------", false)]
        public void SetSheetNameTest(string name, bool expectedValid)
        {
            Worksheet worksheet = new Worksheet();
            Assert.Null(worksheet.SheetName);
            if (expectedValid)
            {
                Validators.ValidateWorksheetName(name);
                Assert.True(true);
            }
            else
            {
                Assert.Throws<FormatException>(() => Validators.ValidateWorksheetName(name));
            }
        }
    }
}
