using System;
using NanoXLSX.Enums;
using Xunit;

namespace NanoXLSX.Test.Core.MiscTest
{
    public class EnumTest
    {
        [Theory(DisplayName = "Test parsing formula errors")]
        [InlineData("#NULL!", Errors.FormulaError.Null)]
        [InlineData("#DIV/0!", Errors.FormulaError.DivisionByZero)]
        [InlineData("#VALUE!", Errors.FormulaError.Value)]
        [InlineData("#REF!", Errors.FormulaError.Reference)]
        [InlineData("#NAME?", Errors.FormulaError.Name)]
        [InlineData("#NUM!", Errors.FormulaError.Number)]
        [InlineData("#N/A", Errors.FormulaError.NotAvailable)]
        [InlineData("#GETTING_DATA", Errors.FormulaError.GettingData)]
        public void TryParseFormulaErrorTest(string value, Errors.FormulaError expected)
        {
            bool result = Errors.TryParseFormulaError(value, out Errors.FormulaError error);

            Assert.True(result);
            Assert.Equal(expected, error);
        }

        [Theory(DisplayName = "Test parsing invalid formula errors")]
        [InlineData(null)]
        [InlineData("")]
        [InlineData(" ")]
        [InlineData("#NAME? ")]
        [InlineData("#name?")]
        [InlineData("#NAME")]
        [InlineData("NAME?")]
        [InlineData("#UNKNOWN!")]
        public void TryParseInvalidFormulaErrorTest(string value)
        {
            bool result = Errors.TryParseFormulaError(value, out Errors.FormulaError error);

            Assert.False(result);
            Assert.Equal(Errors.FormulaError.UnknownError, error);
        }

        [Theory(DisplayName = "Test conversion of formula errors to strings")]
        [InlineData(Errors.FormulaError.Null, "#NULL!")]
        [InlineData(Errors.FormulaError.DivisionByZero, "#DIV/0!")]
        [InlineData(Errors.FormulaError.Value, "#VALUE!")]
        [InlineData(Errors.FormulaError.Reference, "#REF!")]
        [InlineData(Errors.FormulaError.Name, "#NAME?")]
        [InlineData(Errors.FormulaError.Number, "#NUM!")]
        [InlineData(Errors.FormulaError.NotAvailable, "#N/A")]
        [InlineData(Errors.FormulaError.GettingData, "#GETTING_DATA")]
        public void FormulaErrorToStringTest(Errors.FormulaError error, string expected)
        {
            string result = Errors.FormulaErrorToString(error);

            Assert.Equal(expected, result);
        }

        [Theory(DisplayName = "Test conversion of invalid formula errors to strings")]
        [InlineData(Errors.FormulaError.NoError)]
        [InlineData(Errors.FormulaError.UnknownError)]
        [InlineData((Errors.FormulaError)999)]
        public void FormulaErrorToStringFailTest(Errors.FormulaError error)
        {
            Assert.Throws<FormatException>(() => Errors.FormulaErrorToString(error));
        }
    }
}
