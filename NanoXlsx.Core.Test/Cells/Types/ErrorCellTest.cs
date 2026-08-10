using NanoXLSX.Enums;
using NanoXLSX.Test.Core.Utils;
using Xunit;
using static NanoXLSX.Cell;

namespace NanoXLSX.Test.Core.Cells.Types
{
    // Ensure that these tests are executed sequentially, since static repository methods may be called
    [Collection(nameof(SequentialCollection))]
    public class ErrorCellTest
    {
        [Theory(DisplayName = "FormulaError values resolve to standalone error cells")]
        [InlineData(Errors.FormulaError.Null)]
        [InlineData(Errors.FormulaError.DivisionByZero)]
        [InlineData(Errors.FormulaError.Value)]
        [InlineData(Errors.FormulaError.Reference)]
        [InlineData(Errors.FormulaError.Name)]
        [InlineData(Errors.FormulaError.Number)]
        [InlineData(Errors.FormulaError.NotAvailable)]
        [InlineData(Errors.FormulaError.GettingData)]
        public void ErrorValueTest(Errors.FormulaError error)
        {
            Cell cell = new Cell(error, CellType.Default);

            Assert.Equal(CellType.Error, cell.DataType);
            Assert.Equal(error, cell.Value);
            Assert.Null(cell.Formula);
        }

        [Fact(DisplayName = "Changing an error value resolves the new ordinary cell type")]
        public void ErrorToStringTest()
        {
            Cell cell = new Cell(Errors.FormulaError.Reference, CellType.Error);

            cell.Value = "text";

            Assert.Equal(CellType.String, cell.DataType);
            Assert.Equal("text", cell.Value);
            Assert.Null(cell.Formula);
        }

        [Fact(DisplayName = "Forcing an ordinary value to Error preserves the value and creates no formula metadata")]
        public void ExplicitErrorTypeTest()
        {
            Cell cell = new Cell("#DIV/0!", CellType.String);

            cell.DataType = CellType.Error;

            Assert.Equal(CellType.Error, cell.DataType);
            Assert.Equal("#DIV/0!", cell.Value);
            Assert.Null(cell.Formula);
        }
    }
}
