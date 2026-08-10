using NanoXLSX.Test.Core.Utils;
using Xunit;
using static NanoXLSX.Cell;

namespace NanoXLSX.Test.Core.Cells.Types
{
    // Ensure that these tests are executed sequentially, since static repository methods may be called
    [Collection(nameof(SequentialCollection))]
    public class FormulaCellTest
    {
        [Theory(DisplayName = "Formula value cell test: Test creation and value modification")]
        [InlineData("A1", "SUM(A1:A3)")]
        [InlineData("Text", "[1]ExternalSheet!A1")]
        [InlineData("", "B2")]
        public void FormulaValueTest(string initialValue, string expectedValue)
        {
            Cell cell = new Cell(initialValue, CellType.Formula, new Address(0, 0));

            Assert.Equal(CellType.Formula, cell.DataType);
            Assert.Equal(initialValue, cell.Value);
            Assert.NotNull(cell.Formula);
            Assert.Equal(initialValue, cell.Formula.Expression);

            FormulaData formula = cell.Formula;
            cell.Value = expectedValue;

            Assert.Equal(CellType.Formula, cell.DataType);
            Assert.Same(formula, cell.Formula);
            Assert.Equal(expectedValue, cell.Value);
            Assert.Equal(expectedValue, cell.Formula.Expression);
        }

        [Fact(DisplayName = "Changing a string cell to Formula creates synchronized formula metadata")]
        public void StringToFormulaTest()
        {
            Cell cell = new Cell("Initial Value", CellType.String);

            cell.DataType = CellType.Formula;

            Assert.Equal(CellType.Formula, cell.DataType);
            Assert.NotNull(cell.Formula);
            Assert.Equal("Initial Value", cell.Formula.Expression);

            cell.Value = "[1]ExternalSheet!A1";
            Assert.Equal("[1]ExternalSheet!A1", cell.Formula.Expression);
            Assert.True(cell.Formula.HasExternalReferences);
        }

        [Theory(DisplayName = "Changing a formula cell to a non-formula type clears formula metadata")]
        [InlineData(CellType.String)]
        [InlineData(CellType.Number)]
        [InlineData(CellType.Bool)]
        [InlineData(CellType.Empty)]
        [InlineData(CellType.Date)]
        [InlineData(CellType.Time)]
        [InlineData(CellType.Error)]
        public void FormulaToNonFormulaTest(CellType targetType)
        {
            Cell cell = new Cell("A1", CellType.Formula);

            cell.DataType = targetType;

            Assert.Equal(targetType, cell.DataType);
            Assert.Equal("A1", cell.Value);
            Assert.Null(cell.Formula);
        }

        [Fact(DisplayName = "Assigning null to a formula cell resolves Empty and clears formula metadata")]
        public void FormulaToEmptyByValueTest()
        {
            Cell cell = new Cell("A1", CellType.Formula);

            cell.Value = null;

            Assert.Equal(CellType.Empty, cell.DataType);
            Assert.Null(cell.Value);
            Assert.Null(cell.Formula);
        }

        [Fact(DisplayName = "A cached formula error does not change the formula cell type or expression")]
        public void FormulaCachedErrorTest()
        {
            Cell cell = new Cell("A1/0", CellType.Formula);
            FormulaData formula = cell.Formula;

            formula.CachedValue = Enums.Errors.FormulaError.DivisionByZero;
            formula.CachedValueType = CellType.Error;

            Assert.Equal(CellType.Formula, cell.DataType);
            Assert.Equal("A1/0", cell.Value);
            Assert.Same(formula, cell.Formula);
            Assert.Equal("A1/0", cell.Formula.Expression);
            Assert.Equal(Enums.Errors.FormulaError.DivisionByZero, cell.Formula.CachedValue);
            Assert.Equal(CellType.Error, cell.Formula.CachedValueType);
        }

        [Fact(DisplayName = "Linked formula cells retain cached-value behavior")]
        public void LinkedFormulaValueTest()
        {
            Cell cell = new Cell(null, CellType.Formula);
            cell.Formula.MasterCellAddress = "A1";

            cell.Value = "cached value";

            Assert.Equal(CellType.Formula, cell.DataType);
            Assert.Equal("cached value", cell.Value);
            Assert.Null(cell.Formula.Expression);
        }

        [Fact(DisplayName = "Copying a formula cell creates independent formula metadata")]
        public void FormulaCopyTest()
        {
            Cell cell = new Cell("[1]ExternalSheet!A1", CellType.Formula, new Address(2, 3));

            Cell copy = cell.Copy();

            Assert.Equal(cell.DataType, copy.DataType);
            Assert.Equal(cell.Value, copy.Value);
            Assert.NotNull(copy.Formula);
            Assert.NotSame(cell.Formula, copy.Formula);
            Assert.Equal(cell.Formula.Expression, copy.Formula.Expression);
            Assert.Equal(cell.Formula.HasExternalReferences, copy.Formula.HasExternalReferences);
        }
    }
}
