using System.Collections.Generic;
using NanoXLSX.Exceptions;
using NanoXLSX.Styles;
using Xunit;
using FormatException = NanoXLSX.Exceptions.FormatException;

namespace NanoXLSX.Test.Core.CellTest
{
    public class CellReferenceTest
    {
        [Theory(DisplayName = "Test of all AddCellReference overloads")]
        [InlineData(0)] // row, column
        [InlineData(1)] // row, column, style
        [InlineData(2)] // address
        [InlineData(3)] // address, style
        public void AddCellReferenceOverloadTest(int overload)
        {
            Workbook workbook = new Workbook("Sheet1");
            DefinedName name = workbook.AddDefinedNameFormula("FormulaName", "SUM(A1:A2)");
            Style style = (Style)BasicStyles.Bold.Copy();
            IReadOnlyList<Address> addresses;
            switch (overload)
            {
                case 0:
                    addresses = workbook.CurrentWorksheet.AddCellReference(name, 1, 1, 7);
                    break;
                case 1:
                    addresses = workbook.CurrentWorksheet.AddCellReference(name, 1, 1, style, 7);
                    break;
                case 2:
                    addresses = workbook.CurrentWorksheet.AddCellReference(name, "B2", 7);
                    break;
                default:
                    addresses = workbook.CurrentWorksheet.AddCellReference(name, "B2", style, 7);
                    break;
            }

            Assert.Single(addresses);
            Assert.Equal(new Address("B2"), addresses[0]);
            Cell cell = workbook.CurrentWorksheet.Cells["B2"];
            Assert.Equal(Cell.CellType.Formula, cell.DataType);
            Assert.Equal("FormulaName", cell.Value);
            Assert.Equal("FormulaName", cell.Formula.Expression);
            Assert.Same(name, cell.Formula.DefinedNameReference);
            Assert.Equal("7", cell.Formula.CachedValue);
            Assert.Equal(Cell.CellType.Number, cell.Formula.CachedValueType);
            Assert.Null(cell.Formula.FormulaRange);
            if (overload == 1 || overload == 3)
            {
                Assert.Equal(style.GetHashCode(), cell.CellStyle.GetHashCode());
            }
            else
            {
                Assert.Null(cell.CellStyle);
            }
        }

        [Fact(DisplayName = "Test of cell and constant defined name references")]
        public void CellAndConstantReferenceTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            DefinedName cellName = workbook.AddDefinedNameCell("CellName", workbook.CurrentWorksheet, "A1");
            DefinedName constantName = workbook.AddDefinedNameConstant("ConstantName", true);

            workbook.CurrentWorksheet.AddCellReference(cellName, "B1");
            workbook.CurrentWorksheet.AddCellReference(constantName, "B2", 999);

            Assert.Null(workbook.CurrentWorksheet.Cells["B1"].Formula.FormulaRange);
            Assert.Equal(Cell.CellType.Number, workbook.CurrentWorksheet.Cells["B1"].Formula.CachedValueType);
            Assert.Equal("TRUE", workbook.CurrentWorksheet.Cells["B2"].Formula.CachedValue);
            Assert.Equal(Cell.CellType.Bool, workbook.CurrentWorksheet.Cells["B2"].Formula.CachedValueType);
        }

        [Theory(DisplayName = "Test of range defined name array dimensions")]
        [InlineData("A1:A3", "D4", "D4,D5,D6")]
        [InlineData("A1:C1", "D4", "D4,E4,F4")]
        [InlineData("A1:B3", "D4", "D4,D5,D6,E4,E5,E6")]
        [InlineData("A1:A1", "D4", "D4")]
        public void RangeReferenceTest(string sourceRange, string target, string expectedAddresses)
        {
            Workbook workbook = new Workbook("Sheet1");
            DefinedName name = workbook.AddDefinedNameRange("RangeName", workbook.CurrentWorksheet, sourceRange);
            workbook.CurrentWorksheet.AddCell("overwritten", "E5");

            IReadOnlyList<Address> addresses = workbook.CurrentWorksheet.AddCellReference(name, target);
            string[] expected = expectedAddresses.Split(',');
            Assert.Equal(expected.Length, addresses.Count);
            for (int i = 0; i < expected.Length; i++)
            {
                Assert.Equal(new Address(expected[i]), addresses[i]);
            }

            Cell master = workbook.CurrentWorksheet.Cells[target];
            Assert.Equal(FormulaData.FormulaType.Array, master.Formula.Type);
            Assert.Equal(new Range(new Address(expected[0]), new Address(expected[expected.Length - 1])).ToString(), master.Formula.FormulaRange);
            Assert.Null(master.Formula.MasterCellAddress);
            for (int i = 1; i < expected.Length; i++)
            {
                Cell dependent = workbook.CurrentWorksheet.Cells[expected[i]];
                Assert.Equal(Cell.CellType.Formula, dependent.DataType);
                Assert.Equal(FormulaData.FormulaType.Array, dependent.Formula.Type);
                Assert.Equal(target, dependent.Formula.MasterCellAddress);
                Assert.Equal(master.Formula.FormulaRange, dependent.Formula.FormulaRange);
                Assert.Null(dependent.Formula.DefinedNameReference);
            }
        }

        [Fact(DisplayName = "Test of a styled range defined name reference")]
        public void StyledRangeReferenceTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            DefinedName name = workbook.AddDefinedNameRange("RangeName", workbook.CurrentWorksheet, "A1:B2");
            Style style = (Style)BasicStyles.Italic.Copy();
            IReadOnlyList<Address> addresses = workbook.CurrentWorksheet.AddCellReference(name, "C3", style, "cached");

            Assert.Equal(4, addresses.Count);
            foreach (Address address in addresses)
            {
                Assert.Equal(style.GetHashCode(), workbook.CurrentWorksheet.Cells[address.ToString()].CellStyle.GetHashCode());
            }
            Assert.Equal("cached", workbook.CurrentWorksheet.Cells["C3"].Formula.CachedValue);
            Assert.Equal(Cell.CellType.String, workbook.CurrentWorksheet.Cells["C3"].Formula.CachedValueType);
        }

        [Fact(DisplayName = "Test of invalid AddCellReference inputs")]
        public void InvalidReferenceTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            DefinedName name = workbook.AddDefinedNameConstant("Name", 1);
            Assert.Throws<WorksheetException>(() => workbook.CurrentWorksheet.AddCellReference(null, "A1"));
            Assert.Throws<WorksheetException>(() => workbook.CurrentWorksheet.AddCellReference(null, 0, 0));
            Assert.Throws<WorksheetException>(() => workbook.CurrentWorksheet.AddCellReference(null, 0, 0, (Style)BasicStyles.Bold.Copy()));
            Assert.Throws<WorksheetException>(() => new Cell().SetReference(null));
            Assert.Throws<FormatException>(() => workbook.CurrentWorksheet.AddCellReference(name, "invalid"));
            Assert.Throws<RangeException>(() => workbook.CurrentWorksheet.AddCellReference(name, -1, 0));
        }
    }
}
