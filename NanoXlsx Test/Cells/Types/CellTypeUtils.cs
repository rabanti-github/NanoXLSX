using NanoXLSX;
using NanoXLSX.Styles;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;
using static NanoXLSX.Cell;

namespace NanoXLSX_Test.Cells.Types
{
    // Ensure that these tests are executed sequentially, since static repository methods are called 
    [Collection(nameof(SequentialCollection))]
    public class CellTypeUtils
    {

        Address cellAddress;
        Workbook workbook;
        Worksheet worksheet;

        public CellTypeUtils()
        {
            cellAddress = new Address(0, 0);
            workbook = new Workbook(true);
            worksheet = workbook.CurrentWorksheet;
        }

        public Address CellAddress
        {
            get
            {
                return cellAddress;
            }
        }

        public void AssertCellCreation<T>(T initialValue, T expectedValue, CellType expectedType, Func<T, T, bool> comparer)
        {
            AssertCellCreation<T>(initialValue, expectedValue, expectedType, comparer, null);
        }

        public void AssertStyledCellCreation<T>(T initialValue, T expectedValue, CellType expectedType, Func<T, T, bool> comparer, Style style)
        {
            AssertCellCreation<T>(initialValue, expectedValue, expectedType, comparer, style);
        }

        private void AssertCellCreation<T>(T initialValue, T expectedValue, CellType expectedType, Func<T, T, bool> comparer, Style style)
        {
            Cell actualCell = new Cell(initialValue, Cell.CellType.DEFAULT, this.cellAddress);
            if (style != null)
            {
                actualCell.WorksheetReference = worksheet;
                actualCell.SetStyle(style);
            }
            Assert.True(comparer.Invoke(initialValue, (T)actualCell.Value));
            Assert.Equal(typeof(T), actualCell.Value.GetType());
            Assert.Equal(expectedType, actualCell.DataType);
            actualCell.Value = expectedValue;
            Assert.True(comparer.Invoke(expectedValue, (T)actualCell.Value));
            if (style != null)
            {
                // Note: Date and Time styles are set internally and are not asserted if style is null.
                // The same applies to merged styles. These must be asserted separately
                Assert.Equal(style, actualCell.CellStyle);
            }

        }

        public Cell CreateVariantCell<T>(T value, Address cellAddress, bool referenceWorksheet = false, Style style = null)
        {
            Cell givenCell = new Cell(value, CellType.DEFAULT, cellAddress);
            if (style != null || referenceWorksheet)
            {
                givenCell.WorksheetReference = worksheet;
            }
            if (style != null)
            {
                givenCell.SetStyle(style);
            }
            return givenCell;
        }

    }
}
