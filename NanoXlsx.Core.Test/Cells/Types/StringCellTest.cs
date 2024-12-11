using NanoXLSX;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;
using static NanoXLSX.Cell;

namespace NanoXLSX.Test.Cells.Types
{
    // Ensure that these tests are executed sequentially, since static repository methods may be called 
    [Collection(nameof(SequentialCollection))]
    public class StringCellTest
    {
        CellTypeUtils utils;

        public StringCellTest()
        {
            utils = new CellTypeUtils();
        }


        [Theory(DisplayName = "String value cell test: Test of the cell values, as well as proper modification")]
        [InlineData("")]
        [InlineData(null)] // Empty cell
        [InlineData("Text")]
        [InlineData(" ")]
        [InlineData("start\tend")]
        public void StringsCellTest(string value)
        {
            utils.AssertCellCreation<string>("Initial Value", value, CellType.STRING, CompareString);
        }

        [Theory(DisplayName = "Test of the string comparison method on cells")]
        [InlineData(null, null, 0)]
        [InlineData(null, "X", -1)]
        [InlineData("x", null, 1)]
        [InlineData("", "", 0)]
        [InlineData(" ", " ", 0)]
        [InlineData("a", "b", -1)]
        [InlineData("9", "8", 1)]
        public void StringCellComparisonTest(string value1, string value2, int expectedResult)
        {
            Cell cell1 = utils.CreateVariantCell<string>(value1, utils.CellAddress);
            Cell cell2 = utils.CreateVariantCell<string>(value2, utils.CellAddress);
            int comparison = String.Compare(cell1.Value as string, cell2.Value as string);
            Assert.Equal(comparison, expectedResult);
        }

        private static bool CompareString(string current, string other)
        {
            if (current == null && other == null)
            {
                return true;
            }
            else if (current == null)
            {
                return false;
            }
            return current.Equals(other);
        }

    }
}
