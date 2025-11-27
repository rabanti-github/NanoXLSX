using NanoXLSX.Styles;
using NanoXLSX.Test.Core.Utils;
using Xunit;
using static NanoXLSX.Cell;

namespace NanoXLSX.Test.Core.Cells.Types
{
    // Ensure that these tests are executed sequentially, since static repository methods may be called 
    [Collection(nameof(SequentialCollection))]
    public class BooleanCellTest
    {
        CellTypeUtils utils;

        public BooleanCellTest()
        {
            utils = new CellTypeUtils();
        }


        [Fact(DisplayName = "Bool value cell test: Test of the cell values, as well as proper modification")]
        public void BoolCellTest()
        {
            utils.AssertCellCreation<bool>(true, true, CellType.Bool, (curent, other) => { return curent.Equals(other); });
            utils.AssertCellCreation<bool>(true, false, CellType.Bool, (curent, other) => { return curent.Equals(other); });
        }


        [Fact(DisplayName = "Bool value cell test with style")]
        public void BoolCellTest2()
        {
            Style style = BasicStyles.Bold;
            utils.AssertStyledCellCreation<bool>(true, true, CellType.Bool, (curent, other) => { return curent.Equals(other); }, style);
            utils.AssertStyledCellCreation<bool>(true, false, CellType.Bool, (curent, other) => { return curent.Equals(other); }, style);
        }


        [Theory(DisplayName = "Test of the bool comparison method on cells")]
        [InlineData(true, true, 0)]
        [InlineData(false, false, 0)]
        [InlineData(true, false, 1)]
        [InlineData(false, true, -1)]
        public void BoolCellComparisonTest(bool value1, bool value2, int expectedResult)
        {
            Cell cell1 = utils.CreateVariantCell<bool>(value1, utils.CellAddress);
            Cell cell2 = utils.CreateVariantCell<bool>(value2, utils.CellAddress);
            int comparison = ((bool)cell1.Value).CompareTo(cell2.Value);
            Assert.Equal(comparison, expectedResult);
        }




    }
}
