using NanoXLSX;
using System;
using Xunit;
using static NanoXLSX.Cell;

namespace NanoXLSX_Test.Cells.Types
{
    // Ensure that these tests are executed sequentially, since static repository methods may be called 
    [Collection(nameof(SequentialCollection))]
    public class DateTimeCellTest
    {
        CellTypeUtils utils;

        public DateTimeCellTest()
        {
            utils = new CellTypeUtils();
        }

        [Fact(DisplayName = "DateTime value cell test: Test of the cell values, as well as proper modification")]
        public void DateCellTest()
        {
            // Date is hard to parametrize, therefore hardcoded
            DateTime defaultDateTime = new DateTime(2020, 11, 1, 11, 22, 13, 99);
            utils.AssertCellCreation<DateTime>(defaultDateTime, new DateTime(1900, 1, 1), CellType.DATE, (current, other) => { return current.Equals(other); });
            utils.AssertCellCreation<DateTime>(defaultDateTime, new DateTime(9999, 12, 31, 23, 59, 59), CellType.DATE, (current, other) => { return current.Equals(other); });
        }


        [Fact(DisplayName = "TimeSpan value cell test: Test of the cell values, as well as proper modification")]
        public void TimeSpanCellTest()
        {
            // TimeSpan is hard to parametrize, therefore hardcoded
            TimeSpan defaultTime = new TimeSpan(0, 22, 11, 7, 135);
            utils.AssertCellCreation<TimeSpan>(defaultTime, new TimeSpan(0, 0, 0), CellType.TIME, (current, other) => { return current.Equals(other); });
            utils.AssertCellCreation<TimeSpan>(defaultTime, new TimeSpan(2958465, 23, 59, 59, 999), CellType.TIME, (current, other) => { return current.Equals(other); });
        }

        [Fact(DisplayName = "Test of the DateTime comparison method on cells")]
        public void DateCellComparisonTest()
        {
            // Hard to parametrize, thus hardcoded
            DateTime baseDate = new DateTime(2020, 11, 5, 12, 23, 7, 157);
            DateTime nearBelowBase = new DateTime(2020, 11, 5, 12, 23, 7, 156);
            DateTime belowBase = new DateTime(2020, 11, 5, 8, 23, 7, 157);
            DateTime nearAboveBase = new DateTime(2020, 11, 5, 12, 23, 7, 158);
            DateTime aboveBase = new DateTime(2020, 12, 5, 12, 23, 7, 156);

            Cell baseCell = utils.CreateVariantCell(baseDate, utils.CellAddress);
            Cell equalCell = utils.CreateVariantCell(baseDate, utils.CellAddress);
            Cell nearBelowCell = utils.CreateVariantCell(nearBelowBase, utils.CellAddress);
            Cell nearAboveCell = utils.CreateVariantCell(nearAboveBase, utils.CellAddress);
            Cell belowCell = utils.CreateVariantCell(belowBase, utils.CellAddress);
            Cell aboveCell = utils.CreateVariantCell(aboveBase, utils.CellAddress);

            Assert.Equal(0, DateTime.Compare((DateTime)baseCell.Value, (DateTime)equalCell.Value));
            Assert.Equal(1, DateTime.Compare((DateTime)baseCell.Value, (DateTime)nearBelowCell.Value));
            Assert.Equal(-1, DateTime.Compare((DateTime)baseCell.Value, (DateTime)nearAboveCell.Value));
            Assert.Equal(1, DateTime.Compare((DateTime)baseCell.Value, (DateTime)belowCell.Value));
            Assert.Equal(-1, DateTime.Compare((DateTime)baseCell.Value, (DateTime)aboveCell.Value));
        }

        [Fact(DisplayName = "Test of the TimeSpan comparison method on cells")]
        public void TimeSpanCellComparisonTest()
        {
            // Hard to parametrize, thus hardcoded
            TimeSpan baseTime = new TimeSpan(1, 5, 7, 22, 113);
            TimeSpan nearBelowBase = new TimeSpan(1, 5, 7, 22, 112);
            TimeSpan belowBase = new TimeSpan(0, 5, 7, 22, 113);
            TimeSpan nearAboveBase = new TimeSpan(1, 5, 7, 22, 114);
            TimeSpan aboveBase = new TimeSpan(1, 5, 17, 22, 113);

            Cell baseCell = utils.CreateVariantCell(baseTime, utils.CellAddress);
            Cell equalCell = utils.CreateVariantCell(baseTime, utils.CellAddress);
            Cell nearBelowCell = utils.CreateVariantCell(nearBelowBase, utils.CellAddress);
            Cell nearAboveCell = utils.CreateVariantCell(nearAboveBase, utils.CellAddress);
            Cell belowCell = utils.CreateVariantCell(belowBase, utils.CellAddress);
            Cell aboveCell = utils.CreateVariantCell(aboveBase, utils.CellAddress);

            Assert.Equal(0, TimeSpan.Compare((TimeSpan)baseCell.Value, (TimeSpan)equalCell.Value));
            Assert.Equal(1, TimeSpan.Compare((TimeSpan)baseCell.Value, (TimeSpan)nearBelowCell.Value));
            Assert.Equal(-1, TimeSpan.Compare((TimeSpan)baseCell.Value, (TimeSpan)nearAboveCell.Value));
            Assert.Equal(1, TimeSpan.Compare((TimeSpan)baseCell.Value, (TimeSpan)belowCell.Value));
            Assert.Equal(-1, TimeSpan.Compare((TimeSpan)baseCell.Value, (TimeSpan)aboveCell.Value));
        }

    }
}
