using NanoXLSX;
using NanoXLSX.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace NanoXLSX_Test.Worksheets
{
    public class PaneTest
    {

        [Fact(DisplayName = "Test of the get function of the PaneSplitTopHeight property")]
        public void PaneSplitTopHeightTest()
        {
            Worksheet worksheet = new Worksheet();
            Assert.Null(worksheet.PaneSplitTopHeight);
            worksheet.SetSplit(10f, 22.2f, new Address("A2"), Worksheet.WorksheetPane.bottomLeft);
            Assert.NotNull(worksheet.PaneSplitTopHeight);
            Assert.Equal(22.2f, worksheet.PaneSplitTopHeight);
            worksheet.ResetSplit();
            Assert.Null(worksheet.PaneSplitTopHeight);
        }

        [Fact(DisplayName = "Test of the get function of the PaneSplitLeftWidth property")]
        public void PaneSplitLeftWidthTest()
        {
            Worksheet worksheet = new Worksheet();
            Assert.Null(worksheet.PaneSplitLeftWidth);
            worksheet.SetSplit(11.1f, 20f, new Address("A2"), Worksheet.WorksheetPane.bottomLeft);
            Assert.NotNull(worksheet.PaneSplitLeftWidth);
            Assert.Equal(11.1f, worksheet.PaneSplitLeftWidth);
            worksheet.ResetSplit();
            Assert.Null(worksheet.PaneSplitLeftWidth);
        }

        [Fact(DisplayName = "Test of the get function of the FreezeSplitPanes property")]
        public void FreezeSplitPanesTest()
        {
            Worksheet worksheet = new Worksheet();
            Assert.Null(worksheet.FreezeSplitPanes);
            worksheet.SetSplit(2, 2, true, new Address("D4"), Worksheet.WorksheetPane.bottomRight);
            Assert.NotNull(worksheet.FreezeSplitPanes);
            Assert.Equal(true, worksheet.FreezeSplitPanes);
            worksheet.ResetSplit();
            Assert.Null(worksheet.FreezeSplitPanes);
        }

        [Fact(DisplayName = "Test of the get function of the PaneSplitTopLeftCell property")]
        public void PaneSplitTopLeftCellTest()
        {
            Worksheet worksheet = new Worksheet();
            Assert.Null(worksheet.PaneSplitTopLeftCell);
            worksheet.SetSplit(10f, 22.2f, new Address("C4"), Worksheet.WorksheetPane.bottomLeft);
            Assert.NotNull(worksheet.PaneSplitTopLeftCell);
            Assert.Equal("C4", worksheet.PaneSplitTopLeftCell.Value.GetAddress());
            worksheet.ResetSplit();
            Assert.Null(worksheet.PaneSplitTopLeftCell);
        }

        [Fact(DisplayName = "Test of the get function of the PaneSplitAddress property")]
        public void PaneSplitAddressTest()
        {
            Worksheet worksheet = new Worksheet();
            Assert.Null(worksheet.PaneSplitAddress);
            worksheet.SetSplit(2, 2, true, new Address("D4"), Worksheet.WorksheetPane.bottomRight);
            Assert.NotNull(worksheet.PaneSplitAddress);
            Assert.Equal("C3", worksheet.PaneSplitAddress.Value.GetAddress());
            worksheet.ResetSplit();
            Assert.Null(worksheet.PaneSplitAddress);
        }

        [Fact(DisplayName = "Test of the get function of the ActivePane property")]
        public void ActivePaneTest()
        {
            Worksheet worksheet = new Worksheet();
            Assert.Null(worksheet.ActivePane);
            worksheet.SetSplit(2, 2, true, new Address("D4"), Worksheet.WorksheetPane.bottomRight);
            Assert.NotNull(worksheet.ActivePane);
            Assert.Equal(Worksheet.WorksheetPane.bottomRight, worksheet.ActivePane);
            worksheet.ResetSplit();
            Assert.Null(worksheet.ActivePane);
        }

        [Theory(DisplayName = "Test of the SetHorizontalSplit function with height definition")]
        [InlineData(22.2f, "B2", Worksheet.WorksheetPane.bottomLeft)]
        [InlineData(0f, "B2", Worksheet.WorksheetPane.bottomLeft)]
        [InlineData(500f, "B2", Worksheet.WorksheetPane.bottomLeft)]
        [InlineData(22.2f, "X1", Worksheet.WorksheetPane.bottomLeft)]
        [InlineData(0f, "A1", Worksheet.WorksheetPane.bottomLeft)]
        [InlineData(500f, "XFD1048576", Worksheet.WorksheetPane.bottomLeft)]
        [InlineData(22.2f, "B2", Worksheet.WorksheetPane.topRight)]
        [InlineData(0f, "B2", Worksheet.WorksheetPane.bottomRight)]
        [InlineData(500f, "B2", Worksheet.WorksheetPane.topLeft)]
        public void SetHorizontalSplitTest(float height, string topLeftCellAddress, Worksheet.WorksheetPane activePane)
        {
            Worksheet worksheet = new Worksheet();
            AssertInitializedPaneSplit(worksheet);
            Address address = new Address(topLeftCellAddress);
            worksheet.SetHorizontalSplit(height, address, activePane);
            Assert.Equal(height, worksheet.PaneSplitTopHeight);
            Assert.Equal(address, worksheet.PaneSplitTopLeftCell);
            Assert.Equal(activePane, worksheet.ActivePane);
            Assert.Null(worksheet.FreezeSplitPanes);
            Assert.Null(worksheet.PaneSplitAddress);
            Assert.Null(worksheet.PaneSplitLeftWidth);
        }

        [Theory(DisplayName = "Test of the SetHorizontalSplit function with row definition")]
        [InlineData(3, false, "A1", Worksheet.WorksheetPane.bottomLeft)]
        [InlineData(10, true, "K11", Worksheet.WorksheetPane.bottomLeft)]
        [InlineData(3, false, "E2", Worksheet.WorksheetPane.bottomRight)]
        [InlineData(10, true, "L100", Worksheet.WorksheetPane.bottomRight)]
        [InlineData(3, false, "F3", Worksheet.WorksheetPane.topLeft)]
        [InlineData(10, true, "M200", Worksheet.WorksheetPane.topLeft)]
        [InlineData(3, false, "F3", Worksheet.WorksheetPane.topRight)]
        [InlineData(10, true, "M11", Worksheet.WorksheetPane.topRight)]
        public void SetHorizontalSplitTest2(int rowNumber, bool freeze, string topLeftCellAddress, Worksheet.WorksheetPane activePane)
        {
            Worksheet worksheet = new Worksheet();
            AssertInitializedPaneSplit(worksheet);
            Address address = new Address(topLeftCellAddress);
            worksheet.SetHorizontalSplit(rowNumber, freeze, address, activePane);
            Assert.Null(worksheet.PaneSplitLeftWidth);
            Assert.Null(worksheet.PaneSplitTopHeight);
            Address expectedAddress = new Address(0, rowNumber);
            Assert.Equal(expectedAddress.GetAddress(), worksheet.PaneSplitAddress.Value.GetAddress());
            Assert.Equal(freeze, worksheet.FreezeSplitPanes);
            Assert.Equal(address, worksheet.PaneSplitTopLeftCell);
            Assert.Equal(activePane, worksheet.ActivePane);
        }

        [Theory(DisplayName = "Test of the failing SetHorizontalSplit function")]
        [InlineData(3, false, "A1", true)]
        [InlineData(3, true, "A1", false)]
        [InlineData(100, false, "R100", true)]
        [InlineData(100, true, "R100", false)]
        public void SetHorizontalSplitFailTest(int rowNumber, bool freeze, string topLeftCellAddress, bool expectedValid)
        {
            Worksheet worksheet = new Worksheet();
            AssertInitializedPaneSplit(worksheet);
            Address address = new Address(topLeftCellAddress);
            if (expectedValid)
            {
                worksheet.SetHorizontalSplit(rowNumber, freeze, address, Worksheet.WorksheetPane.bottomLeft);
            }
            else
            {
                Assert.Throws<WorksheetException>(() => worksheet.SetHorizontalSplit(rowNumber, freeze, address, Worksheet.WorksheetPane.bottomLeft));
            }
        }

        [Theory(DisplayName = "Test of the SetVerticalSplit function with width definition")]
        [InlineData(22.2f, "B2", Worksheet.WorksheetPane.bottomLeft)]
        [InlineData(0f, "B2", Worksheet.WorksheetPane.bottomLeft)]
        [InlineData(500f, "B2", Worksheet.WorksheetPane.bottomLeft)]
        [InlineData(22.2f, "X1", Worksheet.WorksheetPane.bottomLeft)]
        [InlineData(0f, "A1", Worksheet.WorksheetPane.bottomLeft)]
        [InlineData(500f, "XFD1048576", Worksheet.WorksheetPane.bottomLeft)]
        [InlineData(22.2f, "B2", Worksheet.WorksheetPane.topRight)]
        [InlineData(0f, "B2", Worksheet.WorksheetPane.bottomRight)]
        [InlineData(500f, "B2", Worksheet.WorksheetPane.topLeft)]
        public void SetVerticalSplitTest(float width, string topLeftCellAddress, Worksheet.WorksheetPane activePane)
        {
            Worksheet worksheet = new Worksheet();
            AssertInitializedPaneSplit(worksheet);
            Address address = new Address(topLeftCellAddress);
            worksheet.SetVerticalSplit(width, address, activePane);
            Assert.Equal(width, worksheet.PaneSplitLeftWidth);
            Assert.Equal(address, worksheet.PaneSplitTopLeftCell);
            Assert.Equal(activePane, worksheet.ActivePane);
            Assert.Null(worksheet.FreezeSplitPanes);
            Assert.Null(worksheet.PaneSplitAddress);
            Assert.Null(worksheet.PaneSplitTopHeight);
        }

        [Theory(DisplayName = "Test of the SetVerticalSplit function with column definition")]
        [InlineData(3, false, "A1", Worksheet.WorksheetPane.bottomLeft)]
        [InlineData(10, true, "K11", Worksheet.WorksheetPane.bottomLeft)]
        [InlineData(3, false, "E2", Worksheet.WorksheetPane.bottomRight)]
        [InlineData(10, true, "L100", Worksheet.WorksheetPane.bottomRight)]
        [InlineData(3, false, "F3", Worksheet.WorksheetPane.topLeft)]
        [InlineData(10, true, "M200", Worksheet.WorksheetPane.topLeft)]
        [InlineData(3, false, "F3", Worksheet.WorksheetPane.topRight)]
        [InlineData(10, true, "M11", Worksheet.WorksheetPane.topRight)]
        public void SetVerticalSplitTest2(int columnNumber, bool freeze, string topLeftCellAddress, Worksheet.WorksheetPane activePane)
        {
            Worksheet worksheet = new Worksheet();
            AssertInitializedPaneSplit(worksheet);
            Address address = new Address(topLeftCellAddress);
            worksheet.SetVerticalSplit(columnNumber, freeze, address, activePane);
            Assert.Null(worksheet.PaneSplitLeftWidth);
            Assert.Null(worksheet.PaneSplitTopHeight);
            Address expectedAddress = new Address(columnNumber, 0);
            Assert.Equal(expectedAddress.GetAddress(), worksheet.PaneSplitAddress.Value.GetAddress());
            Assert.Equal(freeze, worksheet.FreezeSplitPanes);
            Assert.Equal(address, worksheet.PaneSplitTopLeftCell);
            Assert.Equal(activePane, worksheet.ActivePane);
        }

        [Theory(DisplayName = "Test of the failing SetVerticalSplit function")]
        [InlineData(3, false, "A1", true)]
        [InlineData(3, true, "A1", false)]
        [InlineData(100, false, "R100", true)]
        [InlineData(100, true, "R100", false)]
        public void SetVerticalSplitFailTest(int columnNumber, bool freeze, string topLeftCellAddress, bool expectedValid)
        {
            Worksheet worksheet = new Worksheet();
            AssertInitializedPaneSplit(worksheet);
            Address address = new Address(topLeftCellAddress);
            if (expectedValid)
            {
                worksheet.SetVerticalSplit(columnNumber, freeze, address, Worksheet.WorksheetPane.bottomLeft);
            }
            else
            {
                Assert.Throws<WorksheetException>(() => worksheet.SetVerticalSplit(columnNumber, freeze, address, Worksheet.WorksheetPane.bottomLeft));
            }
        }

        [Theory(DisplayName = "Test of the SetSplit function with height and width definition")]
        [InlineData(22.2f, 11.1f, "B2", Worksheet.WorksheetPane.bottomLeft)]
        [InlineData(0f, 0f, "B2", Worksheet.WorksheetPane.bottomLeft)]
        [InlineData(500f, 200f, "B2", Worksheet.WorksheetPane.bottomLeft)]
        [InlineData(22.2f, 0f, "X1", Worksheet.WorksheetPane.bottomLeft)]
        [InlineData(null, 0f, "A1", Worksheet.WorksheetPane.bottomLeft)]
        [InlineData(500f, null, "XFD1048576", Worksheet.WorksheetPane.bottomLeft)]
        [InlineData(null, 22.2f, "B2", Worksheet.WorksheetPane.topRight)]
        [InlineData(0f, null, "B2", Worksheet.WorksheetPane.bottomRight)]
        [InlineData(null, 500f, "B2", Worksheet.WorksheetPane.topLeft)]
        public void SetSplitTest(float? height, float? width, string topLeftCellAddress, Worksheet.WorksheetPane activePane)
        {
            Worksheet worksheet = new Worksheet();
            AssertInitializedPaneSplit(worksheet);
            Address address = new Address(topLeftCellAddress);
            worksheet.SetSplit(width, height, address, activePane);
            Assert.Equal(height, worksheet.PaneSplitTopHeight);
            Assert.Equal(width, worksheet.PaneSplitLeftWidth);
            Assert.Equal(address, worksheet.PaneSplitTopLeftCell);
            Assert.Equal(activePane, worksheet.ActivePane);
            Assert.Null(worksheet.FreezeSplitPanes);
            Assert.Null(worksheet.PaneSplitAddress);
           
        }

        [Theory(DisplayName = "Test of the SetSplit function with column and definition")]
        [InlineData(3, 3, false, "A1", Worksheet.WorksheetPane.bottomLeft)]
        [InlineData(10, 2, true, "K11", Worksheet.WorksheetPane.bottomLeft)]
        [InlineData(3, 1, false, "E2", Worksheet.WorksheetPane.bottomRight)]
        [InlineData(10, 99, true, "L100", Worksheet.WorksheetPane.bottomRight)]
        [InlineData(3, null, false, "F3", Worksheet.WorksheetPane.topLeft)]
        [InlineData(null, 1, true, "M200", Worksheet.WorksheetPane.topLeft)]
        [InlineData(3, null,  false, "F3", Worksheet.WorksheetPane.topRight)]
        [InlineData(null, 10, true, "M11", Worksheet.WorksheetPane.topRight)]
        public void SetSplitTest2(int? columnNumber, int? rowNumber, bool freeze, string topLeftCellAddress, Worksheet.WorksheetPane activePane)
        {
            Worksheet worksheet = new Worksheet();
            AssertInitializedPaneSplit(worksheet);
            Address address = new Address(topLeftCellAddress);
            worksheet.SetSplit(columnNumber, rowNumber, freeze, address, activePane);
            Assert.Null(worksheet.PaneSplitLeftWidth);
            Assert.Null(worksheet.PaneSplitTopHeight);
            int column = columnNumber.GetValueOrDefault(0);
            int row = rowNumber.GetValueOrDefault(0);
            Address expectedAddress = new Address(column, row);
            Assert.Equal(expectedAddress.GetAddress(), worksheet.PaneSplitAddress.Value.GetAddress());
            Assert.Equal(freeze, worksheet.FreezeSplitPanes);
            Assert.Equal(address, worksheet.PaneSplitTopLeftCell);
            Assert.Equal(activePane, worksheet.ActivePane);
        }

        [Theory(DisplayName = "Test of the failing SetSplit function")]
        [InlineData(3, 3, false, "A1", true)]
        [InlineData(3, 0, true, "A1", false)]
        [InlineData(100, 1, false, "R100", true)]
        [InlineData(100, 1, true, "R100", false)]
        [InlineData(3, 3, false, "B2", true)]
        [InlineData(3, 0, true, "B2", false)]
        [InlineData(17, 1, false, "R100", true)]
        [InlineData(16, 100, true, "R100", false)]
        [InlineData(3, null, true, "E1", true)]
        [InlineData(null, 99, true, "R100", true)]
        [InlineData(3, null, true, "A1", false)]
        [InlineData(null, 101, true, "R100", false)]
        [InlineData(null, null, true, "A1", true)]

        public void SetSplitFailTest(int? columnNumber, int? rowNumber, bool freeze, string topLeftCellAddress, bool expectedValid)
        {
            Worksheet worksheet = new Worksheet();
            AssertInitializedPaneSplit(worksheet);
            Address address = new Address(topLeftCellAddress);
            if (expectedValid)
            {
                worksheet.SetSplit(columnNumber, rowNumber, freeze, address, Worksheet.WorksheetPane.bottomLeft);
            }
            else
            {
                Assert.Throws<WorksheetException>(() => worksheet.SetSplit(columnNumber, rowNumber, freeze, address, Worksheet.WorksheetPane.bottomLeft));
            }
        }

        [Fact(DisplayName = "Test of the ResetSplit function on a horizontal split with a height definition")]
        public void ResetSplitTest()
        {
            Worksheet worksheet = new Worksheet();
            AssertInitializedPaneSplit(worksheet);
            worksheet.SetHorizontalSplit(22.2f, new Address("A1"), Worksheet.WorksheetPane.bottomLeft);
            worksheet.ResetSplit();
            AssertInitializedPaneSplit(worksheet);
        }

        [Fact(DisplayName = "Test of the ResetSplit function on a horizontal split with a row definition")]
        public void ResetSplitTest2()
        {
            Worksheet worksheet = new Worksheet();
            AssertInitializedPaneSplit(worksheet);
            worksheet.SetHorizontalSplit(5, true, new Address("R6"), Worksheet.WorksheetPane.bottomLeft);
            worksheet.ResetSplit();
            AssertInitializedPaneSplit(worksheet);
        }


        [Fact(DisplayName = "Test of the ResetSplit function on a vertical split with a width definition")]
        public void ResetSplitTest3()
        {
            Worksheet worksheet = new Worksheet();
            AssertInitializedPaneSplit(worksheet);
            worksheet.SetVerticalSplit(22.2f, new Address("A1"), Worksheet.WorksheetPane.bottomLeft);
            worksheet.ResetSplit();
            AssertInitializedPaneSplit(worksheet);
        }

        [Fact(DisplayName = "Test of the ResetSplit function on a vertical split with a column definition")]
        public void ResetSplitTest4()
        {
            Worksheet worksheet = new Worksheet();
            AssertInitializedPaneSplit(worksheet);
            worksheet.SetVerticalSplit(5, true, new Address("R6"), Worksheet.WorksheetPane.bottomLeft);
            worksheet.ResetSplit();
            AssertInitializedPaneSplit(worksheet);
        }

        [Fact(DisplayName = "Test of the ResetSplit function on a split with a width and height definition")]
        public void ResetSplitTest5()
        {
            Worksheet worksheet = new Worksheet();
            AssertInitializedPaneSplit(worksheet);
            worksheet.SetSplit(22.2f, 22.2f, new Address("A1"), Worksheet.WorksheetPane.bottomLeft);
            worksheet.ResetSplit();
            AssertInitializedPaneSplit(worksheet);
        }

        [Fact(DisplayName = "Test of the ResetSplit function on a split with a column and row definition")]
        public void ResetSplitTest6()
        {
            Worksheet worksheet = new Worksheet();
            AssertInitializedPaneSplit(worksheet);
            worksheet.SetSplit(5, 5, true, new Address("R6"), Worksheet.WorksheetPane.bottomLeft);
            worksheet.ResetSplit();
            AssertInitializedPaneSplit(worksheet);
        }

        private static void AssertInitializedPaneSplit(Worksheet worksheet)
        {
            Assert.Null(worksheet.PaneSplitLeftWidth);
            Assert.Null(worksheet.PaneSplitTopHeight);
            Assert.Null(worksheet.PaneSplitTopLeftCell);
            Assert.Null(worksheet.ActivePane);
            Assert.Null(worksheet.FreezeSplitPanes);
            Assert.Null(worksheet.PaneSplitAddress);
        }


    }
}
