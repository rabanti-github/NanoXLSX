﻿using NanoXLSX;
using NanoXLSX.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace NanoXLSX_Test.Worksheets
{
    public class RowTest
    {
        public enum RowProperty
        {
            Hidden,
            Height
        }

        [Fact(DisplayName = "Test of the AddHiddenRow function with a row number")]
        public void AddHiddenRowTest()
        {
            Worksheet worksheet = new Worksheet();
            Assert.Empty(worksheet.HiddenRows);
            worksheet.AddHiddenRow(2);
            Assert.Single(worksheet.HiddenRows);
            Assert.Contains(worksheet.HiddenRows, item => item.Key == 2);
            Assert.True(worksheet.HiddenRows[2]);
            worksheet.AddHiddenRow(2); // Should not add an additional entry
            Assert.Single(worksheet.HiddenRows);
        }

        [Theory(DisplayName = "Test of the failing AddHiddenRow function with an invalid row number")]
        [InlineData(-1)]
        [InlineData(-100)]
        [InlineData(1048576)]
        public void AddHiddenRowFailTest(int value)
        {
            Worksheet worksheet = new Worksheet();
            Assert.Throws<RangeException>(() => worksheet.AddHiddenRow(value));
        }

        [Fact(DisplayName = "Test of the GetLastRowNumber function with an empty worksheet")]
        public void GetLastRowNumberTest()
        {
            Worksheet worksheet = new Worksheet();
            int row = worksheet.GetLastRowNumber();
            Assert.Equal(-1, row);
        }

        [Theory(DisplayName = "Test of the GetLastRowNumber function with defined rows on an empty worksheet")]
        [InlineData(RowProperty.Height)]
        [InlineData(RowProperty.Hidden)]
        public void GetLastRowNumberTest2(RowProperty rowProperty)
        {
            Worksheet worksheet = new Worksheet();
            if (rowProperty == RowProperty.Hidden)
            {
                worksheet.AddHiddenRow(0);
                worksheet.AddHiddenRow(1);
                worksheet.AddHiddenRow(2);
            }
            else
            {
                worksheet.SetRowHeight(0, 22.2f);
                worksheet.SetRowHeight(1, 33.3f);
                worksheet.SetRowHeight(2, 44.4f);
            }
            int row = worksheet.GetLastRowNumber();
            Assert.Equal(2, row);
        }

        [Theory(DisplayName = "Test of the GetLastRowNumber function with defined rows on an empty worksheet, where the row definition has gaps")]
        [InlineData(RowProperty.Height)]
        [InlineData(RowProperty.Hidden)]
        public void GetLastRowNumberTest3(RowProperty rowProperty)
        {
            Worksheet worksheet = new Worksheet();
            if (rowProperty == RowProperty.Hidden)
            {
                worksheet.AddHiddenRow(0);
                worksheet.AddHiddenRow(1);
                worksheet.AddHiddenRow(10);
            }
            else
            {
                worksheet.SetRowHeight(0, 22.2f);
                worksheet.SetRowHeight(1, 33.3f);
                worksheet.SetRowHeight(10, 44.4f);
            }
            int row = worksheet.GetLastRowNumber();
            Assert.Equal(10, row);
        }

        [Theory(DisplayName = "Test of the GetLastRowNumber function with defined rows where cells are defined below the last row")]
        [InlineData(RowProperty.Height)]
        [InlineData(RowProperty.Hidden)]
        public void GetLastRowNumberTest4(RowProperty rowProperty)
        {
            Worksheet worksheet = new Worksheet();
            if (rowProperty == RowProperty.Hidden)
            {
                worksheet.AddHiddenRow(0);
                worksheet.AddHiddenRow(1);
                worksheet.AddHiddenRow(10);
            }
            else
            {
                worksheet.SetRowHeight(0, 22.2f);
                worksheet.SetRowHeight(1, 33.3f);
                worksheet.SetRowHeight(10, 44.4f);
            }
            worksheet.AddCell("test", "E5");
            int row = worksheet.GetLastRowNumber();
            Assert.Equal(10, row);
        }

        [Theory(DisplayName = "Test of the GetLastRowNumber function with defined rows where cells are defined above the last row")]
        [InlineData(RowProperty.Height)]
        [InlineData(RowProperty.Hidden)]
        public void GetLastRowNumberTest5(RowProperty rowProperty)
        {
            Worksheet worksheet = new Worksheet();
            if (rowProperty == RowProperty.Hidden)
            {
                worksheet.AddHiddenRow(0);
                worksheet.AddHiddenRow(1);
                worksheet.AddHiddenRow(2);
            }
            else
            {
                worksheet.SetRowHeight(0, 22.2f);
                worksheet.SetRowHeight(1, 33.3f);
                worksheet.SetRowHeight(2, 44.4f);
            }
            worksheet.AddCell("test", "F5");
            int row = worksheet.GetLastRowNumber();
            Assert.Equal(4, row);
        }

        [Fact(DisplayName = "Test of the GetLastDataRowNumber function with an empty worksheet")]
        public void GetLastDataRowNumberTest()
        {
            Worksheet worksheet = new Worksheet();
            int row = worksheet.GetLastDataRowNumber();
            Assert.Equal(-1, row);
        }

        [Theory(DisplayName = "Test of the GetLastDataRowNumber function with defined rows on an empty worksheet")]
        [InlineData(RowProperty.Height)]
        [InlineData(RowProperty.Hidden)]
        public void GetLastDataRowNumberTest2(RowProperty rowProperty)
        {
            Worksheet worksheet = new Worksheet();
            if (rowProperty == RowProperty.Hidden)
            {
                worksheet.AddHiddenRow(0);
                worksheet.AddHiddenRow(1);
                worksheet.AddHiddenRow(2);
            }
            else
            {
                worksheet.SetRowHeight(0, 22.2f);
                worksheet.SetRowHeight(1, 33.3f);
                worksheet.SetRowHeight(2, 44.4f);
            }
            int row = worksheet.GetLastDataRowNumber();
            Assert.Equal(-1, row);
        }

        [Theory(DisplayName = "Test of the GetLastDataRowNumber function with defined rows where cells are defined below the last row")]
        [InlineData(RowProperty.Height)]
        [InlineData(RowProperty.Hidden)]
        public void GetLastDataRowNumberTest3(RowProperty rowProperty)
        {
            Worksheet worksheet = new Worksheet();
            if (rowProperty == RowProperty.Hidden)
            {
                worksheet.AddHiddenRow(0);
                worksheet.AddHiddenRow(1);
                worksheet.AddHiddenRow(10);
            }
            else
            {
                worksheet.SetRowHeight(0, 22.2f);
                worksheet.SetRowHeight(1, 33.3f);
                worksheet.SetRowHeight(10, 44.4f);
            }

            worksheet.AddCell("test", "E5");
            int row = worksheet.GetLastDataRowNumber();
            Assert.Equal(4, row);
        }

        [Theory(DisplayName = "Test of the GetLastDataRowNumber function with defined rows where cells are defined above the last row")]
        [InlineData(RowProperty.Height)]
        [InlineData(RowProperty.Hidden)]
        public void GetLastDataRowNumberTest4(RowProperty rowProperty)
        {
            Worksheet worksheet = new Worksheet();
            if (rowProperty == RowProperty.Hidden)
            {
                worksheet.AddHiddenRow(0);
                worksheet.AddHiddenRow(1);
                worksheet.AddHiddenRow(2);
            }
            else
            {
                worksheet.SetRowHeight(0, 22.2f);
                worksheet.SetRowHeight(1, 33.3f);
                worksheet.SetRowHeight(3, 44.4f);
            }

            worksheet.AddCell("test", "F5");
            int row = worksheet.GetLastDataRowNumber();
            Assert.Equal(4, row);
        }

        [Fact(DisplayName = "Test of the GetCurrentColumnNumber function")]
        public void GetCurrentRowNumberTest()
        {
            Worksheet worksheet = new Worksheet();
            Assert.Equal(0, worksheet.GetCurrentRowNumber());
            worksheet.CurrentCellDirection = Worksheet.CellDirection.RowToRow;
            worksheet.AddNextCell("test");
            worksheet.AddNextCell("test");
            Assert.Equal(2, worksheet.GetCurrentRowNumber());
            worksheet.CurrentCellDirection = Worksheet.CellDirection.ColumnToColumn;
            worksheet.AddNextCell("test");
            worksheet.AddNextCell("test");
            Assert.Equal(2, worksheet.GetCurrentRowNumber()); // should not change
            worksheet.GoToNextRow();
            Assert.Equal(3, worksheet.GetCurrentRowNumber());
            worksheet.GoToNextRow(2);
            Assert.Equal(5, worksheet.GetCurrentRowNumber());
            worksheet.GoToNextColumn(2);
            Assert.Equal(0, worksheet.GetCurrentRowNumber()); // should reset
        }

        [Fact(DisplayName = "Test of the GoToNextRow function")]
        public void GoToNextRowTest()
        {
            Worksheet worksheet = new Worksheet();
            Assert.Equal(0, worksheet.GetCurrentRowNumber());
            worksheet.GoToNextRow();
            Assert.Equal(1, worksheet.GetCurrentRowNumber());
            worksheet.GoToNextRow(5);
            Assert.Equal(6, worksheet.GetCurrentRowNumber());
            worksheet.GoToNextRow(-2);
            Assert.Equal(4, worksheet.GetCurrentRowNumber());
            worksheet.GoToNextRow(0);
            Assert.Equal(4, worksheet.GetCurrentRowNumber());
        }

        [Theory(DisplayName = "Test of the failing GoToNextRow function on invalid values")]
        [InlineData(0, -1)]
        [InlineData(10, -12)]
        [InlineData(0, 1048576)]
        [InlineData(0, 1248575)]
        public void GoToNextRowTest2(int initialValue, int value)
        {
            Worksheet worksheet = new Worksheet();
            worksheet.SetCurrentRowNumber(initialValue);
            Assert.Equal(initialValue, worksheet.GetCurrentRowNumber());
            Assert.Throws<RangeException>(() => worksheet.GoToNextRow(value));
        }

        [Fact(DisplayName = "Test of the RemoveRowHeight function")]
        public void RemoveRowHeightTest()
        {
            Worksheet worksheet = new Worksheet();
            worksheet.SetRowHeight(2, 22.2f);
            worksheet.SetRowHeight(4, 33.3f);
            Assert.Equal(2, worksheet.RowHeights.Count);
            worksheet.RemoveRowHeight(2);
            Assert.Single(worksheet.RowHeights);
            worksheet.RemoveRowHeight(3); // Should not cause anything
            worksheet.RemoveRowHeight(-1); // Should not cause anything
            Assert.Single(worksheet.RowHeights);
        }

        [Theory(DisplayName = "Test of the SetCurrentRowNumber function")]
        [InlineData(0)]
        [InlineData(3)]
        [InlineData(1048575)]
        public void SetCurrentRowNumberTest(int row)
        {
            Worksheet worksheet = new Worksheet();
            Assert.Equal(0, worksheet.GetCurrentRowNumber());
            worksheet.GoToNextRow();
            worksheet.SetCurrentRowNumber(row);
            Assert.Equal(row, worksheet.GetCurrentRowNumber());
        }

        [Theory(DisplayName = "Test of the failing SetCurrentRowNumber function")]
        [InlineData(-1)]
        [InlineData(-10)]
        [InlineData(1048576)]
        public void SetCurrentRowNumberFailTest(int row)
        {
            Worksheet worksheet = new Worksheet();
            Assert.Throws<RangeException>(() => worksheet.SetCurrentRowNumber(row));
        }

        [Theory(DisplayName = "Test of the SetRowHeight function")]
        [InlineData(0f)]
        [InlineData(0.1f)]
        [InlineData(10f)]
        [InlineData(255f)]
        public void SetRowHeightTest(float height)
        {
            Worksheet worksheet = new Worksheet();
            Assert.Empty(worksheet.RowHeights);
            worksheet.SetRowHeight(0, height);
            Assert.Single(worksheet.RowHeights);
            Assert.Equal(height, worksheet.RowHeights[0]);
            worksheet.SetRowHeight(0, Worksheet.DEFAULT_ROW_HEIGHT);
            Assert.Single(worksheet.RowHeights); // No removal so far
            Assert.Equal(Worksheet.DEFAULT_ROW_HEIGHT, worksheet.RowHeights[0]);
        }

        [Theory(DisplayName = "Test of the failing SetRowHeight function")]
        [InlineData(-1, 0f)]
        [InlineData(1048576, 0.0f)]
        [InlineData(0, -10f)]
        [InlineData(0, 409.51f)]
        [InlineData(0, 500f)]
        public void SetRowHeightFailTest(int rowNumber, float height)
        {
            Worksheet worksheet = new Worksheet();
            Assert.Throws<RangeException>(() => worksheet.SetRowHeight(rowNumber, height));
        }


    }
}
