/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2021
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Exceptions;
using NanoXLSX.Styles;

namespace NanoXLSX
{
    /// <summary>
    /// Class to provide access to the current worksheet with a shortened syntax. Note: The WS object can be null if the workbook was created without a worksheet. 
    /// The object will be available as soon as the current worksheet is defined
    /// </summary>
    public class Shortener
    {
        private Worksheet currentWorksheet;

        /// <summary>
        /// Sets the worksheet accessed by the shortener
        /// </summary>
        /// <param name="worksheet">Current worksheet</param>
        public void SetCurrentWorksheet(Worksheet worksheet)
        {
            currentWorksheet = worksheet;
        }

        /// <summary>
        /// Sets a value into the current cell and moves the cursor to the next cell (column or row depending on the defined cell direction)
        /// </summary>
        /// <exception cref="WorksheetException">Throws a WorksheetException if no worksheet was defined</exception>
        /// <param name="cellValue">Value to set</param>
        public void Value(object cellValue)
        {
            NullCheck();
            currentWorksheet.AddNextCell(cellValue);
        }

        /// <summary>
        /// Sets a value with style into the current cell and moves the cursor to the next cell (column or row depending on the defined cell direction)
        /// </summary>
        /// <exception cref="WorksheetException">Throws a WorksheetException if no worksheet was defined</exception>
        /// <param name="cellValue">Value to set</param>
        /// <param name="style">Style to apply</param>
        public void Value(object cellValue, Style style)
        {
            NullCheck();
            currentWorksheet.AddNextCell(cellValue, style);
        }

        /// <summary>
        /// Sets a formula into the current cell and moves the cursor to the next cell (column or row depending on the defined cell direction)
        /// </summary>
        /// <exception cref="WorksheetException">Throws a WorksheetException if no worksheet was defined</exception>
        /// <param name="cellFormula">Formula to set</param>
        public void Formula(string cellFormula)
        {
            NullCheck();
            currentWorksheet.AddNextCellFormula(cellFormula);
        }

        /// <summary>
        /// Sets a formula with style into the current cell and moves the cursor to the next cell (column or row depending on the defined cell direction)
        /// </summary>
        /// <exception cref="WorksheetException">Throws a WorksheetException if no worksheet was defined</exception>
        /// <param name="cellFormula">Formula to set</param>
        /// <param name="style">Style to apply</param>
        public void Formula(string cellFormula, Style style)
        {
            NullCheck();
            currentWorksheet.AddNextCellFormula(cellFormula, style);
        }

        /// <summary>
        /// Moves the cursor one row down
        /// </summary>
        public void Down()
        {
            NullCheck();
            currentWorksheet.GoToNextRow();
        }

        /// <summary>
        /// Moves the cursor the number of defined rows down
        /// </summary>
        /// <param name="numberOfRows">Number of rows to move</param>
        public void Down(int numberOfRows)
        {
            NullCheck();
            currentWorksheet.GoToNextRow(numberOfRows);
        }

        /// <summary>
        /// Moves the cursor one column to the right
        /// </summary>
        public void Right()
        {
            NullCheck();
            currentWorksheet.GoToNextColumn();
        }

        /// <summary>
        /// Moves the cursor the number of defined columns to the right
        /// </summary>
        /// <param name="numberOfColumns">Number of columns to move</param>
        public void Right(int numberOfColumns)
        {
            NullCheck();
            currentWorksheet.GoToNextColumn(numberOfColumns);
        }

        /// <summary>
        /// Internal method to check whether the worksheet is null
        /// </summary>
        private void NullCheck()
        {
            if (currentWorksheet == null)
            {
                throw new WorksheetException("No worksheet was defined");
            }
        }


    }
}
