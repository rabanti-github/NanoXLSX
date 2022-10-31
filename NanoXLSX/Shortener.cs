/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2022
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Shared.Exceptions;
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
        private readonly Workbook workbookReference;

        /// <summary>
        /// Constructor with workbook reference
        /// </summary>
        /// <param name="reference">Workbook reference</param>
        public Shortener(Workbook reference)
        {
            this.workbookReference = reference;
            this.currentWorksheet = reference.CurrentWorksheet;
        }

        /// <summary>
        /// Sets the worksheet accessed by the shortener
        /// </summary>
        /// <param name="worksheet">Current worksheet</param>
        public void SetCurrentWorksheet(Worksheet worksheet)
        {
            workbookReference.SetCurrentWorksheet(worksheet);
            currentWorksheet = worksheet;
        }

        /// <summary>
        /// Sets the worksheet accessed by the shortener, invoked by the workbook
        /// </summary>
        /// <param name="worksheet">Current worksheet</param>
        internal void SetCurrentWorksheetInternal(Worksheet worksheet)
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
        /// <param name="keepColumnPosition">If true, the column position is preserved, otherwise set to 0</param>
        public void Down(int numberOfRows, bool keepColumnPosition = false)
        {
            NullCheck();
            currentWorksheet.GoToNextRow(numberOfRows, keepColumnPosition);
        }

        /// <summary>
        /// Moves the cursor one row up
        /// </summary>
        /// <remarks>An exception will be thrown if the row number is below 0/></remarks>
        public void Up()
        {
            NullCheck();
            currentWorksheet.GoToNextRow(-1);
        }

        /// <summary>
        /// Moves the cursor the number of defined rows up
        /// </summary>
        /// <param name="numberOfRows">Number of rows to move</param>
        /// <param name="keepColumnosition">If true, the column position is preserved, otherwise set to 0</param>
        /// <remarks>An exception will be thrown if the row number is below 0. Values can be also negative. However, this is the equivalent of the function <see cref="Down(int, bool)"/></remarks>
        public void Up(int numberOfRows, bool keepColumnosition = false)
        {
            NullCheck();
            currentWorksheet.GoToNextRow(-1*numberOfRows, keepColumnosition);
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
        /// <param name="keepRowPosition">If true, the row position is preserved, otherwise set to 0</param>
        public void Right(int numberOfColumns, bool keepRowPosition = false)
        {
            NullCheck();
            currentWorksheet.GoToNextColumn(numberOfColumns, keepRowPosition);
        }

        /// <summary>
        /// Moves the cursor one column to the left
        /// </summary>
        /// <remarks>An exception will be thrown if the column number is below 0</remarks>
        public void Left()
        {
            NullCheck();
            currentWorksheet.GoToNextColumn(-1);
        }

        /// <summary>
        /// Moves the cursor the number of defined columns to the left
        /// </summary>
        /// <param name="numberOfColumns">Number of columns to move</param>
        /// <param name="keepRowRowPosition">If true, the row position is preserved, otherwise set to 0</param>
        /// <remarks>An exception will be thrown if the column number is below 0. Values can be also negative. However, this is the equivalent of the function <see cref="Right(int, bool)"/></remarks>
        public void Left(int numberOfColumns, bool keepRowRowPosition = false)
        {
            NullCheck();
            currentWorksheet.GoToNextColumn(-1*numberOfColumns, keepRowRowPosition);
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
