/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2021
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using NanoXLSX.Exceptions;
using NanoXLSX.Styles;
using FormatException = NanoXLSX.Exceptions.FormatException;

namespace NanoXLSX
{
    /// <summary>
    /// Class representing a worksheet of a workbook
    /// </summary>
    public class Worksheet
    {
        #region constants
        /// <summary>
        /// Threshold, using when floats are compared
        /// </summary>
        private const float FLOAT_TRESHOLD = 0.0001f;
        /// <summary>
        /// Maximum number of characters a worksheet name can have
        /// </summary>
        public static readonly int MAX_WORKSHEER_NAME_LENGTH = 31;
        /// <summary>
        /// Default column width as constant
        /// </summary>
        public static readonly float DEFAULT_COLUMN_WIDTH = 10f;
        /// <summary>
        /// Default row height as constant
        /// </summary>
        public static readonly float DEFAULT_ROW_HEIGHT = 15f;
        /// <summary>
        /// Maximum column number (zero-based) as constant
        /// </summary>
        public static readonly int MAX_COLUMN_NUMBER = 16383;
        /// <summary>
        /// Minimum column number (zero-based) as constant
        /// </summary>
        public static readonly int MIN_COLUMN_NUMBER = 0;
        /// <summary>
        /// Minimum column width as constant
        /// </summary>
        public static readonly float MIN_COLUMN_WIDTH = 0f;
        /// <summary>
        /// Minimum row height as constant
        /// </summary>
        public static readonly float MIN_ROW_HEIGHT = 0f;
        /// <summary>
        /// Maximum column width as constant
        /// </summary>
        public static readonly float MAX_COLUMN_WIDTH = 255f;
        /// <summary>
        /// Maximum row number (zero-based) as constant
        /// </summary>
        public static readonly int MAX_ROW_NUMBER = 1048575;
        /// <summary>
        /// Minimum row number (zero-based) as constant
        /// </summary>
        public static readonly int MIN_ROW_NUMBER = 0;
        /// <summary>
        /// Maximum row height as constant
        /// </summary>
        public static readonly float MAX_ROW_HEIGHT = 409.5f;
        #endregion

        #region enums
        /// <summary>
        /// Enum to define the direction when using AddNextCell method
        /// </summary>
        public enum CellDirection
        {
            /// <summary>The next cell will be on the same row (A1,B1,C1...)</summary>
            ColumnToColumn,
            /// <summary>The next cell will be on the same column (A1,A2,A3...)</summary>
            RowToRow,
            /// <summary>The address of the next cell will be not changed when adding a cell (for manual definition of cell addresses)</summary>
            Disabled
        }

        /// <summary>
        /// Enum to define the possible protection types when protecting a worksheet
        /// </summary>
        public enum SheetProtectionValue
        {
            // sheet, // Is always on 1 if protected
            /// <summary>If selected, the user can edit objects if the worksheets is protected</summary>
            objects,
            /// <summary>If selected, the user can edit scenarios if the worksheets is protected</summary>
            scenarios,
            /// <summary>If selected, the user can format cells if the worksheets is protected</summary>
            formatCells,
            /// <summary>If selected, the user can format columns if the worksheets is protected</summary>
            formatColumns,
            /// <summary>If selected, the user can format rows if the worksheets is protected</summary>
            formatRows,
            /// <summary>If selected, the user can insert columns if the worksheets is protected</summary>
            insertColumns,
            /// <summary>If selected, the user can insert rows if the worksheets is protected</summary>
            insertRows,
            /// <summary>If selected, the user can insert hyper links if the worksheets is protected</summary>
            insertHyperlinks,
            /// <summary>If selected, the user can delete columns if the worksheets is protected</summary>
            deleteColumns,
            /// <summary>If selected, the user can delete rows if the worksheets is protected</summary>
            deleteRows,
            /// <summary>If selected, the user can select locked cells if the worksheets is protected</summary>
            selectLockedCells,
            /// <summary>If selected, the user can sort cells if the worksheets is protected</summary>
            sort,
            /// <summary>If selected, the user can use auto filters if the worksheets is protected</summary>
            autoFilter,
            /// <summary>If selected, the user can use pivot tables if the worksheets is protected</summary>
            pivotTables,
            /// <summary>If selected, the user can select unlocked cells if the worksheets is protected</summary>
            selectUnlockedCells
        }

        /// <summary>
        /// Enum to define the pane position or active pane in a slip worksheet
        /// </summary>
        public enum WorksheetPane
        {
            /// <summary>The pane is located in the bottom right of the split worksheet</summary>
            bottomRight,
            /// <summary>The pane is located in the top right of the split worksheet</summary>
            topRight,
            /// <summary>The pane is located in the bottom left of the split worksheet</summary>
            bottomLeft,
            /// <summary>The pane is located in the top left of the split worksheet</summary>
            topLeft
        }
        #endregion

        #region privateFields
        private Style activeStyle;
        private Range? autoFilterRange;
        private readonly Dictionary<string, Cell> cells;
        private readonly Dictionary<int, Column> columns;
        private string sheetName;
        private int currentRowNumber;
        private int currentColumnNumber;
        private float defaultRowHeight;
        private float defaultColumnWidth;
        private readonly Dictionary<int, float> rowHeights;
        private readonly Dictionary<int, bool> hiddenRows;
        private readonly Dictionary<string, Range> mergedCells;
        private readonly List<SheetProtectionValue> sheetProtectionValues;
        private bool useActiveStyle;
        private string sheetProtectionPassword;
        private Range? selectedCells;
        private bool? freezeSplitPanes;
        private float? paneSplitLeftWidth;
        private float? paneSplitTopHeight;
        private Address? paneSplitTopLeftCell;
        private Address? paneSplitAddress;
        private WorksheetPane? activePane;
        private int sheetID;
        #endregion

        #region properties
        /// <summary>
        /// Gets the range of the auto-filter. Wrapped to Nullable to provide null as value. If null, no auto-filter is applied
        /// </summary>
        public Range? AutoFilterRange
        {
            get { return autoFilterRange; }
        }

        /// <summary>
        /// Gets the cells of the worksheet as dictionary with the cell address as key and the cell object as value
        /// </summary>
        public Dictionary<string, Cell> Cells
        {
            get { return cells; }
        }

        /// <summary>
        /// Gets all columns with non-standard properties, like auto filter applied or a special width as dictionary with the zero-based column index as key and the column object as value
        /// </summary>
        public Dictionary<int, Column> Columns
        {
            get { return columns; }
        }

        /// <summary>
        /// Gets or sets the direction when using AddNextCell method
        /// </summary>
        public CellDirection CurrentCellDirection { get; set; }

        /// <summary>
        /// Gets or sets the default column width
        /// </summary>
        /// <exception cref="RangeException">Throws a RangeException exception if the passed width is out of range (set)</exception>
        public float DefaultColumnWidth
        {
            get { return defaultColumnWidth; }
            set
            {
                if (value < MIN_COLUMN_WIDTH || value > MAX_COLUMN_WIDTH)
                {
                    throw new RangeException(RangeException.GENERAL, "The passed default column width is out of range (" + MIN_COLUMN_WIDTH + " to " + MAX_COLUMN_WIDTH + ")");
                }
                defaultColumnWidth = value;
            }
        }

        /// <summary>
        /// Gets or sets the default Row height
        /// </summary>
        /// <exception cref="RangeException">Throws a RangeException exception if the passed height is out of range (set)</exception>
        public float DefaultRowHeight
        {
            get { return defaultRowHeight; }
            set
            {
                if (value < MIN_ROW_HEIGHT || value > MAX_ROW_HEIGHT)
                {
                    throw new RangeException(RangeException.GENERAL, "The passed default row height is out of range (" + MIN_ROW_HEIGHT + " to " + MAX_ROW_HEIGHT + ")");
                }
                defaultRowHeight = value;
            }
        }

        /// <summary>
        /// Gets the hidden rows as dictionary with the zero-based row number as key and a boolean as value. True indicates hidden, false visible.
        /// </summary>
        /// <remarks>Entries with the value false are not affecting the worksheet. These entries can be removed</remarks>
        public Dictionary<int, bool> HiddenRows
        {
            get { return hiddenRows; }
        }

        /// <summary>
        /// Gets the merged cells (only references) as dictionary with the cell address as key and the range object as value
        /// </summary>
        public Dictionary<string, Range> MergedCells
        {
            get { return mergedCells; }
        }

        /// <summary>
        /// Gets defined row heights as dictionary with the zero-based row number as key and the height (float from 0 to 409.5) as value
        /// </summary>
        public Dictionary<int, float> RowHeights
        {
            get { return rowHeights; }
        }

        /// <summary>
        /// Gets the cell range of selected cells of this worksheet. Null if no cells are selected
        /// </summary>
        public Range? SelectedCells
        {
            get { return selectedCells; }
        }

        /// <summary>
        /// Gets or sets the internal ID of the worksheet
        /// </summary>
        public int SheetID { get => sheetID;
            set
            {
                if (value < 1)
                {
                    throw new FormatException("InvalidIDException", "The ID " + value + " is invalid. Worksheet IDs must be >0");
                }
                sheetID = value;
            } 
        }

        /// <summary>
        /// Gets or sets the name of the worksheet
        /// </summary>
        public string SheetName
        {
            get { return sheetName; }
            set { SetSheetname(value); }
        }

        /// <summary>
        /// Gets the password used for sheet protection. See <see cref="SetSheetProtectionPassword"/> to set the password
        /// </summary>
        public string SheetProtectionPassword
        {
            get { return sheetProtectionPassword; }
        }

        /// <summary>
        /// Gets the list of SheetProtectionValues. These values define the allowed actions if the worksheet is protected
        /// </summary>
        public List<SheetProtectionValue> SheetProtectionValues
        {
            get { return sheetProtectionValues; }
        }

        /// <summary>
        /// Gets or sets whether the worksheet is protected. If true, protection is enabled
        /// </summary>
        public bool UseSheetProtection { get; set; }

        /// <summary>
        /// Gets or sets the Reference to the parent Workbook
        /// </summary>
        public Workbook WorkbookReference { get; set; }

        /// <summary>
        /// Gets or sets whether the worksheet is hidden. If true, the worksheet is not listed in the worksheet tabs of the workbook
        /// </summary>
        public bool Hidden { get; set; }

        /// <summary>
        /// Gets the height of the upper, horizontal split pane, measured from the top of the window.<br/>
        /// The value is nullable. If null, no horizontal split of the worksheet is applied.<br/>
        /// The value is only applicable to split the worksheet into panes, but not to freeze them.<br/>
        /// See also: <see cref="PaneSplitAddress"/>
        /// </summary>
        /// <remarks>Note: This value will be modified to the Excel-internal representation, 
        /// calculated by <see cref="Utils.GetInternalPaneSplitHeight(float)"/>.</remarks>
        public float? PaneSplitTopHeight
        {
            get { return paneSplitTopHeight; }
        }

        /// <summary>
        /// Gets the width of the left, vertical split pane, measured from the left of the window.<br/>
        /// The value is nullable. If null, no vertical split of the worksheet is applied<br/>
        /// The value is only applicable to split the worksheet into panes, but not to freeze them.<br/>
        /// See also: <see cref="PaneSplitAddress"/>
        /// </summary>
        /// <remarks>Note: This value will be modified to the Excel-internal representation, 
        /// calculated by <see cref="Utils.GetInternalColumnWidth(float, float, float)"/>.</remarks>
        public float? PaneSplitLeftWidth
        {
            get { return paneSplitLeftWidth; }
        }

        /// <summary>
        /// Gets whether split panes are frozen.<br/>
        /// The value is nullable. If null, no freezing is applied. This property also does not apply if <see cref="PaneSplitAddress"/> is null
        /// </summary>
        public bool? FreezeSplitPanes
        {
            get { return freezeSplitPanes; }
        }

        /// <summary>
        /// Gets the Top Left cell address of the bottom right pane if applicable and splitting is applied.<br/>
        /// The column is only relevant for vertical split, whereas the row component is only relevant for a horizontal split.<br/>
        /// The value is nullable. If null, no splitting was defined.
        /// </summary>
        public Address? PaneSplitTopLeftCell
        {
            get { return paneSplitTopLeftCell; }
        }

        /// <summary>
        /// Gets the split address for frozen panes or if pane split was defined in number of columns and / or rows.<br/> 
        /// For vertical splits, only the column component is considered. For horizontal splits, only the row component is considered.<br/>
        /// The value is nullable. If null, no frozen panes or split by columns / rows are applied to the worksheet. 
        /// However, splitting can still be applied, if the value is defined in characters.<br/>
        /// See also: <see cref="PaneSplitLeftWidth"/> and <see cref="PaneSplitTopHeight"/> for splitting in characters (without freezing)
        /// </summary>
        public Address? PaneSplitAddress
        {
            get { return paneSplitAddress; }
        }


        /// <summary>
        /// Gets the active Pane is splitting is applied.<br/>
        /// The value is nullable. If null, no splitting was defined
        /// </summary>
        public WorksheetPane? ActivePane
        {
            get { return activePane; }
        }

        /// <summary>
        /// Gets the active Style of the worksheet. If null, no style is defined as active
        /// </summary>
        public Style ActiveStyle
        {
            get { return activeStyle; }
        }

        #endregion


        #region constructors
        /// <summary>
        /// Default Constructor
        /// </summary>
        public Worksheet()
        {
            CurrentCellDirection = CellDirection.ColumnToColumn;
            cells = new Dictionary<string, Cell>();
            currentRowNumber = 0;
            currentColumnNumber = 0;
            defaultColumnWidth = DEFAULT_COLUMN_WIDTH;
            defaultRowHeight = DEFAULT_ROW_HEIGHT;
            rowHeights = new Dictionary<int, float>();
            mergedCells = new Dictionary<string, Range>();
            sheetProtectionValues = new List<SheetProtectionValue>();
            hiddenRows = new Dictionary<int, bool>();
            columns = new Dictionary<int, Column>();
            activeStyle = null;
            WorkbookReference = null;
        }

        /// <summary>
        /// Constructor with name and sheet ID
        /// </summary>
        /// <param name="name">Name of the worksheet</param>
        /// <param name="id">ID of the worksheet (for internal use)</param>
        /// <param name="reference">Reference to the parent Workbook</param>
        public Worksheet(string name, int id, Workbook reference)
            : this()
        {
            SetSheetname(name);
            SheetID = id;
            WorkbookReference = reference;
        }

        #endregion

        #region methods_AddNextCell

        /// <summary>
        /// Adds an object to the next cell position. If the type of the value does not match with one of the supported data types, it will be casted to a String. 
        /// A prepared object of the type Cell will not be casted but adjusted
        /// </summary>
        /// <remarks>Recognized are the following data types: Cell (prepared object), string, int, double, float, long, DateTime, TimeSpan, bool. 
        /// All other types will be casted into a string using the default ToString() method</remarks>
        /// <param name="value">Unspecified value to insert</param>
        /// <exception cref="RangeException">Throws a RangeException if the next cell is out of range (on row or column)</exception>
        public void AddNextCell(object value)
        {
            AddNextCell(CastValue(value, currentColumnNumber, currentRowNumber), true, null);
        }


        /// <summary>
        /// Adds an object to the next cell position. If the type of the value does not match with one of the supported data types, it will be casted to a String. 
        /// A prepared object of the type Cell will not be casted but adjusted
        /// </summary>
        /// <remarks>Recognized are the following data types: Cell (prepared object), string, int, double, float, long, DateTime, TimeSpan, bool. 
        /// All other types will be casted into a string using the default ToString() method</remarks>
        /// <param name="value">Unspecified value to insert</param>
        /// <param name="style">Style object to apply on this cell</param>
        /// <exception cref="RangeException">Throws a RangeException if the next cell is out of range (on row or column)</exception>
        /// <exception cref="StyleException">Throws a StyleException if the default style was malformed</exception>
        public void AddNextCell(object value, Style style)
        {
            AddNextCell(CastValue(value, currentColumnNumber, currentRowNumber), true, style);
        }


        /// <summary>
        /// Method to insert a generic cell to the next cell position
        /// </summary>
        /// <param name="cell">Cell object to insert</param>
        /// <param name="incremental">If true, the address value (row or column) will be incremented, otherwise not</param>
        /// <param name="style">If not null, the defined style will be applied to the cell, otherwise no style or the default style will be applied</param>
        /// <remarks>Recognized are the following data types: string, int, double, float, long, DateTime, TimeSpan, bool. 
        /// All other types will be casted into a string using the default ToString() method</remarks>
        /// <exception cref="StyleException">Throws a StyleException if the default style was malformed</exception>
        private void AddNextCell(Cell cell, bool incremental, Style style)
        {
            if (activeStyle != null && useActiveStyle && style == null)
            {
                cell.SetStyle(activeStyle);
            }
            else if (style != null)
            {
                cell.SetStyle(style);
            }
            else if (cell.DataType == Cell.CellType.DATE && cell.CellStyle == null)
            {
                cell.SetStyle(BasicStyles.DateFormat);
            }
            else
            {
                if (cell.DataType == Cell.CellType.TIME && cell.CellStyle == null)
                {
                    cell.SetStyle(BasicStyles.TimeFormat);
                }
            }
            string address = cell.CellAddress;
            if (cells.ContainsKey(address))
            {
                cells[address] = cell;
            }
            else
            {
                cells.Add(address, cell);
            }
            if (incremental)
            {
                if (CurrentCellDirection == CellDirection.ColumnToColumn)
                {
                    currentColumnNumber++;
                }
                else if (CurrentCellDirection == CellDirection.RowToRow)
                {
                    currentRowNumber++;
                }
                else
                {
                    // disabled / no-op
                }
            }
            else
            {
                if (CurrentCellDirection == CellDirection.ColumnToColumn)
                {
                    currentColumnNumber = cell.ColumnNumber + 1;
                    currentRowNumber = cell.RowNumber;
                }
                else if (CurrentCellDirection == CellDirection.RowToRow)
                {
                    currentColumnNumber = cell.ColumnNumber;
                    currentRowNumber = cell.RowNumber + 1;
                }
                else
                {
                    // disabled / no-op
                }
            }
        }

        /// <summary>
        /// Method to cast a value or align an object of the type Cell to the context of the worksheet
        /// </summary>
        /// <param name="value">Unspecified value or object of the type Cell</param>
        /// <param name="column">Column index</param>
        /// <param name="row">Row index</param>
        /// <returns>Cell object</returns>
        private Cell CastValue(object value, int column, int row)
        {
            Cell c;
            if (value != null && value.GetType() == typeof(Cell))
            {
                c = (Cell)value;
                c.CellAddress2 = new Address(column, row);
            }
            else
            {
                c = new Cell(value, Cell.CellType.DEFAULT, column, row);
            }
            return c;
        }


        #endregion

        #region methods_AddCell

        /// <summary>
        /// Adds an object to the defined cell address. If the type of the value does not match with one of the supported data types, it will be casted to a String. 
        /// A prepared object of the type Cell will not be casted but adjusted
        /// </summary>
        /// <param name="value">Unspecified value to insert</param>
        /// <param name="columnAddress">Column number (zero based)</param>
        /// <param name="rowAddress">Row number (zero based)</param>
        /// <remarks>Recognized are the following data types: Cell (prepared object), string, int, double, float, long, DateTime, TimeSpan, bool. 
        /// All other types will be casted into a string using the default ToString() method</remarks>
        /// <exception cref="RangeException">Throws an RangeException if the passed cell address is out of range</exception>
        public void AddCell(object value, int columnAddress, int rowAddress)
        {
            AddNextCell(CastValue(value, columnAddress, rowAddress), false, null);
        }

        /// <summary>
        /// Adds an object to the defined cell address. If the type of the value does not match with one of the supported data types, it will be casted to a String. 
        /// A prepared object of the type Cell will not be casted but adjusted
        /// </summary>
        /// <param name="value">Unspecified value to insert</param>
        /// <param name="columnAddress">Column number (zero based)</param>
        /// <param name="rowAddress">Row number (zero based)</param>
        /// <param name="style">Style to apply on the cell</param>
        /// <remarks>Recognized are the following data types: Cell (prepared object), string, int, double, float, long, DateTime, TimeSpan, bool. 
        /// All other types will be casted into a string using the default ToString() method</remarks>
        /// <exception cref="StyleException">Throws an UndefinedStyleException if the passed style is malformed</exception>
        /// <exception cref="RangeException">Throws an RangeException if the passed cell address is out of range</exception>
        public void AddCell(object value, int columnAddress, int rowAddress, Style style)
        {
            AddNextCell(CastValue(value, columnAddress, rowAddress), false, style);
        }


        /// <summary>
        /// Adds an object to the defined cell address. If the type of the value does not match with one of the supported data types, it will be casted to a String. 
        /// A prepared object of the type Cell will not be casted but adjusted
        /// </summary>
        /// <param name="value">Unspecified value to insert</param>
        /// <param name="address">Cell address in the format A1 - XFD1048576</param>
        /// <remarks>Recognized are the following data types: Cell (prepared object), string, int, double, float, long, DateTime, TimeSpan, bool. 
        /// All other types will be casted into a string using the default ToString() method</remarks>
        /// <exception cref="RangeException">Throws an RangeException if the passed cell address is out of range</exception>
        /// <exception cref="Exceptions.FormatException">Throws a FormatException if the passed cell address is malformed</exception>
        public void AddCell(object value, string address)
        {
            int column;
            int row;
            Cell.ResolveCellCoordinate(address, out column, out row);
            AddCell(value, column, row);
        }

        /// <summary>
        /// Adds an object to the defined cell address. If the type of the value does not match with one of the supported data types, it will be casted to a String. 
        /// A prepared object of the type Cell will not be casted but adjusted
        /// </summary>
        /// <param name="value">Unspecified value to insert</param>
        /// <param name="address">Cell address in the format A1 - XFD1048576</param>
        /// <param name="style">Style to apply on the cell</param>
        /// <remarks>Recognized are the following data types: Cell (prepared object), string, int, double, float, long, DateTime, TimeSpan, 
        /// bool. All other types will be casted into a string using the default ToString() method</remarks>
        /// <exception cref="StyleException">Throws an UndefinedStyleException if the passed style is malformed</exception>
        /// <exception cref="RangeException">Throws an RangeException if the passed cell address is out of range</exception>
        /// <exception cref="Exceptions.FormatException">Throws a FormatException if the passed cell address is malformed</exception>
        public void AddCell(object value, string address, Style style)
        {
            int column;
            int row;
            Cell.ResolveCellCoordinate(address, out column, out row);
            AddCell(value, column, row, style);
        }

        #endregion

        #region methods_AddCellFormula

        /// <summary>
        /// Adds a cell formula as string to the defined cell address
        /// </summary>
        /// <param name="formula">Formula to insert</param>
        /// <param name="address">Cell address in the format A1 - XFD1048576</param>
        /// <exception cref="RangeException">Throws an RangeException if the passed cell address is out of range</exception>
        /// <exception cref="Exceptions.FormatException">Throws a FormatException if the passed cell address is malformed</exception>
        public void AddCellFormula(string formula, string address)
        {
            int column;
            int row;
            Cell.ResolveCellCoordinate(address, out column, out row);
            Cell c = new Cell(formula, Cell.CellType.FORMULA, column, row);
            AddNextCell(c, false, null);
        }

        /// <summary>
        /// Adds a cell formula as string to the defined cell address
        /// </summary>
        /// <param name="formula">Formula to insert</param>
        /// <param name="address">Cell address in the format A1 - XFD1048576</param>
        /// <param name="style">Style to apply on the cell</param>
        /// <exception cref="StyleException">Throws an UndefinedStyleException if the passed style was malformed</exception>
        /// <exception cref="RangeException">Throws an RangeException if the passed cell address is out of range</exception>
        /// <exception cref="Exceptions.FormatException">Throws a FormatException if the passed cell address is malformed</exception>
        public void AddCellFormula(string formula, string address, Style style)
        {
            int column;
            int row;
            Cell.ResolveCellCoordinate(address, out column, out row);
            Cell c = new Cell(formula, Cell.CellType.FORMULA, column, row);
            AddNextCell(c, false, style);
        }

        /// <summary>
        /// Adds a cell formula as string to the defined cell address
        /// </summary>
        /// <param name="formula">Formula to insert</param>
        /// <param name="columnAddress">Column number (zero based)</param>
        /// <param name="rowAddress">Row number (zero based)</param>
        /// <exception cref="RangeException">Throws an RangeException if the passed cell address is out of range</exception>
        public void AddCellFormula(string formula, int columnAddress, int rowAddress)
        {
            Cell c = new Cell(formula, Cell.CellType.FORMULA, columnAddress, rowAddress);
            AddNextCell(c, false, null);
        }

        /// <summary>
        /// Adds a cell formula as string to the defined cell address
        /// </summary>
        /// <param name="formula">Formula to insert</param>
        /// <param name="columnAddress">Column number (zero based)</param>
        /// <param name="rowAddress">Row number (zero based)</param>
        /// <param name="style">Style to apply on the cell</param>
        /// <exception cref="RangeException">Throws an RangeException if the passed cell address is out of range</exception>
        public void AddCellFormula(string formula, int columnAddress, int rowAddress, Style style)
        {
            Cell c = new Cell(formula, Cell.CellType.FORMULA, columnAddress, rowAddress);
            AddNextCell(c, false, style);
        }

        /// <summary>
        /// Adds a formula as string to the next cell position
        /// </summary>
        /// <param name="formula">Formula to insert</param>
        /// <exception cref="RangeException">Trows a RangeException if the next cell is out of range (on row or column)</exception>
        public void AddNextCellFormula(string formula)
        {
            Cell c = new Cell(formula, Cell.CellType.FORMULA, currentColumnNumber, currentRowNumber);
            AddNextCell(c, true, null);
        }

        /// <summary>
        /// Adds a formula as string to the next cell position
        /// </summary>
        /// <param name="formula">Formula to insert</param>
        /// <param name="style">Style to apply on the cell</param>
        /// <exception cref="RangeException">Trows a RangeException if the next cell is out of range (on row or column)</exception>
        public void AddNextCellFormula(string formula, Style style)
        {
            Cell c = new Cell(formula, Cell.CellType.FORMULA, currentColumnNumber, currentRowNumber);
            AddNextCell(c, true, style);
        }

        #endregion

        #region methods_AddCellRange

        /// <summary>
        /// Adds a list of object values to a defined cell range. If the type of the a particular value does not match with one of the supported data types, it will be casted to a String. 
        /// Prepared objects of the type Cell will not be casted but adjusted
        /// </summary>
        /// <param name="values">List of unspecified objects to insert</param>
        /// <param name="startAddress">Start address</param>
        /// <param name="endAddress">End address</param>
        /// <remarks>The data types in the passed list can be mixed. Recognized are the following data types: string, int, double, float, long, DateTime, TimeSpan, bool. 
        /// All other types will be casted into a string using the default ToString() method</remarks>
        /// <exception cref="RangeException">Throws an RangeException if the number of cells resolved from the range differs from the number of passed values</exception>
        public void AddCellRange(IReadOnlyList<object> values, Address startAddress, Address endAddress)
        {
            AddCellRangeInternal(values, startAddress, endAddress, null);
        }

        /// <summary>
        /// Adds a list of object values to a defined cell range. If the type of the a particular value does not match with one of the supported data types, it will be casted to a String. 
        /// Prepared objects of the type Cell will not be casted but adjusted
        /// </summary>
        /// <param name="values">List of unspecified objects to insert</param>
        /// <param name="startAddress">Start address</param>
        /// <param name="endAddress">End address</param>
        /// <param name="style">Style to apply on the all cells of the range</param>
        /// <remarks>The data types in the passed list can be mixed. Recognized are the following data types: Cell (prepared object), string, int, double, float, long, DateTime, TimeSpan, bool. 
        /// All other types will be casted into a string using the default ToString() method</remarks>
        /// <exception cref="RangeException">Throws an RangeException if the number of cells resolved from the range differs from the number of passed values</exception>
        /// <exception cref="StyleException">Throws an UndefinedStyleException if the passed style is malformed</exception>
        public void AddCellRange(IReadOnlyList<object> values, Address startAddress, Address endAddress, Style style)
        {
            AddCellRangeInternal(values, startAddress, endAddress, style);
        }

        /// <summary>
        /// Adds a list of object values to a defined cell range. If the type of the a particular value does not match with one of the supported data types, it will be casted to a String. 
        /// Prepared objects of the type Cell will not be casted but adjusted
        /// </summary>
        /// <param name="values">List of unspecified objects to insert</param>
        /// <param name="cellRange">Cell range as string in the format like A1:D1 or X10:X22</param>
        /// <remarks>The data types in the passed list can be mixed. Recognized are the following data types: Cell (prepared object), string, int, double, float, long, DateTime, TimeSpan, bool. 
        /// All other types will be casted into a string using the default ToString() method</remarks>
        /// <exception cref="RangeException">Throws an RangeException if the number of cells resolved from the range differs from the number of passed values</exception>
        /// <exception cref="Exceptions.FormatException">Throws a FormatException if the passed cell range is malformed</exception>
        public void AddCellRange(IReadOnlyList<object> values, string cellRange)
        {
            Range range = Cell.ResolveCellRange(cellRange);
            AddCellRangeInternal(values, range.StartAddress, range.EndAddress, null);
        }

        /// <summary>
        /// Adds a list of object values to a defined cell range. If the type of the a particular value does not match with one of the supported data types, it will be casted to a String. 
        /// Prepared objects of the type Cell will not be casted but adjusted
        /// </summary>
        /// <param name="values">List of unspecified objects to insert</param>
        /// <param name="cellRange">Cell range as string in the format like A1:D1 or X10:X22</param>
        /// <param name="style">Style to apply on the all cells of the range</param>
        /// <remarks>The data types in the passed list can be mixed. Recognized are the following data types: Cell (prepared object), string, int, double, float, long, DateTime, TimeSpan, bool. 
        /// All other types will be casted into a string using the default ToString() method</remarks>
        /// <exception cref="RangeException">Throws an RangeException if the number of cells resolved from the range differs from the number of passed values</exception>
        /// <exception cref="StyleException">Throws an UndefinedStyleException if the passed style is malformed</exception>
        /// <exception cref="Exceptions.FormatException">Throws a FormatException if the passed cell range is malformed</exception>
        public void AddCellRange(IReadOnlyList<object> values, string cellRange, Style style)
        {
            Range range = Cell.ResolveCellRange(cellRange);
            AddCellRangeInternal(values, range.StartAddress, range.EndAddress, style);
        }

        /// <summary>
        /// Internal function to add a generic list of value to the defined cell range
        /// </summary>
        /// <typeparam name="T">Data type of the generic value list</typeparam>
        /// <param name="values">List of values</param>
        /// <param name="startAddress">Start address</param>
        /// <param name="endAddress">End address</param>
        /// <param name="style">Style to apply on the all cells of the range</param>
        /// <remarks>The data types in the passed list can be mixed. Recognized are the following data types: Cell (prepared object), string, int, double, float, long, DateTime, TimeSpan, bool. 
        /// All other types will be casted into a string using the default ToString() method</remarks>
        /// <exception cref="RangeException">Throws an RangeException if the number of cells differs from the number of passed values</exception>
        private void AddCellRangeInternal<T>(IReadOnlyList<T> values, Address startAddress, Address endAddress, Style style)
        {
            List<Address> addresses = Cell.GetCellRange(startAddress, endAddress) as List<Address>;
            if (values.Count != addresses.Count)
            {
                throw new RangeException(RangeException.GENERAL, "The number of passed values (" + values.Count + ") differs from the number of cells within the range (" + addresses.Count + ")");
            }
            List<Cell> list = Cell.ConvertArray(values) as List<Cell>;
            int len = values.Count;
            for (int i = 0; i < len; i++)
            {
                list[i].RowNumber = addresses[i].Row;
                list[i].ColumnNumber = addresses[i].Column;
                AddNextCell(list[i], false, style);
            }
        }
        #endregion

        #region methods_RemoveCell
        /// <summary>
        /// Removes a previous inserted cell at the defined address
        /// </summary>
        /// <param name="columnAddress">Column number (zero based)</param>
        /// <param name="rowAddress">Row number (zero based)</param>
        /// <returns>Returns true if the cell could be removed (existed), otherwise false (did not exist)</returns>
        /// <exception cref="RangeException">Throws an RangeException if the passed cell address is out of range</exception>
        public bool RemoveCell(int columnAddress, int rowAddress)
        {
            string address = Cell.ResolveCellAddress(columnAddress, rowAddress);
            return cells.Remove(address);
        }

        /// <summary>
        /// Removes a previous inserted cell at the defined address
        /// </summary>
        /// <param name="address">Cell address in the format A1 - XFD1048576</param>
        /// <returns>Returns true if the cell could be removed (existed), otherwise false (did not exist)</returns>
        /// <exception cref="RangeException">Throws an RangeException if the passed cell address is out of range</exception>
        /// <exception cref="Exceptions.FormatException">Throws a FormatException if the passed cell address is malformed</exception>
        public bool RemoveCell(string address)
        {
            int row;
            int column;
            Cell.ResolveCellCoordinate(address, out column, out row);
            return RemoveCell(column, row);
        }
        #endregion

        #region methods_setStyle

        /// <summary>
        /// Sets the passed style on the passed cell range. If cells are already existing, the style will be added or replaced. Otherwise, an empty (numeric) cell will be added with the assigned style
        /// </summary>
        /// <param name="cellRange">Cell range to apply the style</param>
        /// <param name="style">Style to apply</param>
        /// <remarks>Note: This method may invalidate an existing date value since dates are defined by specific style. The result of a redefinition will be a number, instead of a date</remarks>
        public void SetStyle(Range cellRange, Style style)
        {
            IReadOnlyList<Address> addresses = cellRange.ResolveEnclosedAddresses();
            foreach (Address address in addresses)
            {
                String key = address.GetAddress();
                if (this.cells.ContainsKey(key))
                {
                    cells[key].SetStyle(style);
                }
                else
                {
                    AddCell(null, address.Column, address.Row, style);
                }
            }
        }

        /// <summary>
        /// Sets the passed style on the passed cell range, derived from a start and end address. If cells are already existing, the style will be added or replaced. 
        /// Otherwise, an empty (numeric) cell will be added with the assigned style
        /// </summary>
        /// <param name="startAddress">Start address of the cell range</param>
        /// <param name="endAddress">End address of the cell range</param>
        /// <param name="style">Style to apply</param>
        /// <remarks>Note: This method may invalidate an existing date value since dates are defined by specific style. The result of a redefinition will be a number, instead of a date</remarks>
        public void SetStyle(Address startAddress, Address endAddress, Style style)
        {
            SetStyle(new Range(startAddress, endAddress), style);
        }

        /// <summary>
        /// Sets the passed style on the passed (singular) cell address. If the cell is already existing, the style will be added or replaced.
        /// Otherwise, an empty (numeric) cell will be added with the assigned style
        /// </summary>
        /// <param name="address">Cell address to apply the style</param>
        /// <param name="style">Style to apply</param>
        /// <remarks>Note: This method may invalidate an existing date value since dates are defined by specific style. The result of a redefinition will be a number, instead of a date</remarks>
        public void SetStyle(Address address, Style style)
        {
            SetStyle(address, address, style);
        }

        /// <summary>
        /// Sets the passed style on the passed address expression. Such an expression may be a single cell or a cell range.
        /// If the cell is already existing, the style will be added or replaced. Otherwise, an empty (numeric) cell or cell range will be added with the assigned style
        /// </summary>
        /// <param name="addressExpression">Expression of a cell address or range of addresses</param>
        /// <param name="style">Style to apply</param>
        /// <remarks>Note: This method may invalidate an existing date value since dates are defined by specific style. The result of a redefinition will be a number, instead of a date</remarks>
        public void SetStyle(string addressExpression, Style style)
        {
            Cell.AddressScope scope = Cell.GetAddressScope(addressExpression);
            if (scope == Cell.AddressScope.SingleAddress)
            {
                Address address = new Address(addressExpression);
                SetStyle(address, style);
            }
            else if (scope == Cell.AddressScope.Range)
            {
                Range range = new Range(addressExpression);
                SetStyle(range, style);
            }
            else
            {
                throw new FormatException("InvalidAddressExpression", "The passed address'" + addressExpression + "' is neither a cell address, nor a range");
            }
        }

        #endregion

        #region common_methods

        /// <summary>
        /// Method to add allowed actions if the worksheet is protected. If one or more values are added, UseSheetProtection will be set to true
        /// </summary>
        /// <param name="typeOfProtection">Allowed action on the worksheet or cells</param>
        public void AddAllowedActionOnSheetProtection(SheetProtectionValue typeOfProtection)
        {
            if (!sheetProtectionValues.Contains(typeOfProtection))
            {
                if (typeOfProtection == SheetProtectionValue.selectLockedCells && !sheetProtectionValues.Contains(SheetProtectionValue.selectUnlockedCells))
                {
                    sheetProtectionValues.Add(SheetProtectionValue.selectUnlockedCells);
                }
                sheetProtectionValues.Add(typeOfProtection);
                UseSheetProtection = true;
            }
        }

        /// <summary>
        /// Sets the defined column as hidden
        /// </summary>
        /// <param name="columnNumber">Column number to hide on the worksheet</param>
        /// <exception cref="RangeException">Throws a RangeException if the passed column number is out of range</exception>
        public void AddHiddenColumn(int columnNumber)
        {
            SetColumnHiddenState(columnNumber, true);
        }

        /// <summary>
        /// Sets the defined column as hidden
        /// </summary>
        /// <param name="columnAddress">Column address to hide on the worksheet</param>
        /// <exception cref="RangeException">Throws an RangeException if the passed column address is out of range</exception>
        public void AddHiddenColumn(string columnAddress)
        {
            int columnNumber = Cell.ResolveColumn(columnAddress);
            SetColumnHiddenState(columnNumber, true);
        }

        /// <summary>
        /// Sets the defined row as hidden
        /// </summary>
        /// <param name="rowNumber">Row number to hide on the worksheet</param>
        /// <exception cref="RangeException">Throws an RangeException if the passed row number is out of range</exception>
        public void AddHiddenRow(int rowNumber)
        {
            SetRowHiddenState(rowNumber, true);
        }

        /// <summary>
        /// Clears the active style of the worksheet. All later added calls will contain no style unless another active style is set
        /// </summary>
        public void ClearActiveStyle()
        {
            useActiveStyle = false;
            activeStyle = null;
        }

        /// <summary>
        /// Gets the cell of the specified address
        /// </summary>
        /// <param name="address">Address of the cell</param>
        /// <returns>Cell object</returns>
        /// <exception cref="WorksheetException">Trows a WorksheetException if the cell was not found on the cell table of this worksheet</exception>
        public Cell GetCell(Address address)
        {
            if (!cells.ContainsKey(address.GetAddress()))
            {
                throw new WorksheetException("CellNotFoundException", "The cell with the address " + address.GetAddress() + " does not exist in this worksheet");
            }
            return cells[address.GetAddress()];
        }

        /// <summary>
        /// Gets the cell of the specified column and row number (zero-based)
        /// </summary>
        /// <param name="columnNumber">Column number of the cell</param>
        /// <param name="rowNumber">Row number of the cell</param>
        /// <returns>Cell object</returns>
        /// <exception cref="WorksheetException">Trows a WorksheetException if the cell was not found on the cell table of this worksheet</exception>
        public Cell GetCell(int columnNumber, int rowNumber)
        {
            return GetCell(new Address(columnNumber, rowNumber));
        }

        /// <summary>
        /// Gets whether the specified address exists in the worksheet. Existing means that a value was stored at the address
        /// </summary>
        /// <param name="address">Address to check</param>
        /// <returns>
        ///   <c>true</c> if the cell exists, otherwise <c>false</c>.
        /// </returns>
        public bool HasCell(Address address)
        {
            return cells.ContainsKey(address.GetAddress());
        }

        /// <summary>
        /// Gets whether the specified address exists in the worksheet. Existing means that a value was stored at the address
        /// </summary>
        /// <param name="columnNumber">Column number of the cell to check (zero-based)</param>
        /// <param name="rowNumber">Row number of the cell to check (zero-based)</param>
        /// <returns>
        ///   <c>true</c> if the cell exists, otherwise <c>false</c>.
        /// </returns>
        public bool HasCell(int columnNumber, int rowNumber)
        {
            return HasCell(new Address(columnNumber, rowNumber));
        }

        /// <summary>
        /// Resets the defined column, if existing. The corresponding instance will be removed from <see cref="Columns"/>.
        /// </summary>
        /// <remarks>If the column is inside an autoFilter-Range, the column cannot be entirely removed from <see cref="Columns"/>. The hidden state will be set to false and width to default, in this case.</remarks>
        /// <param name="columnNumber">Column number to reset (zero-based)</param>
        public void ResetColumn(int columnNumber)
        {
            if (columns.ContainsKey(columnNumber) && !columns[columnNumber].HasAutoFilter) // AutoFilters cannot have gaps 
            {
                columns.Remove(columnNumber);
            }
            else
            {
                columns[columnNumber].IsHidden = false;
                columns[columnNumber].Width = DEFAULT_COLUMN_WIDTH;
            }
        }

        /// <summary>
        /// Gets the last existing column number in the current worksheet (zero-based)
        /// </summary>
        /// <returns>Zero-based column number. In case of a empty worksheet, -1 will be returned</returns>
        /// <remarks>GetLastColumnNumber() will not return the last column with data in any case. If there is a formated but empty cell (or many) beyond the last cell with data, 
        /// GetLastColumnNumber() will return the column number of this empty cell. Use <see cref="GetLastDataColumnNumber"/> in this case.</remarks>
        public int GetLastColumnNumber()
        {
            return GetLastAddress(true, false);
        }

        /// <summary>
        /// Gets the last existing column number with data in the current worksheet (zero-based)
        /// </summary>
        /// <returns>Zero-based column number. In case of a empty worksheet, -1 will be returned</returns>
        /// <remarks>GetLastDataColumnNumber() will ignore formatted but empty cells beyond the last column with data. 
        /// If you want the last defined column, use <see cref="GetLastColumnNumber"/> instead.</remarks>
        public int GetLastDataColumnNumber()
        {
            return GetLastAddress(true, true);
        }

        /// <summary>
        /// Gets the last existing row number in the current worksheet (zero-based)
        /// </summary>
        /// <returns>Zero-based row number. In case of a empty worksheet, -1 will be returned</returns>
        /// <remarks>GetLastRowNumber() will not return the last row with data in any case. If there is a formated but empty cell (or many) beyond the last cell with data, 
        /// GetLastRowNumber() will return the row number of this empty cell. Use <see cref="GetLastDataRowNumber"/> in this case.</remarks>
        public int GetLastRowNumber()
        {
            return GetLastAddress(false, false);
        }


        /// <summary>
        /// Gets the last existing row number with data in the current worksheet (zero-based)
        /// </summary>
        /// <returns>Zero-based row number. In case of a empty worksheet, -1 will be returned</returns>
        /// <remarks>GetLastDataColumnNumber() will ignore formatted but empty cells beyond the last column with data. 
        /// If you want the last defined column, use <see cref="GetLastColumnNumber"/> instead.</remarks>
        public int GetLastDataRowNumber()
        {
            return GetLastAddress(false, true);
        }

        /// <summary>
        ///  Gets the last existing cell in the current worksheet (bottom right)
        /// </summary>
        /// <returns>Cell Address</returns>
        /// <remarks>GetLastCellAddress() will not return the last cell with data in any case. If there is a formated but empty cell (or many) beyond the last cell with data, 
        /// GetLastCellAddress() will return the address of this empty cell. Use <see cref="GetLastDataCellAddress"/> in this case.</remarks>

        public Address GetLastCellAddress()
        {
            int lastRow = GetLastRowNumber();
            int lastColumn = GetLastColumnNumber();
            return new Address(lastColumn, lastRow);
        }

        /// <summary>
        ///  Gets the last existing cell with data in the current worksheet (bottom right)
        /// </summary>
        /// <returns>Cell Address</returns>
        /// <remarks>GetLastDataCellAddress() will ignore formatted but empty cells beyond the last cell with data. 
        /// If you want the last defined cell, use <see cref="GetLastCellAddress"/> instead.</remarks>

        public Address GetLastDataCellAddress()
        {
            int lastRow = GetLastDataRowNumber();
            int lastColumn = GetLastDataColumnNumber();
            return new Address(lastColumn, lastRow);
        }

        /// <summary>
        /// Gets the last existing row or column number of the current worksheet (zero-based)
        /// </summary>
        /// <param name="column">If true, the output will be the last column, otherwise the last row</param>
        /// <param name="ignoreEmpty">If true, empty cells are ignored and the last column or row is this one with a value</param>
        /// <returns>Last row or column number (zero-based)</returns>
        private int GetLastAddress(bool column, bool ignoreEmpty)
        {
            int max = -1;
            int number;
            foreach (KeyValuePair<string, Cell> cell in cells)
            {
                number = column ? cell.Value.ColumnNumber : cell.Value.RowNumber;
                if (ignoreEmpty && cell.Value.Value != null && cell.Value.Value.ToString() != String.Empty && number > max)
                {
                    max = number;
                }
                else if (!ignoreEmpty && number > max)
                {
                    max = number;
                }
            }
            return max;
        }

        /// <summary>
        /// Gets the current column number (zero based)
        /// </summary>
        /// <returns>Column number (zero-based)</returns>
        public int GetCurrentColumnNumber()
        {
            return currentColumnNumber;
        }

        /// <summary>
        /// Gets the current row number (zero based)
        /// </summary>
        /// <returns>Row number (zero-based)</returns>
        public int GetCurrentRowNumber()
        {
            return currentRowNumber;
        }

        /// <summary>
        /// Moves the current position to the next column
        /// </summary>
        public void GoToNextColumn()
        {
            currentColumnNumber++;
            currentRowNumber = 0;
        }

        /// <summary>
        /// Moves the current position to the next column with the number of cells to move
        /// </summary>
        /// <param name="numberOfColumns">Number of columns to move</param>
        public void GoToNextColumn(int numberOfColumns)
        {
            for (int i = 0; i < numberOfColumns; i++)
            {
                GoToNextColumn();
            }
        }

        /// <summary>
        /// Moves the current position to the next row (use for a new line)
        /// </summary>
        public void GoToNextRow()
        {
            currentRowNumber++;
            currentColumnNumber = 0;
        }

        /// <summary>
        /// Moves the current position to the next row with the number of cells to move (use for a new line)
        /// </summary>
        /// <param name="numberOfRows">Number of rows to move</param>
        public void GoToNextRow(int numberOfRows)
        {
            for (int i = 0; i < numberOfRows; i++)
            {
                GoToNextRow();
            }
        }

        /// <summary>
        /// Merges the defined cell range
        /// </summary>
        /// <param name="cellRange">Range to merge</param>
        /// <returns>Returns the validated range of the merged cells (e.g. 'A1:B12')</returns>
        /// <exception cref="RangeException">Throws an RangeException if the passed cell range is out of range</exception>
        public string MergeCells(Range cellRange)
        {
            return MergeCells(cellRange.StartAddress, cellRange.EndAddress);
        }

        /// <summary>
        /// Merges the defined cell range
        /// </summary>
        /// <param name="cellRange">Range to merge (e.g. 'A1:B12')</param>
        /// <returns>Returns the validated range of the merged cells (e.g. 'A1:B12')</returns>
        /// <exception cref="RangeException">Throws an RangeException if the passed cell range is out of range</exception>
        /// <exception cref="Exceptions.FormatException">Throws a FormatException if the passed cell range is malformed</exception>
        public string MergeCells(string cellRange)
        {
            Range range = Cell.ResolveCellRange(cellRange);
            return MergeCells(range.StartAddress, range.EndAddress);
        }

        /// <summary>
        /// Merges the defined cell range
        /// </summary>
        /// <param name="startAddress">Start address of the merged cell range</param>
        /// <param name="endAddress">End address of the merged cell range</param>
        /// <returns>Returns the validated range of the merged cells (e.g. 'A1:B12')</returns>
        /// <exception cref="RangeException">Throws an RangeException if one of the passed cell addresses is out of range</exception>
        public string MergeCells(Address startAddress, Address endAddress)
        {
            string key = startAddress + ":" + endAddress;
            Range value = new Range(startAddress, endAddress);
            if (!mergedCells.ContainsKey(key))
            {
                mergedCells.Add(key, value);
            }
            return key;
        }

        /// <summary>
        /// Method to recalculate the auto filter (columns) of this worksheet. This is an internal method. There is no need to use it. It must be public to require access from the LowLevel class
        /// </summary>
        public void RecalculateAutoFilter()
        {
            if (autoFilterRange == null)
            { return; }
            int start = autoFilterRange.Value.StartAddress.Column;
            int end = autoFilterRange.Value.EndAddress.Column;
            int endRow = 0;
            foreach (KeyValuePair<string, Cell> item in Cells)
            {
                if (item.Value.ColumnNumber < start || item.Value.ColumnNumber > end)
                { continue; }
                if (item.Value.RowNumber > endRow)
                { endRow = item.Value.RowNumber; }
            }
            Column c;
            for (int i = start; i <= end; i++)
            {
                if (!columns.ContainsKey(i))
                {
                    c = new Column(i);
                    c.HasAutoFilter = true;
                    columns.Add(i, c);
                }
                else
                {
                    columns[i].HasAutoFilter = true;
                }
            }
            Range temp = new Range();
            temp.StartAddress = new Address(start, 0);
            temp.EndAddress = new Address(end, endRow);
            autoFilterRange = temp;
        }

        /// <summary>
        /// Method to recalculate the collection of columns of this worksheet. This is an internal method. There is no need to use it. It must be public to require access from the LowLevel class
        /// </summary>
        public void RecalculateColumns()
        {
            List<int> columnsToDelete = new List<int>();
            foreach (KeyValuePair<int, Column> col in columns)
            {
                if (!col.Value.HasAutoFilter && !col.Value.IsHidden && Math.Abs(col.Value.Width - DEFAULT_COLUMN_WIDTH) <= FLOAT_TRESHOLD)
                {
                    columnsToDelete.Add(col.Key);
                }

                if (!col.Value.HasAutoFilter && !col.Value.IsHidden && Math.Abs(col.Value.Width - DEFAULT_COLUMN_WIDTH) <= FLOAT_TRESHOLD)
                {
                    columnsToDelete.Add(col.Key);
                }
            }
            foreach (int index in columnsToDelete)
            {
                columns.Remove(index);
            }
        }

        /// <summary>
        /// Removes auto filters from the worksheet
        /// </summary>
        public void RemoveAutoFilter()
        {
            autoFilterRange = null;
        }

        /// <summary>
        /// Sets a previously defined, hidden column as visible again
        /// </summary>
        /// <param name="columnNumber">Column number to make visible again</param>
        /// <exception cref="RangeException">Throws an RangeException if the passed column number is out of range</exception>
        public void RemoveHiddenColumn(int columnNumber)
        {
            SetColumnHiddenState(columnNumber, false);
        }

        /// <summary>
        /// Sets a previously defined, hidden column as visible again
        /// </summary>
        /// <param name="columnAddress">Column address to make visible again</param>
        /// <exception cref="RangeException">Throws an RangeException if the column address out of range</exception>
        public void RemoveHiddenColumn(string columnAddress)
        {
            int columnNumber = Cell.ResolveColumn(columnAddress);
            SetColumnHiddenState(columnNumber, false);
        }

        /// <summary>
        /// Sets a previously defined, hidden row as visible again
        /// </summary>
        /// <param name="rowNumber">Row number to hide on the worksheet</param>
        /// <exception cref="RangeException">Throws an RangeException if the passed row number is out of range</exception>
        public void RemoveHiddenRow(int rowNumber)
        {
            SetRowHiddenState(rowNumber, false);
        }

        /// <summary>
        /// Removes the defined merged cell range
        /// </summary>
        /// <param name="range">Cell range to remove the merging</param>
        /// <exception cref="RangeException">Throws a UnkownRangeException if the passed cell range was not merged earlier</exception>
        public void RemoveMergedCells(string range)
        {
            range = Utils.ToUpper(range);
            if (!mergedCells.ContainsKey(range))
            {
                throw new RangeException("UnknownRangeException", "The cell range " + range + " was not found in the list of merged cell ranges");
            }

            List<Address> addresses = Cell.GetCellRange(range) as List<Address>;
            Cell cell;
            foreach (Address address in addresses)
            {
                if (cells.ContainsKey(address.ToString()))
                {
                    cell = cells[address.ToString()];
                    cell.DataType = Cell.CellType.DEFAULT; // resets the type
                    if (cell.Value == null)
                    {
                        cell.Value = string.Empty;
                    }
                }
            }
            mergedCells.Remove(range);
        }

        /// <summary>
        /// Removes the cell selection of this worksheet
        /// </summary>
        public void RemoveSelectedCells()
        {
            selectedCells = null;
        }

        /// <summary>
        /// Removes the defined, non-standard row height
        /// </summary>
        /// <param name="rowNumber">Row number (zero-based)</param>
        public void RemoveRowHeight(int rowNumber)
        {
            if (rowHeights.ContainsKey(rowNumber))
            {
                rowHeights.Remove(rowNumber);
            }
        }

        /// <summary>
        /// Removes an allowed action on the current worksheet or its cells
        /// </summary>
        /// <param name="value">Allowed action on the worksheet or cells</param>
        public void RemoveAllowedActionOnSheetProtection(SheetProtectionValue value)
        {
            if (sheetProtectionValues.Contains(value)){
                sheetProtectionValues.Remove(value);
            }
        }

        /// <summary>
        /// Sets the active style of the worksheet. This style will be assigned to all later added cells
        /// </summary>
        /// <param name="style">Style to set as active style</param>
        public void SetActiveStyle(Style style)
        {
            useActiveStyle = true;
            activeStyle = style;
        }

        /// <summary>
        /// Sets the column auto filter within the defined column range
        /// </summary>
        /// <param name="startColumn">Column number with the first appearance of an auto filter drop down</param>
        /// <param name="endColumn">Column number with the last appearance of an auto filter drop down</param>
        /// <exception cref="RangeException">Throws an RangeException if the start or end address out of range</exception>
        public void SetAutoFilter(int startColumn, int endColumn)
        {
            string start = Cell.ResolveCellAddress(startColumn, 0);
            string end = Cell.ResolveCellAddress(endColumn, 0);
            if (endColumn < startColumn)
            {
                SetAutoFilter(end + ":" + start);
            }
            else
            {
                SetAutoFilter(start + ":" + end);
            }
        }

        /// <summary>
        /// Sets the column auto filter within the defined column range
        /// </summary>
        /// <param name="range">Range to apply auto filter on. The range could be 'A1:C10' for instance. The end row will be recalculated automatically when saving the file</param>
        /// <exception cref="RangeException">Throws an RangeException if the passed range out of range</exception>
        /// <exception cref="Exceptions.FormatException">Throws an FormatException if the passed range is malformed</exception>
        public void SetAutoFilter(string range)
        {
            autoFilterRange = Cell.ResolveCellRange(range);
            RecalculateAutoFilter();
            RecalculateColumns();
        }

        /// <summary>
        /// Sets the defined column as hidden or visible
        /// </summary>
        /// <param name="columnNumber">Column number to hide on the worksheet</param>
        /// <param name="state">If true, the column will be hidden, otherwise be visible</param>
        /// <exception cref="RangeException">Throws an RangeException if the column number out of range</exception>
        private void SetColumnHiddenState(int columnNumber, bool state)
        {
            if (columnNumber > MAX_COLUMN_NUMBER || columnNumber < MIN_COLUMN_NUMBER)
            {
                throw new RangeException(RangeException.GENERAL, "The column number (" + columnNumber + ") is out of range. Range is from " +
                    MIN_COLUMN_NUMBER + " to " + MAX_COLUMN_NUMBER + " (" + (MAX_COLUMN_NUMBER + 1) + " columns).");
            }
            if (columns.ContainsKey(columnNumber) && state)
            {
                columns[columnNumber].IsHidden = true;
            }
            else if (state)
            {
                Column c = new Column(columnNumber);
                c.IsHidden = true;
                columns.Add(columnNumber, c);
            }
            else
            {
                // no-op
            }
        }

        /// <summary>
        /// Sets the width of the passed column address
        /// </summary>
        /// <param name="columnAddress">Column address (A - XFD)</param>
        /// <param name="width">Width from 0 to 255.0</param>
        /// <exception cref="RangeException">Throws an RangeException:<br></br>a) If the passed column address is out of range<br></br>b) if the column width is out of range (0 - 255.0)</exception>
        public void SetColumnWidth(string columnAddress, float width)
        {
            int columnNumber = Cell.ResolveColumn(columnAddress);
            SetColumnWidth(columnNumber, width);
        }

        /// <summary>
        /// Sets the width of the passed column number (zero-based)
        /// </summary>
        /// <param name="columnNumber">Column number (zero-based, from 0 to 16383)</param>
        /// <param name="width">Width from 0 to 255.0</param>
        /// <exception cref="RangeException">Throws an RangeException:<br></br>a) If the passed column number is out of range<br></br>b) if the column width is out of range (0 - 255.0)</exception>
        public void SetColumnWidth(int columnNumber, float width)
        {
            if (columnNumber > MAX_COLUMN_NUMBER || columnNumber < MIN_COLUMN_NUMBER)
            {
                throw new RangeException(RangeException.GENERAL, "The column number (" + columnNumber + ") is out of range. Range is from " +
                    MIN_COLUMN_NUMBER + " to " + MAX_COLUMN_NUMBER + " (" + (MAX_COLUMN_NUMBER + 1) + " columns).");
            }
            if (width < MIN_COLUMN_WIDTH || width > MAX_COLUMN_WIDTH)
            {
                throw new RangeException(RangeException.GENERAL, "The column width (" + width + ") is out of range. Range is from " + MIN_COLUMN_WIDTH + " to " + MAX_COLUMN_WIDTH + " (chars).");
            }
            if (columns.ContainsKey(columnNumber))
            {
                columns[columnNumber].Width = width;
            }
            else
            {
                Column c = new Column(columnNumber);
                c.Width = width;
                columns.Add(columnNumber, c);
            }
        }

        /// <summary>
        /// Set the current cell address
        /// </summary>
        /// <param name="columnAddress">Column number (zero based)</param>
        /// <param name="rowAddress">Row number (zero based)</param>
        /// <exception cref="RangeException">Throws an RangeException if one of the passed cell addresses is out of range</exception>
        public void SetCurrentCellAddress(int columnAddress, int rowAddress)
        {
            SetCurrentColumnNumber(columnAddress);
            SetCurrentRowNumber(rowAddress);
        }

        /// <summary>
        /// Set the current cell address
        /// </summary>
        /// <param name="address">Cell address in the format A1 - XFD1048576</param>
        /// <exception cref="RangeException">Throws an RangeException if the passed cell address is out of range</exception>
        /// <exception cref="Exceptions.FormatException">Throws a FormatException if the passed cell address is malformed</exception>
        public void SetCurrentCellAddress(string address)
        {
            int row;
            int column;
            Cell.ResolveCellCoordinate(address, out column, out row);
            SetCurrentCellAddress(column, row);
        }

        /// <summary>
        /// Sets the current column number (zero based)
        /// </summary>
        /// <param name="columnNumber">Column number (zero based)</param>
        /// <exception cref="RangeException">Throws an RangeException if the number is out of the valid range. Range is from 0 to 16383 (16384 columns)</exception>
        public void SetCurrentColumnNumber(int columnNumber)
        {
            if (columnNumber > MAX_COLUMN_NUMBER || columnNumber < MIN_COLUMN_NUMBER)
            {
                throw new RangeException(RangeException.GENERAL, "The column number (" + columnNumber + ") is out of range. Range is from " +
                    MIN_COLUMN_NUMBER + " to " + MAX_COLUMN_NUMBER + " (" + (MAX_COLUMN_NUMBER + 1) + " columns).");
            }
            currentColumnNumber = columnNumber;
        }

        /// <summary>
        /// Sets the current row number (zero based)
        /// </summary>
        /// <param name="rowNumber">Row number (zero based)</param>
        /// <exception cref="RangeException">Throws an RangeException if the number is out of the valid range. Range is from 0 to 1048575 (1048576 rows)</exception>
        public void SetCurrentRowNumber(int rowNumber)
        {
            if (rowNumber > MAX_ROW_NUMBER || rowNumber < 0)
            {
                throw new RangeException(RangeException.GENERAL, "The row number (" + rowNumber + ") is out of range. Range is from 0 to " + MAX_ROW_NUMBER + " (" + (MAX_ROW_NUMBER + 1) + " rows).");
            }
            currentRowNumber = rowNumber;
        }

        /// <summary>
        /// Sets the selected cells on this worksheet
        /// </summary>
        /// <param name="range">Cell range to select</param>
        public void SetSelectedCells(Range range)
        {
            selectedCells = range;
        }

        /// <summary>
        /// Sets the selected cells on this worksheet
        /// </summary>
        /// <param name="startAddress">Start address of the range</param>
        /// <param name="endAddress">End address of the range</param>
        public void SetSelectedCells(Address startAddress, Address endAddress)
        {
            selectedCells = new Range(startAddress, endAddress);
        }

        /// <summary>
        /// Sets the selected cells on this worksheet
        /// </summary>
        /// <param name="range">Cell range to select</param>
        public void SetSelectedCells(string range)
        {
            selectedCells = Cell.ResolveCellRange(range);
        }

        /// <summary>
        /// Sets or removes the password for worksheet protection. If set, UseSheetProtection will be also set to true
        /// </summary>
        /// <param name="password">Password (UTF-8) to protect the worksheet. If the password is null or empty, no password will be used</param>
        public void SetSheetProtectionPassword(string password)
        {
            if (string.IsNullOrEmpty(password))
            {
                sheetProtectionPassword = null;
            }
            else
            {
                sheetProtectionPassword = password;
                UseSheetProtection = true;
            }
        }

        /// <summary>
        /// Sets the height of the passed row number (zero-based)
        /// </summary>
        /// <param name="rowNumber">Row number (zero-based, 0 to 1048575)</param>
        /// <param name="height">Height from 0 to 409.5</param>
        /// <exception cref="RangeException">Throws an RangeException:<br></br>a) If the passed row number is out of range<br></br>b) if the row height is out of range (0 - 409.5)</exception>
        public void SetRowHeight(int rowNumber, float height)
        {
            if (rowNumber > MAX_ROW_NUMBER || rowNumber < MIN_ROW_NUMBER)
            {
                throw new RangeException(RangeException.GENERAL, "The row number (" + rowNumber + ") is out of range. Range is from " +
                    MIN_ROW_NUMBER + " to " + MAX_ROW_NUMBER + " (" + (MAX_ROW_NUMBER + 1) + " rows).");
            }
            if (height < MIN_ROW_HEIGHT || height > MAX_ROW_HEIGHT)
            {
                throw new RangeException(RangeException.GENERAL, "The row height (" + height + ") is out of range. Range is from " + MIN_ROW_HEIGHT + " to " + MAX_ROW_HEIGHT + " (equals 546px).");
            }
            if (rowHeights.ContainsKey(rowNumber))
            {
                rowHeights[rowNumber] = height;
            }
            else
            {
                rowHeights.Add(rowNumber, height);
            }
        }

        /// <summary>
        /// Sets the defined row as hidden or visible
        /// </summary>
        /// <param name="rowNumber">Row number to make visible again</param>
        /// <param name="state">If true, the row will be hidden, otherwise visible</param>
        /// <exception cref="RangeException">Throws an RangeException if the passed row number was out of range</exception>
        private void SetRowHiddenState(int rowNumber, bool state)
        {
            if (rowNumber > MAX_ROW_NUMBER || rowNumber < MIN_ROW_NUMBER)
            {
                throw new RangeException(RangeException.GENERAL, "The row number (" + rowNumber + ") is out of range. Range is from " +
                    MIN_ROW_NUMBER + " to " + MAX_ROW_NUMBER + " (" + (MAX_ROW_NUMBER + 1) + " rows).");
            }
            if (hiddenRows.ContainsKey(rowNumber))
            {
                if (state)
                {
                    hiddenRows.Add(rowNumber, true);
                }
                else
                {
                    hiddenRows.Remove(rowNumber);
                }
            }
            else if (state)
            {
                hiddenRows.Add(rowNumber, true);
            }
            else
            {
                // no-op
            }
        }

        /// <summary>
        /// Validates and sets the worksheet name
        /// </summary>
        /// <param name="name">Name to set</param>
        /// <exception cref="Exceptions.FormatException">Throws a FormatException if the worksheet name is too long (max. 31) or contains illegal characters [  ]  * ? / \</exception>
        public void SetSheetname(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                throw new FormatException("the worksheet name must be between 1 and " + MAX_WORKSHEER_NAME_LENGTH + " characters");
            }
            if (name.Length > MAX_WORKSHEER_NAME_LENGTH)
            {
                throw new FormatException("the worksheet name must be between 1 and " + MAX_WORKSHEER_NAME_LENGTH + " characters");
            }
            Regex rx = new Regex(@"[\[\]\*\?/\\]");
            Match mx = rx.Match(name);
            if (mx.Captures.Count > 0)
            {
                throw new FormatException(@"the worksheet name must not contain the characters [  ]  * ? / \ ");
            }
            sheetName = name;
        }

        /// <summary>
        /// Sets the name of the worksheet
        /// </summary>
        /// <param name="name">Name of the worksheet</param>
        /// <param name="sanitize">If true, the filename will be sanitized automatically according to the specifications of Excel</param>
        /// <exception cref="WorksheetException">WorksheetException Thrown if no workbook is referenced. This information is necessary to determine whether the name already exists</exception>
        public void SetSheetName(string name, bool sanitize)
        {
            if (WorkbookReference == null)
            {
                throw new WorksheetException("MissingReferenceException", "The worksheet name cannot be sanitized because no workbook is referenced");
            }
            sheetName = ""; // Empty name (temporary) to prevent conflicts during sanitizing
            sheetName = SanitizeWorksheetName(name, WorkbookReference);
        }

        /// <summary>
        /// Sets the horizontal split of the worksheet into two panes. The measurement in characters cannot be used to freeze panes
        /// </summary>
        /// <param name="topPaneHeight">Height (similar to row height) from top of the worksheet to the split line in characters</param>
        /// <param name="topLeftCell">Top Left cell address of the bottom right pane (if applicable). Only the row component is important in a horizontal split</param>
        /// <param name="activePane">Active pane in the split window</param>
        public void SetHorizontalSplit(float topPaneHeight, Address topLeftCell, WorksheetPane activePane)
        {
            SetSplit(null, topPaneHeight, topLeftCell, activePane);
        }

        /// <summary>
        /// Sets the horizontal split of the worksheet into two panes. The measurement in rows can be used to split and freeze panes
        /// </summary>
        /// <param name="numberOfRowsFromTop">Number of rows from top of the worksheet to the split line. The particular row heights are considered</param>
        /// <param name="freeze">If true, all panes are frozen, otherwise remains movable</param>
        /// <param name="topLeftCell">Top Left cell address of the bottom right pane (if applicable). Only the row component is important in a horizontal split</param>
        /// <param name="activePane">Active pane in the split window</param>
        /// <exception cref="WorksheetException">WorksheetException Thrown if the row number of the top left cell is smaller the split panes number of rows from top, if freeze is applied</exception>
        public void SetHorizontalSplit(int numberOfRowsFromTop, bool freeze, Address topLeftCell, WorksheetPane activePane)
        {
            SetSplit(null, numberOfRowsFromTop, freeze, topLeftCell, activePane);
        }

        /// <summary>
        /// Sets the vertical split of the worksheet into two panes. The measurement in characters cannot be used to freeze panes
        /// </summary>
        /// <param name="leftPaneWidth">Width (similar to column width) from left of the worksheet to the split line in characters</param>
        /// <param name="topLeftCell">Top Left cell address of the bottom right pane (if applicable). Only the column component is important in a vertical split</param>
        /// <param name="activePane">Active pane in the split window</param>
        public void SetVerticalSplit(float leftPaneWidth, Address topLeftCell, WorksheetPane activePane)
        {
            SetSplit(leftPaneWidth, null, topLeftCell, activePane);
        }

        /// <summary>
        /// Sets the vertical split of the worksheet into two panes. The measurement in columns can be used to split and freeze panes
        /// </summary>
        /// <param name="numberOfColumnsFromLeft">Number of columns from left of the worksheet to the split line. The particular column widths are considered</param>
        /// <param name="freeze">If true, all panes are frozen, otherwise remains movable</param>
        /// <param name="topLeftCell">Top Left cell address of the bottom right pane (if applicable). Only the column component is important in a vertical split</param>
        /// <param name="activePane">Active pane in the split window</param>
        /// <exception cref="WorksheetException">WorksheetException Thrown if the column number of the top left cell is smaller the split panes number of columns from left, 
        /// if freeze is applied</exception>
        public void SetVerticalSplit(int numberOfColumnsFromLeft, bool freeze, Address topLeftCell, WorksheetPane activePane)
        {
            SetSplit(numberOfColumnsFromLeft, null, freeze, topLeftCell, activePane);
        }

        /// <summary>
        /// Sets the horizontal and vertical split of the worksheet into four panes. The measurement in rows and columns can be used to split and freeze panes
        /// </summary>
        /// <param name="numberOfColumnsFromLeft">Number of columns from left of the worksheet to the split line. The particular column widths are considered.<br/>
        /// The parameter is nullable. If left null, the method acts identical to <see cref="SetHorizontalSplit(int, bool, Address, WorksheetPane)"/></param>
        /// <param name="numberOfRowsFromTop">Number of rows from top of the worksheet to the split line. The particular row heights are considered.<br/>
        /// The parameter is nullable. If left null, the method acts identical to <see cref="SetVerticalSplit(int, bool, Address, WorksheetPane)"/></param>
        /// <param name="freeze">If true, all panes are frozen, otherwise remains movable</param>
        /// <param name="topLeftCell">Top Left cell address of the bottom right pane (if applicable)</param>
        /// <param name="activePane">Active pane in the split window</param>
        /// <exception cref="WorksheetException">WorksheetException Thrown if the address of the top left cell is smaller the split panes address, if freeze is applied</exception>
        public void SetSplit(int? numberOfColumnsFromLeft, int? numberOfRowsFromTop, bool freeze, Address topLeftCell, WorksheetPane activePane)
        {
            if (freeze)
            {
                if (numberOfColumnsFromLeft != null && topLeftCell.Column < numberOfColumnsFromLeft.Value)
                {
                    throw new WorksheetException("InvalidTopLeftCellException", "The column number " + topLeftCell.Column +
                        " is not valid for a frozen, vertical split with the split pane column number " + numberOfColumnsFromLeft.Value);
                }
                if (numberOfRowsFromTop != null && topLeftCell.Row < numberOfRowsFromTop.Value)
                {
                    throw new WorksheetException("InvalidTopLeftCellException", "The row number " + topLeftCell.Row +
                        " is not valid for a frozen, horizontal split height the split pane row number " + numberOfRowsFromTop.Value);
                }
            }
            this.paneSplitLeftWidth = null;
            this.paneSplitTopHeight = null;
            this.freezeSplitPanes = freeze;
            int row = numberOfRowsFromTop != null ? numberOfRowsFromTop.Value : 0;
            int column = numberOfColumnsFromLeft != null ? numberOfColumnsFromLeft.Value : 0;
            this.paneSplitAddress = new Address(column, row);
            this.paneSplitTopLeftCell = topLeftCell;
            this.activePane = activePane;
        }

        /// <summary>
        /// Sets the horizontal and vertical split of the worksheet into four panes. The measurement in characters cannot be used to freeze panes
        /// </summary>
        /// <param name="leftPaneWidth">Width (similar to column width) from left of the worksheet to the split line in characters.<br/>
        /// The parameter is nullable. If left null, the method acts identical to <see cref="SetHorizontalSplit(float, Address, WorksheetPane)"/></param>
        /// <param name="topPaneHeight">Height (similar to row height) from top of the worksheet to the split line in characters.<br/>
        /// The parameter is nullable. If left null, the method acts identical to <see cref="SetVerticalSplit(float, Address, WorksheetPane)"/></param>
        /// <param name="topLeftCell">Top Left cell address of the bottom right pane (if applicable)</param>
        /// <param name="activePane">Active pane in the split window</param>
        public void SetSplit(float? leftPaneWidth, float? topPaneHeight, Address topLeftCell, WorksheetPane activePane)
        {
            this.paneSplitLeftWidth = leftPaneWidth;
            this.paneSplitTopHeight = topPaneHeight;
            this.freezeSplitPanes = null;
            this.paneSplitAddress = null;
            this.paneSplitTopLeftCell = topLeftCell;
            this.activePane = activePane;
        }

        /// <summary>
        /// Resets splitting of the worksheet into panes, as well as their freezing 
        /// </summary>
        public void ResetSplit()
        {
            this.paneSplitLeftWidth = null;
            this.paneSplitTopHeight = null;
            this.freezeSplitPanes = null;
            this.paneSplitAddress = null;
            this.paneSplitTopLeftCell = null;
            this.activePane = null;
        }


        #region static_methods

        /// <summary>
        /// Sanitizes a worksheet name
        /// </summary>
        /// <param name="input">Name to sanitize</param>
        /// <param name="workbook">Workbook reference</param>
        /// <returns>Name of the sanitized worksheet</returns>
        public static string SanitizeWorksheetName(string input, Workbook workbook)
        {

            if (string.IsNullOrEmpty(input))
            {
                input = "Sheet1";
            }
            int len;
            if (input.Length > MAX_WORKSHEER_NAME_LENGTH)
            {
                len = MAX_WORKSHEER_NAME_LENGTH;
            }
            else
            {
                len = input.Length;
            }
            StringBuilder sb = new StringBuilder(MAX_WORKSHEER_NAME_LENGTH);
            char c;
            for (int i = 0; i < len; i++)
            {
                c = input[i];
                if (c == '[' || c == ']' || c == '*' || c == '?' || c == '\\' || c == '/')
                { sb.Append('_'); }
                else
                { sb.Append(c); }
            }

            return GetUnusedWorksheetName(sb.ToString(), workbook);
        }

        /// <summary>
        /// Determines the next unused worksheet name in the passed workbook
        /// </summary>
        /// <param name="name">Original name to start the check</param>
        /// <param name="workbook">Workbook to look for existing worksheets</param>
        /// <returns>Not yet used worksheet name</returns>
        /// <remarks>The 'rare' case where 10^31 Worksheets exists (leads to a crash) is deliberately not handled, 
        /// since such a number of sheets would consume at least a quintillion bytes of RAM... what is vastly out of the 64 bit range</remarks>
        private static string GetUnusedWorksheetName(string name, Workbook workbook)
        {
            string originalName = name;
            int number = 0;
            while (true)
            {
                if (!WorksheetExists(name, workbook))
                { return name; }
                string numberString = Utils.ToString(number);
                if (originalName.Length + numberString.Length > MAX_WORKSHEER_NAME_LENGTH)
                {
                    name = originalName.Substring(0, MAX_WORKSHEER_NAME_LENGTH - numberString.Length - 1) + numberString;
                }
                else
                {
                    name = originalName + numberString;
                }
                number++;
            }
        }

        /// <summary>
        /// Checks whether a worksheet with the given name exists
        /// </summary>
        /// <param name="name">Name to check</param>
        /// <param name="workbook">Workbook reference</param>
        /// <returns>True if the name exits, otherwise false</returns>
        private static bool WorksheetExists(String name, Workbook workbook)
        {
            int len = workbook.Worksheets.Count;
            for (int i = 0; i < len; i++)
            {
                if (workbook.Worksheets[i].SheetName == name)
                {
                    return true;
                }
            }
            return false;
        }

        #endregion


        #endregion

    }
}
