/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using NanoXLSX.Exceptions;
using NanoXLSX.Interfaces;
using NanoXLSX.Registry;
using NanoXLSX.Styles;
using NanoXLSX.Utils;
using FormatException = NanoXLSX.Exceptions.FormatException;

namespace NanoXLSX
{
    /// <summary>
    /// Class representing a worksheet of a workbook
    /// </summary>
    public class Worksheet
    {
        static Worksheet()
        {
            PackageRegistry.Initialize();
        }

        #region constants
        /// <summary>
        /// Maximum number of characters a worksheet name can have
        /// </summary>
        public static readonly int MAX_WORKSHEET_NAME_LENGTH = 31;
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
        /// <summary>
        /// Automatic zoom factor of a worksheet
        /// </summary>
        public const int AUTO_ZOOM_FACTOR = 0;
        /// <summary>
        /// Minimum zoom factor of a worksheet. If set to this value, the zoom is set to automatic
        /// </summary>
        public const int MIN_ZOOM_FACTOR = 10;
        /// <summary>
        /// Maximum zoom factor of a worksheet
        /// </summary>
        public const int MAX_ZOOM_FACTOR = 400;
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

        /// <summary>
        /// Enum to define how a worksheet is displayed in the spreadsheet application (Excel)
        /// </summary>
        public enum SheetViewType
        {
            /// <summary>The worksheet is displayed without pagination (default)</summary>
            normal,
            /// <summary>The worksheet is displayed with indicators where the page would break if it were printed</summary>
            pageBreakPreview,
            /// <summary>The worksheet is displayed like it would be printed</summary>
            pageLayout
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
        private bool hidden;
        private Workbook workbookReference;
        private IPassword sheetProtectionPassword;
        private List<Range> selectedCells;
        private bool? freezeSplitPanes;
        private float? paneSplitLeftWidth;
        private float? paneSplitTopHeight;
        private Address? paneSplitTopLeftCell;
        private Address? paneSplitAddress;
        private WorksheetPane? activePane;
        private int sheetID;
        private SheetViewType viewType;
        private Dictionary<SheetViewType, int> zoomFactor;
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
                    throw new RangeException("The passed default column width is out of range (" + MIN_COLUMN_WIDTH + " to " + MAX_COLUMN_WIDTH + ")");
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
                    throw new RangeException("The passed default row height is out of range (" + MIN_ROW_HEIGHT + " to " + MAX_ROW_HEIGHT + ")");
                }
                defaultRowHeight = value;
            }
        }

        /// <summary>
        /// Gets the hidden rows as dictionary with the zero-based row number as key and a boolean as value. True indicates hidden, false visible.
        /// </summary>
        /// \remark <remarks>Entries with the value false are not affecting the worksheet. These entries can be removed</remarks>
        public Dictionary<int, bool> HiddenRows
        {
            get { return hiddenRows; }
        }

        /// <summary>
        /// Gets defined row heights as dictionary with the zero-based row number as key and the height (float from 0 to 409.5) as value
        /// </summary>
        public Dictionary<int, float> RowHeights
        {
            get { return rowHeights; }
        }

        /// <summary>
        /// Gets the merged cells (only references) as dictionary with the cell address as key and the range object as value
        /// </summary>
        public Dictionary<string, Range> MergedCells
        {
            get { return mergedCells; }
        }

        /// <summary>
        /// Gets the cell ranges of selected cells of this worksheet. Returns ans empty list if no cells are selected
        /// </summary>
        public List<Range> SelectedCells
        {
            get { return selectedCells; }
        }

        /// <summary>
        /// Gets or sets the internal ID of the worksheet
        /// </summary>
        public int SheetID
        {
            get => sheetID;
            set
            {
                if (value < 1)
                {
                    throw new FormatException("The ID " + value + " is invalid. Worksheet IDs must be >0");
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
            set { SetSheetName(value); }
        }

        /// <summary>
        /// Password instance of the worksheet protection. If a password was set, the pain text representation and the hash can be read from the instance
        /// </summary>
        /// \remark <remarks>The password of this property is stored in plain text at runtime but not stored to a worksheet. The plain text password cannot be recovered when loading a workbook. The hash is retrieved and can be reused</remarks>
        public virtual IPassword SheetProtectionPassword
        {
            get { return sheetProtectionPassword; }
            internal set { sheetProtectionPassword = value; }
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
        public Workbook WorkbookReference
        {
            get { return workbookReference; }
            set
            {
                workbookReference = value;
                if (value != null)
                {
                    workbookReference.ValidateWorksheets();
                }
            }
        }

        /// <summary>
        /// gets or sets whether the worksheet is hidden. If true, the worksheet is not listed in the worksheet tabs of the workbook.<br />
        /// If the worksheet is not part of a workbook, or the only one in the workbook, an exception will be thrown.<br />
        /// If the worksheet is the selected one, and attempted to set hidden, an exception will be thrown. Define another selected worksheet prior to this call, in this case.
        /// </summary>
        public bool Hidden
        {
            get { return hidden; }
            set
            {
                hidden = value;
                if (value && workbookReference != null)
                {
                    workbookReference.ValidateWorksheets();
                }
            }
        }

        /// <summary>
        /// Gets the height of the upper, horizontal split pane, measured from the top of the window.<br />
        /// The value is nullable. If null, no horizontal split of the worksheet is applied.<br />
        /// The value is only applicable to split the worksheet into panes, but not to freeze them.<br />
        /// See also: <see cref="PaneSplitAddress"/>
        /// </summary>
        /// \remark <remarks>Note: This value will be modified to the Excel-internal representation, 
        /// calculated by <see cref="Utils.GetInternalPaneSplitHeight(float)"/>.</remarks>
        public float? PaneSplitTopHeight
        {
            get { return paneSplitTopHeight; }
        }

        /// <summary>
        /// Gets the width of the left, vertical split pane, measured from the left of the window.<br />
        /// The value is nullable. If null, no vertical split of the worksheet is applied<br />
        /// The value is only applicable to split the worksheet into panes, but not to freeze them.<br />
        /// See also: <see cref="PaneSplitAddress"/>
        /// </summary>
        /// \remark <remarks>Note: This value will be modified to the Excel-internal representation, 
        /// calculated by <see cref="Utils.GetInternalColumnWidth(float, float, float)"/>.</remarks>
        public float? PaneSplitLeftWidth
        {
            get { return paneSplitLeftWidth; }
        }

        /// <summary>
        /// Gets whether split panes are frozen.<br />
        /// The value is nullable. If null, no freezing is applied. This property also does not apply if <see cref="PaneSplitAddress"/> is null
        /// </summary>
        public bool? FreezeSplitPanes
        {
            get { return freezeSplitPanes; }
        }

        /// <summary>
        /// Gets the Top Left cell address of the bottom right pane if applicable and splitting is applied.<br />
        /// The column is only relevant for vertical split, whereas the row component is only relevant for a horizontal split.<br />
        /// The value is nullable. If null, no splitting was defined.
        /// </summary>
        public Address? PaneSplitTopLeftCell
        {
            get { return paneSplitTopLeftCell; }
        }

        /// <summary>
        /// Gets the split address for frozen panes or if pane split was defined in number of columns and / or rows.<br /> 
        /// For vertical splits, only the column component is considered. For horizontal splits, only the row component is considered.<br />
        /// The value is nullable. If null, no frozen panes or split by columns / rows are applied to the worksheet. 
        /// However, splitting can still be applied, if the value is defined in characters.<br />
        /// See also: <see cref="PaneSplitLeftWidth"/> and <see cref="PaneSplitTopHeight"/> for splitting in characters (without freezing)
        /// </summary>
        public Address? PaneSplitAddress
        {
            get { return paneSplitAddress; }
        }


        /// <summary>
        /// Gets the active Pane is splitting is applied.<br />
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

        /// <summary>
        /// Gets or sets whether grid lines are visible on the current worksheet. Default is true
        /// </summary>
        public bool ShowGridLines { get; set; }

        /// <summary>
        /// Gets or sets whether the column and row headers are visible on the current worksheet. Default is true
        /// </summary>
        public bool ShowRowColumnHeaders { get; set; }

        /// <summary>
        /// Gets or sets whether a ruler is displayed over the column headers. This value only applies if <see cref="ViewType"/> is set to <see cref="SheetViewType.pageLayout"/>. Default is true
        /// </summary>
        public bool ShowRuler { get; set; }

        /// <summary>
        /// Gets or sets how the current worksheet is displayed in the spreadsheet application (Excel)
        /// </summary>
        public SheetViewType ViewType
        {
            get
            {
                return viewType;
            }
            set
            {
                viewType = value;
                SetZoomFactor(value, 100);
            }
        }

        /// <summary>
        /// Gets or sets the zoom factor of the <see cref="ViewType"/> of the current worksheet. If <see cref="AUTO_ZOOM_FACTOR"/>, the zoom factor is set to automatic
        /// </summary>
        /// \remark <remarks>It is possible to add further zoom factors for inactive view types, using the function <see cref="SetZoomFactor(SheetViewType, int)"/> </remarks>
        /// <exception cref="WorksheetException">Throws a WorksheetException if the zoom factor is not <see cref="AUTO_ZOOM_FACTOR"/> or below <see cref="MIN_ZOOM_FACTOR"/> or above <see cref="MAX_ZOOM_FACTOR"/></exception>
        public int ZoomFactor
        {
            set
            {
                SetZoomFactor(viewType, value);
            }
            get
            {
                return zoomFactor[viewType];
            }
        }

        /// <summary>
        /// Gets all defined zoom factors per <see cref="SheetViewType"/> of the current worksheet. Use <see cref="SetZoomFactor(SheetViewType, int)"/> to define the values
        /// </summary>
        public Dictionary<SheetViewType, int> ZoomFactors
        {
            get
            {
                return zoomFactor;
            }
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
            selectedCells = new List<Range>();
            sheetProtectionValues = new List<SheetProtectionValue>();
            hiddenRows = new Dictionary<int, bool>();
            columns = new Dictionary<int, Column>();
            activeStyle = null;
            workbookReference = null;
            viewType = SheetViewType.normal;
            zoomFactor = new Dictionary<SheetViewType, int>();
            zoomFactor.Add(viewType, 100);
            ShowGridLines = true;
            ShowRowColumnHeaders = true;
            ShowRuler = true;
            sheetProtectionPassword = new LegacyPassword(LegacyPassword.PasswordType.WORKSHEET_PROTECTION);
        }

        /// <summary>
        /// Constructor with worksheet name
        /// </summary>
        /// \remark <remarks>Note that the worksheet name is not checked and fully sanitized against other worksheets with this operation. This is later performed when the worksheet is added to the workbook</remarks>
        public Worksheet(string name)
            : this()
        {
            SetSheetName(name);
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
            SetSheetName(name);
            SheetID = id;
            workbookReference = reference;
        }

        #endregion

        #region methods_AddNextCell

        /// <summary>
        /// Adds an object to the next cell position. If the type of the value does not match with one of the supported data types, it will be cast to a String. 
        /// A prepared object of the type Cell will not be cast but adjusted
        /// </summary>
        /// \remark <remarks>Recognized are the following data types: Cell (prepared object), string, int, double, float, long, DateTime, TimeSpan, bool. 
        /// All other types will be cast into a string using the default ToString() method</remarks>
        /// <param name="value">Unspecified value to insert</param>
        /// <exception cref="RangeException">Throws a RangeException if the next cell is out of range (on row or column)</exception>
        public void AddNextCell(object value)
        {
            AddNextCell(CastValue(value, currentColumnNumber, currentRowNumber), true, null);
        }


        /// <summary>
        /// Adds an object to the next cell position. If the type of the value does not match with one of the supported data types, it will be cast to a String. 
        /// A prepared object of the type Cell will not be cast but adjusted
        /// </summary>
        /// \remark <remarks>Recognized are the following data types: Cell (prepared object), string, int, double, float, long, DateTime, TimeSpan, bool. 
        /// All other types will be cast into a string using the default ToString() method</remarks>
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
        /// \remark <remarks>Recognized are the following data types: string, int, double, float, long, DateTime, TimeSpan, bool. All other types will be cast into a string using the default ToString() method.<br />
        /// If the cell object already has a style definition, and a style or active style is defined, the cell style will be merged, otherwise just set</remarks>
        /// <exception cref="StyleException">Throws a StyleException if the default style was malformed</exception>
        private void AddNextCell(Cell cell, bool incremental, Style style)
        {
            // date and time styles are already defined by the passed cell object
            if (style != null || (activeStyle != null && useActiveStyle))
            {
                if (cell.CellStyle == null && useActiveStyle)
                {
                    cell.SetStyle(activeStyle);
                }
                else if (cell.CellStyle == null && style != null)
                {
                    cell.SetStyle(style);
                }
                else if (cell.CellStyle != null && useActiveStyle)
                {
                    Style mixedStyle = (Style)cell.CellStyle.Copy();
                    mixedStyle.Append(activeStyle);
                    cell.SetStyle(mixedStyle);
                }
                else if (cell.CellStyle != null && style != null)
                {
                    Style mixedStyle = (Style)cell.CellStyle.Copy();
                    mixedStyle.Append(style);
                    cell.SetStyle(mixedStyle);
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
        /// Adds an object to the defined cell address. If the type of the value does not match with one of the supported data types, it will be cast to a String. 
        /// A prepared object of the type Cell will not be cast but adjusted
        /// </summary>
        /// <param name="value">Unspecified value to insert</param>
        /// <param name="columnNumber">Column number (zero based)</param>
        /// <param name="rowNumber">Row number (zero based)</param>
        /// \remark <remarks>Recognized are the following data types: Cell (prepared object), string, int, double, float, long, DateTime, TimeSpan, bool. 
        /// All other types will be cast into a string using the default ToString() method</remarks>
        /// <exception cref="RangeException">Throws a RangeException if the passed cell address is out of range</exception>
        public void AddCell(object value, int columnNumber, int rowNumber)
        {
            AddNextCell(CastValue(value, columnNumber, rowNumber), false, null);
        }

        /// <summary>
        /// Adds an object to the defined cell address. If the type of the value does not match with one of the supported data types, it will be cast to a String. 
        /// A prepared object of the type Cell will not be cast but adjusted
        /// </summary>
        /// <param name="value">Unspecified value to insert</param>
        /// <param name="columnNumber">Column number (zero based)</param>
        /// <param name="rowNumber">Row number (zero based)</param>
        /// <param name="style">Style to apply on the cell</param>
        /// \remark <remarks>Recognized are the following data types: Cell (prepared object), string, int, double, float, long, DateTime, TimeSpan, bool. 
        /// All other types will be cast into a string using the default ToString() method</remarks>
        /// <exception cref="StyleException">Throws a StyleException if the passed style is malformed</exception>
        /// <exception cref="RangeException">Throws a RangeException if the passed cell address is out of range</exception>
        public void AddCell(object value, int columnNumber, int rowNumber, Style style)
        {
            AddNextCell(CastValue(value, columnNumber, rowNumber), false, style);
        }


        /// <summary>
        /// Adds an object to the defined cell address. If the type of the value does not match with one of the supported data types, it will be cast to a String. 
        /// A prepared object of the type Cell will not be cast but adjusted
        /// </summary>
        /// <param name="value">Unspecified value to insert</param>
        /// <param name="address">Cell address in the format A1 - XFD1048576</param>
        /// \remark <remarks>Recognized are the following data types: Cell (prepared object), string, int, double, float, long, DateTime, TimeSpan, bool. 
        /// All other types will be cast into a string using the default ToString() method</remarks>
        /// <exception cref="RangeException">Throws a RangeException if the passed cell address is out of range</exception>
        /// <exception cref="NanoXLSX.Exceptions.FormatException">Throws a FormatException if the passed cell address is malformed</exception>
        public void AddCell(object value, string address)
        {
            int column;
            int row;
            Cell.ResolveCellCoordinate(address, out column, out row);
            AddCell(value, column, row);
        }

        /// <summary>
        /// Adds an object to the defined cell address. If the type of the value does not match with one of the supported data types, it will be cast to a String. 
        /// A prepared object of the type Cell will not be cast but adjusted
        /// </summary>
        /// <param name="value">Unspecified value to insert</param>
        /// <param name="address">Cell address in the format A1 - XFD1048576</param>
        /// <param name="style">Style to apply on the cell</param>
        /// \remark <remarks>Recognized are the following data types: Cell (prepared object), string, int, double, float, long, DateTime, TimeSpan, 
        /// bool. All other types will be cast into a string using the default ToString() method</remarks>
        /// <exception cref="StyleException">Throws a StyleException if the passed style is malformed</exception>
        /// <exception cref="RangeException">Throws a RangeException if the passed cell address is out of range</exception>
        /// <exception cref="NanoXLSX.Exceptions.FormatException">Throws a FormatException if the passed cell address is malformed</exception>
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
        /// <exception cref="RangeException">Throws a RangeException if the passed cell address is out of range</exception>
        /// <exception cref="NanoXLSX.Exceptions.FormatException">Throws a FormatException if the passed cell address is malformed</exception>
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
        /// <exception cref="StyleException">Throws a StyleException if the passed style was malformed</exception>
        /// <exception cref="RangeException">Throws a RangeException if the passed cell address is out of range</exception>
        /// <exception cref="NanoXLSX.Exceptions.FormatException">Throws a FormatException if the passed cell address is malformed</exception>
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
        /// <param name="columnNumber">Column number (zero based)</param>
        /// <param name="rowNumber">Row number (zero based)</param>
        /// <exception cref="RangeException">Throws a RangeException if the passed cell address is out of range</exception>
        public void AddCellFormula(string formula, int columnNumber, int rowNumber)
        {
            Cell c = new Cell(formula, Cell.CellType.FORMULA, columnNumber, rowNumber);
            AddNextCell(c, false, null);
        }

        /// <summary>
        /// Adds a cell formula as string to the defined cell address
        /// </summary>
        /// <param name="formula">Formula to insert</param>
        /// <param name="columnNumber">Column number (zero based)</param>
        /// <param name="rowNumber">Row number (zero based)</param>
        /// <param name="style">Style to apply on the cell</param>
        /// <exception cref="RangeException">Throws a RangeException if the passed cell address is out of range</exception>
        public void AddCellFormula(string formula, int columnNumber, int rowNumber, Style style)
        {
            Cell c = new Cell(formula, Cell.CellType.FORMULA, columnNumber, rowNumber);
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
        /// Adds a list of object values to a defined cell range. If the type of the particular value does not match with one of the supported data types, it will be cast to a String. 
        /// Prepared objects of the type Cell will not be cast but adjusted
        /// </summary>
        /// <param name="values">List of unspecified objects to insert</param>
        /// <param name="startAddress">Start address</param>
        /// <param name="endAddress">End address</param>
        /// \remark <remarks>The data types in the passed list can be mixed. Recognized are the following data types: string, int, double, float, long, DateTime, TimeSpan, bool. 
        /// All other types will be cast into a string using the default ToString() method</remarks>
        /// <exception cref="RangeException">Throws a RangeException if the number of cells resolved from the range differs from the number of passed values</exception>
        public void AddCellRange(IReadOnlyList<object> values, Address startAddress, Address endAddress)
        {
            AddCellRangeInternal(values, startAddress, endAddress, null);
        }

        /// <summary>
        /// Adds a list of object values to a defined cell range. If the type of the particular value does not match with one of the supported data types, it will be cast to a String. 
        /// Prepared objects of the type Cell will not be cast but adjusted
        /// </summary>
        /// <param name="values">List of unspecified objects to insert</param>
        /// <param name="startAddress">Start address</param>
        /// <param name="endAddress">End address</param>
        /// <param name="style">Style to apply on the all cells of the range</param>
        /// \remark <remarks>The data types in the passed list can be mixed. Recognized are the following data types: Cell (prepared object), string, int, double, float, long, DateTime, TimeSpan, bool. 
        /// All other types will be cast into a string using the default ToString() method</remarks>
        /// <exception cref="RangeException">Throws a RangeException if the number of cells resolved from the range differs from the number of passed values</exception>
        /// <exception cref="StyleException">Throws a StyleException if the passed style is malformed</exception>
        public void AddCellRange(IReadOnlyList<object> values, Address startAddress, Address endAddress, Style style)
        {
            AddCellRangeInternal(values, startAddress, endAddress, style);
        }

        /// <summary>
        /// Adds a list of object values to a defined cell range. If the type of the particular value does not match with one of the supported data types, it will be cast to a String. 
        /// Prepared objects of the type Cell will not be cast but adjusted
        /// </summary>
        /// <param name="values">List of unspecified objects to insert</param>
        /// <param name="cellRange">Cell range as string in the format like A1:D1 or X10:X22</param>
        /// \remark <remarks>The data types in the passed list can be mixed. Recognized are the following data types: Cell (prepared object), string, int, double, float, long, DateTime, TimeSpan, bool. 
        /// All other types will be cast into a string using the default ToString() method</remarks>
        /// <exception cref="RangeException">Throws a RangeException if the number of cells resolved from the range differs from the number of passed values</exception>
        /// <exception cref="NanoXLSX.Exceptions.FormatException">Throws a FormatException if the passed cell range is malformed</exception>
        public void AddCellRange(IReadOnlyList<object> values, string cellRange)
        {
            Range range = Cell.ResolveCellRange(cellRange);
            AddCellRangeInternal(values, range.StartAddress, range.EndAddress, null);
        }

        /// <summary>
        /// Adds a list of object values to a defined cell range. If the type of the particular value does not match with one of the supported data types, it will be cast to a String. 
        /// Prepared objects of the type Cell will not be cast but adjusted
        /// </summary>
        /// <param name="values">List of unspecified objects to insert</param>
        /// <param name="cellRange">Cell range as string in the format like A1:D1 or X10:X22</param>
        /// <param name="style">Style to apply on the all cells of the range</param>
        /// \remark <remarks>The data types in the passed list can be mixed. Recognized are the following data types: Cell (prepared object), string, int, double, float, long, DateTime, TimeSpan, bool. 
        /// All other types will be cast into a string using the default ToString() method</remarks>
        /// <exception cref="RangeException">Throws a RangeException if the number of cells resolved from the range differs from the number of passed values</exception>
        /// <exception cref="StyleException">Throws a StyleException if the passed style is malformed</exception>
        /// <exception cref="NanoXLSX.Exceptions.FormatException">Throws a FormatException if the passed cell range is malformed</exception>
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
        /// \remark <remarks>The data types in the passed list can be mixed. Recognized are the following data types: Cell (prepared object), string, int, double, float, long, DateTime, TimeSpan, bool. 
        /// All other types will be cast into a string using the default ToString() method</remarks>
        /// <exception cref="RangeException">Throws a RangeException if the number of cells differs from the number of passed values</exception>
        private void AddCellRangeInternal<T>(IReadOnlyList<T> values, Address startAddress, Address endAddress, Style style)
        {
            List<Address> addresses = Cell.GetCellRange(startAddress, endAddress) as List<Address>;
            if (values.Count != addresses.Count)
            {
                throw new RangeException("The number of passed values (" + values.Count + ") differs from the number of cells within the range (" + addresses.Count + ")");
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
        /// <param name="columnNumber">Column number (zero based)</param>
        /// <param name="rowNumber">Row number (zero based)</param>
        /// <returns>Returns true if the cell could be removed (existed), otherwise false (did not exist)</returns>
        /// <exception cref="RangeException">Throws a RangeException if the passed cell address is out of range</exception>
        public bool RemoveCell(int columnNumber, int rowNumber)
        {
            string address = Cell.ResolveCellAddress(columnNumber, rowNumber);
            return cells.Remove(address);
        }

        /// <summary>
        /// Removes a previous inserted cell at the defined address
        /// </summary>
        /// <param name="address">Cell address in the format A1 - XFD1048576</param>
        /// <returns>Returns true if the cell could be removed (existed), otherwise false (did not exist)</returns>
        /// <exception cref="RangeException">Throws a RangeException if the passed cell address is out of range</exception>
        /// <exception cref="NanoXLSX.Exceptions.FormatException">Throws a FormatException if the passed cell address is malformed</exception>
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
        /// Sets the passed style on the passed cell range. If cells are already existing, the style will be added or replaced.
        /// Otherwise, an empty cell will be added with the assigned style. If the passed style is null, all styles will be removed on existing cells and no additional (empty) cells are added to the worksheet
        /// </summary>
        /// <param name="cellRange">Cell range to apply the style</param>
        /// <param name="style">Style to apply</param>
        /// \remark <remarks>Note: This method may invalidate an existing date or time value since dates and times are defined by specific style. The result of a redefinition will be a number, instead of a date or time</remarks>
        public void SetStyle(Range cellRange, Style style)
        {
            IReadOnlyList<Address> addresses = cellRange.ResolveEnclosedAddresses();
            foreach (Address address in addresses)
            {
                string key = address.GetAddress();
                if (this.cells.ContainsKey(key))
                {
                    if (style == null)
                    {
                        cells[key].RemoveStyle();
                    }
                    else
                    {
                        cells[key].SetStyle(style);
                    }
                }
                else
                {
                    if (style != null)
                    {
                        AddCell(null, address.Column, address.Row, style);
                    }
                }
            }
        }

        /// <summary>
        /// Sets the passed style on the passed cell range, derived from a start and end address. If cells are already existing, the style will be added or replaced.
        /// Otherwise, an empty cell will be added with the assigned style. If the passed style is null, all styles will be removed on existing cells and no additional (empty) cells are added to the worksheet
        /// </summary>
        /// <param name="startAddress">Start address of the cell range</param>
        /// <param name="endAddress">End address of the cell range</param>
        /// <param name="style">Style to apply or null to clear the range</param>
        /// \remark <remarks>Note: This method may invalidate an existing date or time value since dates and times are defined by specific style. The result of a redefinition will be a number, instead of a date or time</remarks>
        public void SetStyle(Address startAddress, Address endAddress, Style style)
        {
            SetStyle(new Range(startAddress, endAddress), style);
        }

        /// <summary>
        /// Sets the passed style on the passed (singular) cell address. If the cell is already existing, the style will be added or replaced.
        /// Otherwise, an empty cell will be added with the assigned style. If the passed style is null, all styles will be removed on existing cells and no additional (empty) cells are added to the worksheet
        /// </summary>
        /// <param name="address">Cell address to apply the style</param>
        /// <param name="style">Style to apply or null to clear the range</param>
        /// \remark <remarks>Note: This method may invalidate an existing date or time value since dates and times are defined by specific style. The result of a redefinition will be a number, instead of a date or time</remarks>
        public void SetStyle(Address address, Style style)
        {
            SetStyle(address, address, style);
        }

        /// <summary>
        /// Sets the passed style on the passed address expression. Such an expression may be a single cell or a cell range.
        /// If the cell is already existing, the style will be added or replaced.
        /// Otherwise, an empty cell will be added with the assigned style. If the passed style is null, all styles will be removed on existing cells and no additional (empty) cells are added to the worksheet
        /// </summary>
        /// <param name="addressExpression">Expression of a cell address or range of addresses</param>
        /// <param name="style">Style to apply or null to clear the range</param>
        /// \remark <remarks>Note: This method may invalidate an existing date or time value since dates and times are defined by specific style. The result of a redefinition will be a number, instead of a date or time</remarks>
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
                throw new FormatException("The passed address'" + addressExpression + "' is neither a cell address, nor a range");
            }
        }

        #endregion

        #region boundaryFunctions
        /// <summary>
        /// Gets the first existing column number in the current worksheet (zero-based)
        /// </summary>
        /// <returns>Zero-based column number. In case of an empty worksheet, -1 will be returned</returns>
        /// \remark <remarks>GetFirstColumnNumber() will not return the first column with data in any case. If there is a formatted but empty cell (or many) before the first cell with data, 
        /// GetFirstColumnNumber() will return the column number of this empty cell. Use <see cref="GetFirstDataColumnNumber"/> in this case.</remarks>
        public int GetFirstColumnNumber()
        {
            return GetBoundaryNumber(false, true);
        }

        /// <summary>
        /// Gets the first existing column number with data in the current worksheet (zero-based)
        /// </summary>
        /// <returns>Zero-based column number. In case of an empty worksheet, -1 will be returned</returns>
        /// \remark <remarks>GetFirstDataColumnNumber() will ignore formatted but empty cells before the first column with data. 
        /// If you want the first defined column, use <see cref="GetFirstColumnNumber"/> instead.</remarks>
        public int GetFirstDataColumnNumber()
        {
            return GetBoundaryDataNumber(false, true, true);
        }

        /// <summary>
        /// Gets the first existing row number in the current worksheet (zero-based)
        /// </summary>
        /// <returns>Zero-based row number. In case of an empty worksheet, -1 will be returned</returns>
        /// \remark <remarks>GetFirstRowNumber() will not return the first row with data in any case. If there is a formatted but empty cell (or many) before the first cell with data, 
        /// GetFirstRowNumber() will return the row number of this empty cell. Use <see cref="GetFirstDataRowNumber"/> in this case.</remarks>
        public int GetFirstRowNumber()
        {
            return GetBoundaryNumber(true, true);
        }

        /// <summary>
        /// Gets the first existing row number with data in the current worksheet (zero-based)
        /// </summary>
        /// <returns>Zero-based row number. In case of an empty worksheet, -1 will be returned</returns>
        /// \remark <remarks>GetFirstDataRowNumber() will ignore formatted but empty cells before the first row with data. 
        /// If you want the first defined row, use <see cref="GetFirstRowNumber"/> instead.</remarks>
        public int GetFirstDataRowNumber()
        {
            return GetBoundaryDataNumber(true, true, true);
        }

        /// <summary>
        /// Gets the last existing column number in the current worksheet (zero-based)
        /// </summary>
        /// <returns>Zero-based column number. In case of an empty worksheet, -1 will be returned</returns>
        /// \remark <remarks>GetLastColumnNumber() will not return the last column with data in any case. If there is a formatted (or with the definition of AutoFilter, 
        /// column width or hidden state) but empty cell (or many) after the last cell with data, 
        /// GetLastColumnNumber() will return the column number of this empty cell. Use <see cref="GetLastDataColumnNumber"/> in this case.</remarks>
        public int GetLastColumnNumber()
        {
            return GetBoundaryNumber(false, false);
        }

        /// <summary>
        /// Gets the last existing column number with data in the current worksheet (zero-based)
        /// </summary>
        /// <returns>Zero-based column number. in case of an empty worksheet, -1 will be returned</returns>
        /// \remark <remarks>GetLastDataColumnNumber() will ignore formatted (or with the definition of AutoFilter, column width or hidden state) but empty cells after the last column with data. 
        /// If you want the last defined column, use <see cref="GetLastColumnNumber"/> instead.</remarks>
        public int GetLastDataColumnNumber()
        {
            return GetBoundaryDataNumber(false, false, true);
        }

        /// <summary>
        /// Gets the last existing row number in the current worksheet (zero-based)
        /// </summary>
        /// <returns>Zero-based row number. In case of an empty worksheet, -1 will be returned</returns>
        /// \remark <remarks>GetLastRowNumber() will not return the last row with data in any case. If there is a formatted (or with the definition of row height or hidden state) 
        /// but empty cell (or many) after the last cell with data, 
        /// GetLastRowNumber() will return the row number of this empty cell. Use <see cref="GetLastDataRowNumber"/> in this case.</remarks>
        public int GetLastRowNumber()
        {
            return GetBoundaryNumber(true, false);
        }


        /// <summary>
        /// Gets the last existing row number with data in the current worksheet (zero-based)
        /// </summary>
        /// <returns>Zero-based row number. in case of an empty worksheet, -1 will be returned</returns>
        /// \remark <remarks>GetLastDataColumnNumber() will ignore formatted (or with the definition of row height or hidden state) but empty cells after the last column with data. 
        /// If you want the last defined column, use <see cref="GetLastRowNumber"/> instead.</remarks>
        public int GetLastDataRowNumber()
        {
            return GetBoundaryDataNumber(true, false, true);
        }

        /// <summary>
        ///  Gets the last existing cell in the current worksheet (bottom right)
        /// </summary>
        /// <returns>Nullable Cell Address. If no cell address could be determined, null will be returned</returns>
        /// \remark <remarks>GetLastCellAddress() will not return the last cell with data in any case. If there is a formatted (or with definitions of hidden states, AutoFilters, heights or widths) 
        /// but empty cell (or many) after the last cell with data, 
        /// GetLastCellAddress() will return the address of this empty cell. Use <see cref="GetLastDataCellAddress"/> in this case.</remarks>

        public Address? GetLastCellAddress()
        {
            int lastRow = GetLastRowNumber();
            int lastColumn = GetLastColumnNumber();
            if (lastRow < 0 || lastColumn < 0)
            {
                return null;
            }
            return new Address(lastColumn, lastRow);
        }

        /// <summary>
        ///  Gets the last existing cell with data in the current worksheet (bottom right)
        /// </summary>
        /// <returns>Nullable Cell Address. If no cell address could be determined, null will be returned</returns>
        /// \remark <remarks>GetLastDataCellAddress() will ignore formatted (or with definitions of hidden states, AutoFilters, heights or widths) but empty cells after the last cell with data. 
        /// If you want the last defined cell, use <see cref="GetLastCellAddress"/> instead.</remarks>

        public Address? GetLastDataCellAddress()
        {
            int lastRow = GetLastDataRowNumber();
            int lastColumn = GetLastDataColumnNumber();
            if (lastRow < 0 || lastColumn < 0)
            {
                return null;
            }
            return new Address(lastColumn, lastRow);
        }

        /// <summary>
        ///  Gets the first existing cell in the current worksheet (bottom right)
        /// </summary>
        /// <returns>Nullable Cell Address. If no cell address could be determined, null will be returned</returns>
        /// \remark <remarks>GetFirstCellAddress() will not return the first cell with data in any case. If there is a formatted but empty cell (or many) before the first cell with data, 
        /// GetLastCellAddress() will return the address of this empty cell. Use <see cref="GetFirstDataCellAddress"/> in this case.</remarks>
        public Address? GetFirstCellAddress()
        {
            int firstRow = GetFirstRowNumber();
            int firstColumn = GetFirstColumnNumber();
            if (firstRow < 0 || firstColumn < 0)
            {
                return null;
            }
            return new Address(firstColumn, firstRow);
        }

        /// <summary>
        ///  Gets the first existing cell with data in the current worksheet (bottom right)
        /// </summary>
        /// <returns>Nullable Cell Address. If no cell address could be determined, null will be returned</returns>
        /// \remark <remarks>GetFirstDataCellAddress() will ignore formatted but empty cells before the first cell with data. 
        /// If you want the first defined cell, use <see cref="GetFirstCellAddress"/> instead.</remarks>
        public Address? GetFirstDataCellAddress()
        {
            int firstRow = GetFirstDataRowNumber();
            int firstColumn = GetFirstDataColumnNumber();
            if (firstRow < 0 || firstColumn < 0)
            {
                return null;
            }
            return new Address(firstColumn, firstRow);
        }

        /// <summary>
        /// Gets either the minimum or maximum row or column number, considering only calls with data
        /// </summary>
        /// <param name="row">If true, the min or max row is returned, otherwise the column</param>
        /// <param name="min">If true, the min value of the row or column is defined, otherwise the max value</param>
        /// <param name="ignoreEmpty">If true, empty cell values are ignored, otherwise considered without checking the content</param>
        /// <returns>Min or max number, or -1 if not defined</returns>
        private int GetBoundaryDataNumber(bool row, bool min, bool ignoreEmpty)
        {
            if (cells.Count == 0)
            {
                return -1;
            }
            if (!ignoreEmpty)
            {
                if (row && min)
                {
                    return cells.Min(x => x.Value.RowNumber);
                }
                else if (row)
                {
                    return cells.Max(x => x.Value.RowNumber);
                }
                else if (min)
                {
                    return cells.Min(x => x.Value.ColumnNumber);
                }
                else
                {
                    return cells.Max(x => x.Value.ColumnNumber);
                }
            }
            List<Cell> nonEmptyCells = cells.Values.Where(x => x.Value != null && x.Value.ToString() != string.Empty).ToList();
            if (nonEmptyCells.Count == 0)
            {
                return -1;
            }
            if (row && min)
            {
                return nonEmptyCells.Min(x => x.RowNumber);
            }
            else if (row)
            {
                return nonEmptyCells.Max(x => x.RowNumber);
            }
            else if (min)
            {
                return nonEmptyCells.Min(x => x.ColumnNumber);
            }
            else
            {
                return nonEmptyCells.Max(x => x.ColumnNumber);
            }
        }

        /// <summary>
        /// Gets either the minimum or maximum row or column number, considering all available data
        /// </summary>
        /// <param name="row">If true, the min or max row is returned, otherwise the column</param>
        /// <param name="min">If true, the min value of the row or column is defined, otherwise the max value</param>
        /// <returns>Min or max number, or -1 if not defined</returns>
        private int GetBoundaryNumber(bool row, bool min)
        {
            int cellBoundary = GetBoundaryDataNumber(row, min, false);
            if (row)
            {
                int heightBoundary = -1;
                if (rowHeights.Count > 0)
                {
                    heightBoundary = min ? RowHeights.Min(x => x.Key) : RowHeights.Max(x => x.Key);
                }
                int hiddenBoundary = -1;
                if (hiddenRows.Count > 0)
                {
                    hiddenBoundary = min ? HiddenRows.Min(x => x.Key) : HiddenRows.Max(x => x.Key);
                }
                return min ? GetMinRow(cellBoundary, heightBoundary, hiddenBoundary) : GetMaxRow(cellBoundary, heightBoundary, hiddenBoundary);
            }
            else
            {
                int columnDefBoundary = -1;
                if (columns.Count > 0)
                {
                    columnDefBoundary = min ? Columns.Min(x => x.Key) : Columns.Max(x => x.Key);
                }
                if (min)
                {
                    return cellBoundary >= 0 && cellBoundary < columnDefBoundary ? cellBoundary : columnDefBoundary;
                }
                else
                {
                    return cellBoundary >= 0 && cellBoundary > columnDefBoundary ? cellBoundary : columnDefBoundary;
                }
            }
        }

        /// <summary>
        /// Gets the maximum row coordinate either from cell data, height definitions or hidden rows
        /// </summary>
        /// <param name="cellBoundary">Row number of max cell data</param>
        /// <param name="heightBoundary">Row number of max defined row height</param>
        /// <param name="hiddenBoundary">Row number of max defined hidden row</param>
        /// <returns>Max row number or -1 if nothing valid defined</returns>
        private int GetMaxRow(int cellBoundary, int heightBoundary, int hiddenBoundary)
        {
            int highest = -1;
            if (cellBoundary >= 0)
            {
                highest = cellBoundary;
            }
            if (heightBoundary >= 0 && heightBoundary > highest)
            {
                highest = heightBoundary;
            }
            if (hiddenBoundary >= 0 && hiddenBoundary > highest)
            {
                highest = hiddenBoundary;
            }
            return highest;
        }

        /// <summary>
        /// Gets the minimum row coordinate either from cell data, height definitions or hidden rows
        /// </summary>
        /// <param name="cellBoundary">Row number of min cell data</param>
        /// <param name="heightBoundary">Row number of min defined row height</param>
        /// <param name="hiddenBoundary">Row number of min defined hidden row</param>
        /// <returns>Min row number or -1 if nothing valid defined</returns>
        private int GetMinRow(int cellBoundary, int heightBoundary, int hiddenBoundary)
        {
            int lowest = int.MaxValue;
            if (cellBoundary >= 0)
            {
                lowest = cellBoundary;
            }
            if (heightBoundary >= 0 && heightBoundary < lowest)
            {
                lowest = heightBoundary;
            }
            if (hiddenBoundary >= 0 && hiddenBoundary < lowest)
            {
                lowest = hiddenBoundary;
            }
            return lowest == int.MaxValue ? -1 : lowest;
        }
        #endregion

        #region Insert-Search-Replace

        /// <summary>
        /// Inserts 'count' rows below the specified 'rowNumber'. Existing cells are moved down by the number of new rows.
        /// The inserted, new rows inherits the style of the original cell at the defined row number.
        /// The inserted cells are empty. The values can be set later
        /// </summary>
        /// <remarks>Formulas / references are not adjusted</remarks>
        /// <param name="rowNumber">Row number below which the new row(s) will be inserted.</param>
        /// <param name="numberOfNewRows">Number of rows to insert.</param>
        public void InsertRow(int rowNumber, int numberOfNewRows)
        {
            // All cells below the first row must receive a new address (row + count);
            var upperRow = this.GetRow(rowNumber);

            // Identify all cells below the insertion point to adjust their addresses
            var cellsToChange = this.Cells.Where(c => c.Value.CellAddress2.Row > rowNumber).ToList();

            // Make a copy of the cells to be moved and then delete the original cells;
            Dictionary<string, Cell> newCells = new Dictionary<string, Cell>();
            foreach (var cell in cellsToChange)
            {
                var row = cell.Value.CellAddress2.Row;
                var col = cell.Value.CellAddress2.Column;
                Address newAddress = new Address(col, row + numberOfNewRows);

                Cell newCell = new Cell(cell.Value.Value, cell.Value.DataType, newAddress);
                if (cell.Value.CellStyle != null)
                {
                    newCell.SetStyle(cell.Value.CellStyle); // Apply the style from the "old" cell.
                }
                newCells.Add(newAddress.GetAddress(), newCell);

                // Delete the original cells since the key cannot be changed.
                this.Cells.Remove(cell.Key);
            }

            // Fill the gap with new cells, using the same style as the first row.
            foreach (Cell cell in upperRow)
            {
                for (int i = 0; i < numberOfNewRows; i++)
                {
                    Address newAddress = new Address(cell.CellAddress2.Column, cell.CellAddress2.Row + 1 + i);
                    Cell newCell = new Cell(null, Cell.CellType.EMPTY, newAddress);
                    if (cell.CellStyle != null)
                        newCell.SetStyle(cell.CellStyle);
                    this.Cells.Add(newAddress.GetAddress(), newCell);
                }
            }

            // Re-add the previous cells from the copy back with a new key.
            foreach (KeyValuePair<string, Cell> cellKeyValue in newCells)
            {
                this.Cells.Add(cellKeyValue.Key, cellKeyValue.Value);  //cell.Value is the cell incl. Style etc.
            }
        }

        /// <summary>
        /// Inserts 'count' columns right of the specified 'columnNumber'. Existing cells are moved to the right by the number of new columns.
        /// The inserted, new columns inherits the style of the original cell at the defined column number.
        /// The inserted cells are empty. The values can be set later
        /// </summary>
        /// <remarks>Formulas are not adjusted</remarks>
        /// <param name="columnNumber">Column number right which the new column(s) will be inserted.</param>
        /// <param name="numberOfNewColumns">Number of columns to insert.</param>
        public void InsertColumn(int columnNumber, int numberOfNewColumns)
        {
            var leftColumn = this.GetColumn(columnNumber);
            var cellsToChange = this.Cells.Where(c => c.Value.CellAddress2.Column > columnNumber).ToList();

            Dictionary<string, Cell> newCells = new Dictionary<string, Cell>();
            foreach (var cell in cellsToChange)
            {
                var row = cell.Value.CellAddress2.Row;
                var col = cell.Value.CellAddress2.Column;
                Address newAddress = new Address(col + numberOfNewColumns, row);

                Cell newCell = new Cell(cell.Value.Value, cell.Value.DataType, newAddress);
                if (cell.Value.CellStyle != null)
                {
                    newCell.SetStyle(cell.Value.CellStyle); // Apply the style from the "old" cell.
                }
                newCells.Add(newAddress.GetAddress(), newCell);

                // Delete the original cells since the key cannot be changed.
                this.Cells.Remove(cell.Key);
            }

            // Fill the gap with new cells, using the same style as the first row.
            foreach (Cell cell in leftColumn)
            {
                for (int i = 0; i < numberOfNewColumns; i++)
                {
                    Address newAddress = new Address(cell.CellAddress2.Column + 1 + i, cell.CellAddress2.Row);
                    Cell newCell = new Cell(null, Cell.CellType.EMPTY, newAddress);
                    if (cell.CellStyle != null)
                        newCell.SetStyle(cell.CellStyle);
                    this.Cells.Add(newAddress.GetAddress(), newCell);
                }
            }

            // Re-add the previous cells from the copy back with a new key.
            foreach (KeyValuePair<string, Cell> cellKeyValue in newCells)
            {
                this.Cells.Add(cellKeyValue.Key, cellKeyValue.Value);  //cell.Value is the cell incl. Style etc.
            }
        }

        /// <summary>
        /// Searches for the first occurrence of the value.
        /// </summary>
        /// <param name="searchValue">The value to search for.</param>
        /// <returns>The first cell containing the searched value or null if the value was not found</returns>
        public Cell FirstCellByValue(object searchValue)
        {
            var cell = this.Cells.FirstOrDefault(c =>
                Equals(c.Value.Value, searchValue))
                .Value;
            return cell;
        }

        /// <summary>
        /// Searches for the first occurrence of the expression.
        /// Example: var cell = worksheet.FindCell(c => c.Value?.ToString().Contains("searchValue"));
        /// </summary>
        /// <param name="predicate"></param>
        /// <returns>The first cell containing the searched value or null if the value was not found</returns>
        public Cell FirstOrDefaultCell(Func<Cell, bool> predicate)
        {
            return this.Cells.Values
                .FirstOrDefault(c => c != null && (c.Value == null || predicate(c)));
        }

        /// <summary>
        /// Searches for cells that contain the specified value and returns a list of these cells.
        /// </summary>
        /// <param name="searchValue">The value to search for.</param>
        /// <returns>A list of cells that contain the specified value.</returns>
        public List<Cell> CellsByValue(object searchValue)
        {
            return this.Cells.Where(c =>
                Equals(c.Value.Value, searchValue))
                .Select(c => c.Value)
                .ToList();
        }

        /// <summary>
        /// Replaces all occurrences of 'oldValue' with 'newValue' and returns the number of replacements.
        /// </summary>
        /// <param name="oldValue">Old value</param>
        /// <param name="newValue">New value that should replace the old one</param>
        /// <returns>Count of replaced Cell values</returns>
        public int ReplaceCellValue(object oldValue, object newValue)
        {
            int count = 0;
            List<Cell> foundCells = this.CellsByValue(oldValue);
            foreach (var cell in foundCells)
            {
                cell.Value = newValue;
                count++;
            }
            return count;
        }
        #endregion

        #region common_methods

        /// <summary>
        /// Method to add allowed actions if the worksheet is protected. If one or more values are added, UseSheetProtection will be set to true
        /// </summary>
        /// <param name="typeOfProtection">Allowed action on the worksheet or cells</param>
        /// \remark <remarks>If <see cref="SheetProtectionValue.selectLockedCells"/> is added, <see cref="SheetProtectionValue.selectUnlockedCells"/> is added automatically</remarks>
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
        /// <exception cref="RangeException">Throws a RangeException if the passed column address is out of range</exception>
        public void AddHiddenColumn(string columnAddress)
        {
            int columnNumber = Cell.ResolveColumn(columnAddress);
            SetColumnHiddenState(columnNumber, true);
        }

        /// <summary>
        /// Sets the defined row as hidden
        /// </summary>
        /// <param name="rowNumber">Row number to hide on the worksheet</param>
        /// <exception cref="RangeException">Throws a RangeException if the passed row number is out of range</exception>
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
                throw new WorksheetException("The cell with the address " + address.GetAddress() + " does not exist in this worksheet");
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
        /// <exception cref="RangeException">A RangeException is thrown if the column or row number is invalid</exception>
        public bool HasCell(int columnNumber, int rowNumber)
        {
            return HasCell(new Address(columnNumber, rowNumber));
        }

        /// <summary>
        /// Resets the defined column, if existing. The corresponding instance will be removed from <see cref="Columns"/>.
        /// </summary>
        /// \remark <remarks>If the column is inside an autoFilter-Range, the column cannot be entirely removed from <see cref="Columns"/>. The hidden state will be set to false and width to default, in this case.</remarks>
        /// <param name="columnNumber">Column number to reset (zero-based)</param>
        public void ResetColumn(int columnNumber)
        {
            if (columns.ContainsKey(columnNumber) && !columns[columnNumber].HasAutoFilter) // AutoFilters cannot have gaps 
            {
                columns.Remove(columnNumber);
            }
            else if (columns.ContainsKey(columnNumber))
            {
                columns[columnNumber].IsHidden = false;
                columns[columnNumber].Width = DEFAULT_COLUMN_WIDTH;
            }
        }

        /// <summary>
        /// Gets a row as list of cell objects
        /// </summary>
        /// <param name="rowNumber">Row number (zero-based)</param>
        /// <returns>List of cell objects. If the row doesn't exist, an empty list is returned</returns>
        public IReadOnlyList<Cell> GetRow(int rowNumber)
        {
            List<Cell> list = new List<Cell>();
            foreach (KeyValuePair<string, Cell> cell in cells)
            {
                if (cell.Value.RowNumber == rowNumber)
                {
                    list.Add(cell.Value);
                }
            }
            list.Sort((c1, c2) => (c1.ColumnNumber.CompareTo(c2.ColumnNumber))); // Lambda sort
            return list;
        }

        /// <summary>
        /// Gets a column as list of cell objects
        /// </summary>
        /// <param name="columnAddress">Column address</param>
        /// <exception cref="RangeException">A range exception is thrown if the address is not valid</exception>
        /// <returns>List of cell objects. If the column doesn't exist, an empty list is returned</returns>
        public IReadOnlyList<Cell> GetColumn(string columnAddress)
        {
            int column = Cell.ResolveColumn(columnAddress);
            return GetColumn(column);
        }

        /// <summary>
        /// Gets a column as list of cell objects
        /// </summary>
        /// <param name="columnNumber">Column number (zero-based)</param>
        /// <returns>List of cell objects. If the column doesn't exist, an empty list is returned</returns>
        public IReadOnlyList<Cell> GetColumn(int columnNumber)
        {
            List<Cell> list = new List<Cell>();
            foreach (KeyValuePair<string, Cell> cell in cells)
            {
                if (cell.Value.ColumnNumber == columnNumber)
                {
                    list.Add(cell.Value);
                }
            }
            list.Sort((c1, c2) => (c1.RowNumber.CompareTo(c2.RowNumber))); // Lambda sort
            return list;
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
            Cell.ValidateColumnNumber(currentColumnNumber);
        }

        /// <summary>
        /// Moves the current position to the next column with the number of cells to move
        /// </summary>
        /// <param name="numberOfColumns">Number of columns to move</param>
        /// <param name="keepRowPosition">If true, the row position is preserved, otherwise set to 0</param>
        /// \remark <remarks>The value can also be negative. However, resulting column numbers below 0 or above 16383 will cause an exception</remarks>
        public void GoToNextColumn(int numberOfColumns, bool keepRowPosition = false)
        {
            currentColumnNumber += numberOfColumns;
            if (!keepRowPosition)
            {
                currentRowNumber = 0;
            }
            Cell.ValidateColumnNumber(currentColumnNumber);
        }

        /// <summary>
        /// Moves the current position to the next row (use for a new line)
        /// </summary>
        public void GoToNextRow()
        {
            currentRowNumber++;
            currentColumnNumber = 0;
            Cell.ValidateRowNumber(currentRowNumber);
        }

        /// <summary>
        /// Moves the current position to the next row with the number of cells to move (use for a new line)
        /// </summary>
        /// <param name="numberOfRows">Number of rows to move</param>
        /// <param name="keepColumnPosition">If true, the column position is preserved, otherwise set to 0</param>
        /// \remark <remarks>The value can also be negative. However, resulting row numbers below 0 or above 1048575 will cause an exception</remarks>
        public void GoToNextRow(int numberOfRows, bool keepColumnPosition = false)
        {
            currentRowNumber += numberOfRows;
            if (!keepColumnPosition)
            {
                currentColumnNumber = 0;
            }
            Cell.ValidateRowNumber(currentRowNumber);
        }

        /// <summary>
        /// Merges the defined cell range
        /// </summary>
        /// <param name="cellRange">Range to merge</param>
        /// <returns>Returns the validated range of the merged cells (e.g. 'A1:B12')</returns>
        /// <exception cref="RangeException">Throws a RangeException if the passed cell range is out of range</exception>
        public string MergeCells(Range cellRange)
        {
            return MergeCells(cellRange.StartAddress, cellRange.EndAddress);
        }

        /// <summary>
        /// Merges the defined cell range
        /// </summary>
        /// <param name="cellRange">Range to merge (e.g. 'A1:B12')</param>
        /// <returns>Returns the validated range of the merged cells (e.g. 'A1:B12')</returns>
        /// <exception cref="RangeException">Throws a RangeException if the passed cell range is out of range</exception>
        /// <exception cref="NanoXLSX.Exceptions.FormatException">Throws a FormatException if the passed cell range is malformed</exception>
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
        /// <exception cref="RangeException">Throws a RangeException if one of the passed cell addresses is out of range or if one or more cell addresses are already occupied in another merge range</exception>
        public string MergeCells(Address startAddress, Address endAddress)
        {
            string key = startAddress + ":" + endAddress;
            Range value = new Range(startAddress, endAddress);
            IReadOnlyList<Address> result = value.ResolveEnclosedAddresses();
            foreach (KeyValuePair<string, Range> item in mergedCells)
            {
                if (item.Value.ResolveEnclosedAddresses().Intersect(result).Any())
                {
                    throw new RangeException("The passed range: " + value.ToString() + " contains cells that are already in the defined merge range: " + item.Key);
                }
            }
            mergedCells.Add(key, value);
            return key;
        }

        /// <summary>
        /// Method to recalculate the auto filter (columns) of this worksheet. This is an internal method. There is no need to use it
        /// </summary>
        internal void RecalculateAutoFilter()
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
            autoFilterRange = new Range(start, 0, end, endRow);
        }

        /// <summary>
        /// Method to recalculate the collection of columns of this worksheet. This is an internal method. There is no need to use it
        /// </summary>
        internal void RecalculateColumns()
        {
            List<int> columnsToDelete = new List<int>();
            foreach (KeyValuePair<int, Column> col in columns)
            {
                if (!col.Value.HasAutoFilter && !col.Value.IsHidden && Comparators.CompareDimensions(col.Value.Width, DEFAULT_COLUMN_WIDTH) == 0 && col.Value.DefaultColumnStyle == null)
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
        /// Method to resolve all merged cells of the worksheet. Only the value of the very first cell of the locked cells range will be visible. The other values are still present (set to EMPTY) but will not be stored in the worksheet.<br />
        /// This is an internal method. There is no need to use it
        /// </summary>
        /// <exception cref="StyleException">Throws a StyleException if one of the styles of the merged cells cannot be referenced or is null</exception>
        internal void ResolveMergedCells()
        {
            Style mergeStyle = BasicStyles.MergeCellStyle;
            Cell cell;
            foreach (KeyValuePair<string, Range> range in MergedCells)
            {
                int pos = 0;
                List<Address> addresses = Cell.GetCellRange(range.Value.StartAddress, range.Value.EndAddress) as List<Address>;
                foreach (Address address in addresses)
                {
                    if (!Cells.ContainsKey(address.GetAddress()))
                    {
                        cell = new Cell();
                        cell.DataType = Cell.CellType.EMPTY;
                        cell.RowNumber = address.Row;
                        cell.ColumnNumber = address.Column;
                        AddCell(cell, cell.ColumnNumber, cell.RowNumber);
                    }
                    else
                    {
                        cell = Cells[address.GetAddress()];
                    }
                    if (pos != 0)
                    {
                        cell.DataType = Cell.CellType.EMPTY;
                        if (cell.CellStyle == null)
                        {
                            cell.SetStyle(mergeStyle);
                        }
                        else
                        {
                            Style mixedMergeStyle = cell.CellStyle;
                            // TODO: There should be a better possibility to identify particular style elements that deviates
                            mixedMergeStyle.CurrentCellXf.ForceApplyAlignment = mergeStyle.CurrentCellXf.ForceApplyAlignment;
                            cell.SetStyle(mixedMergeStyle);
                        }
                    }
                    pos++;
                }
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
        /// <exception cref="RangeException">Throws a RangeException if the passed column number is out of range</exception>
        public void RemoveHiddenColumn(int columnNumber)
        {
            SetColumnHiddenState(columnNumber, false);
        }

        /// <summary>
        /// Sets a previously defined, hidden column as visible again
        /// </summary>
        /// <param name="columnAddress">Column address to make visible again</param>
        /// <exception cref="RangeException">Throws a RangeException if the column address out of range</exception>
        public void RemoveHiddenColumn(string columnAddress)
        {
            int columnNumber = Cell.ResolveColumn(columnAddress);
            SetColumnHiddenState(columnNumber, false);
        }

        /// <summary>
        /// Sets a previously defined, hidden row as visible again
        /// </summary>
        /// <param name="rowNumber">Row number to hide on the worksheet</param>
        /// <exception cref="RangeException">Throws a RangeException if the passed row number is out of range</exception>
        public void RemoveHiddenRow(int rowNumber)
        {
            SetRowHiddenState(rowNumber, false);
        }

        /// <summary>
        /// Removes the defined merged cell range
        /// </summary>
        /// <param name="range">Cell range to remove the merging</param>
        /// <exception cref="RangeException">Throws a RangeException if the passed cell range was not merged earlier</exception>
        public void RemoveMergedCells(string range)
        {
            range = ParserUtils.ToUpper(range);
            if (range == null || !mergedCells.ContainsKey(range))
            {
                throw new RangeException("The cell range " + range + " was not found in the list of merged cell ranges");
            }

            List<Address> addresses = Cell.GetCellRange(range) as List<Address>;
            foreach (Address address in addresses)
            {
                if (cells.ContainsKey(address.GetAddress()))
                {
                    Cell cell = cells[address.GetAddress()];
                    if (BasicStyles.MergeCellStyle.Equals(cell.CellStyle))
                    {
                        cell.RemoveStyle();
                    }
                    cell.ResolveCellType(); // resets the type
                }
            }
            mergedCells.Remove(range);
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
            if (sheetProtectionValues.Contains(value))
            {
                sheetProtectionValues.Remove(value);
            }
        }

        /// <summary>
        /// Sets the active style of the worksheet. This style will be assigned to all later added cells
        /// </summary>
        /// <param name="style">Style to set as active style</param>
        public void SetActiveStyle(Style style)
        {
            if (style == null)
            {
                useActiveStyle = false;
            }
            else
            {
                useActiveStyle = true;
            }
            activeStyle = style;
        }

        /// <summary>
        /// Sets the column auto filter within the defined column range
        /// </summary>
        /// <param name="startColumn">Column number with the first appearance of an auto filter drop down</param>
        /// <param name="endColumn">Column number with the last appearance of an auto filter drop down</param>
        /// <exception cref="RangeException">Throws a RangeException if the start or end address out of range</exception>
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
        /// <exception cref="RangeException">Throws a RangeException if the passed range out of range</exception>
        /// <exception cref="NanoXLSX.Exceptions.FormatException">Throws a FormatException if the passed range is malformed</exception>
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
        /// <exception cref="RangeException">Throws a RangeException if the column number out of range</exception>
        private void SetColumnHiddenState(int columnNumber, bool state)
        {
            Cell.ValidateColumnNumber(columnNumber);
            if (columns.ContainsKey(columnNumber))
            {
                columns[columnNumber].IsHidden = state;
            }
            else if (state)
            {
                Column c = new Column(columnNumber);
                c.IsHidden = true;
                columns.Add(columnNumber, c);
            }
            if (!columns[columnNumber].IsHidden && Comparators.CompareDimensions(columns[columnNumber].Width, DEFAULT_COLUMN_WIDTH) == 0 && !columns[columnNumber].HasAutoFilter)
            {
                columns.Remove(columnNumber);
            }
        }

        /// <summary>
        /// Sets the width of the passed column address
        /// </summary>
        /// <param name="columnAddress">Column address (A - XFD)</param>
        /// <param name="width">Width from 0 to 255.0</param>
        /// <exception cref="RangeException">Throws a RangeException:<br />a) If the passed column address is out of range<br />b) if the column width is out of range (0 - 255.0)</exception>
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
        /// <exception cref="RangeException">Throws a RangeException:<br />a) If the passed column number is out of range<br />b) if the column width is out of range (0 - 255.0)</exception>
        public void SetColumnWidth(int columnNumber, float width)
        {
            Cell.ValidateColumnNumber(columnNumber);
            if (width < MIN_COLUMN_WIDTH || width > MAX_COLUMN_WIDTH)
            {
                throw new RangeException("The column width (" + width + ") is out of range. Range is from " + MIN_COLUMN_WIDTH + " to " + MAX_COLUMN_WIDTH + " (chars).");
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
        /// Sets the default column style of the passed column address
        /// </summary>
        /// <param name="columnAddress">Column address (A - XFD)</param>
        /// <param name="style">Style to set as default. If null, the style is cleared</param>
        /// <returns>Assigned style or null if cleared</returns>
        /// <exception cref="RangeException">Throws a RangeException:<br />a) If the passed column address is out of range<br />b) if the column width is out of range (0 - 255.0)</exception>
        public Style SetColumnDefaultStyle(string columnAddress, Style style)
        {
            int columnNumber = Cell.ResolveColumn(columnAddress);
            return SetColumnDefaultStyle(columnNumber, style);
        }
        /// <summary>
        /// Sets the default column style of the passed column number (zero-based)
        /// </summary>
        /// <param name="columnNumber">Column number (zero-based, from 0 to 16383)</param>
        /// <param name="style">Style to set as default. If null, the style is cleared</param>
        /// <returns>Assigned style or null if cleared</returns>
        /// <exception cref="RangeException">Throws a RangeException:<br />a) If the passed column number is out of range<br />b) if the column width is out of range (0 - 255.0)</exception>
        public Style SetColumnDefaultStyle(int columnNumber, Style style)
        {
            Cell.ValidateColumnNumber(columnNumber);
            if (this.columns.ContainsKey(columnNumber))
            {
                return this.columns[columnNumber].SetDefaultColumnStyle(style);
            }
            else
            {
                Column c = new Column(columnNumber);
                Style returnStyle = c.SetDefaultColumnStyle(style);
                this.columns.Add(columnNumber, c);
                return returnStyle;
            }
        }

        /// <summary>
        /// Set the current cell address
        /// </summary>
        /// <param name="columnNumber">Column number (zero based)</param>
        /// <param name="rowNumber">Row number (zero based)</param>
        /// <exception cref="RangeException">Throws a RangeException if one of the passed cell addresses is out of range</exception>
        public void SetCurrentCellAddress(int columnNumber, int rowNumber)
        {
            SetCurrentColumnNumber(columnNumber);
            SetCurrentRowNumber(rowNumber);
        }

        /// <summary>
        /// Set the current cell address
        /// </summary>
        /// <param name="address">Cell address in the format A1 - XFD1048576</param>
        /// <exception cref="RangeException">Throws a RangeException if the passed cell address is out of range</exception>
        /// <exception cref="NanoXLSX.Exceptions.FormatException">Throws a FormatException if the passed cell address is malformed</exception>
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
        /// <exception cref="RangeException">Throws a RangeException if the number is out of the valid range. Range is from 0 to 16383 (16384 columns)</exception>
        public void SetCurrentColumnNumber(int columnNumber)
        {
            Cell.ValidateColumnNumber(columnNumber);
            currentColumnNumber = columnNumber;
        }

        /// <summary>
        /// Sets the current row number (zero based)
        /// </summary>
        /// <param name="rowNumber">Row number (zero based)</param>
        /// <exception cref="RangeException">Throws a RangeException if the number is out of the valid range. Range is from 0 to 1048575 (1048576 rows)</exception>
        public void SetCurrentRowNumber(int rowNumber)
        {
            Cell.ValidateRowNumber(rowNumber);
            currentRowNumber = rowNumber;
        }

        /// <summary>
        /// Adds a range to the selected cells on this worksheet
        /// </summary>
        /// <param name="range">Cell range to add</param>
        public void AddSelectedCells(Range range)
        {
            selectedCells = DataUtils.MergeRange(selectedCells, range).ToList();
        }

        /// <summary>
        /// Adds a range to the selected cells on this worksheet
        /// </summary>
        /// <param name="startAddress">Start address of the range</param>
        /// <param name="endAddress">End address of the range</param>
        public void AddSelectedCells(Address startAddress, Address endAddress)
        {
            AddSelectedCells(new Range(startAddress, endAddress));
        }

        /// <summary>
        /// Adds a range or cell address to the selected cells on this worksheet
        /// </summary>
        /// <param name="rangeOrAddress">Cell range or address to add</param>
        public void AddSelectedCells(string rangeOrAddress)
        {
            Range? resolved = ParseRange(rangeOrAddress);
            if (resolved != null)
            {
                AddSelectedCells(resolved.Value);
            }
        }

        /// <summary>
        /// Adds a single cell address to the selected cells on this worksheet
        /// </summary>
        /// <param name="address">Cell address to add</param>
        public void AddSelectedCells(Address address)
        {
            AddSelectedCells(new Range(address, address));
        }

        /// <summary>
        /// Removes all cell selections of this worksheet
        /// </summary>
        public void ClearSelectedCells()
        {
            selectedCells.Clear();
        }

        /// <summary>
        /// Removes the given range from the selected cell ranges of this worksheet, if existing.
        /// If the passed range is overlapping the ranges of the selected cells, only the intersecting addresses will be removed
        /// </summary>
        /// <param name="range">Range to remove</param>
        public void RemoveSelectedCells(Range range)
        {
            selectedCells = DataUtils.SubtractRange(selectedCells, range).ToList();
        }

        /// <summary>
        /// Removes the given range or cell address from the selected cell ranges of this worksheet, if existing
        /// </summary>
        /// <param name="rangeOrAddress">Range or cell address to remove</param>
        public void RemoveSelectedCells(String rangeOrAddress)
        {
            Range? resolved = ParseRange(rangeOrAddress);
            if (resolved != null)
            {
                RemoveSelectedCells(resolved.Value);
            }
        }

        /// <summary>
        /// Removes the given range from the selected cell ranges of this worksheet, if existing
        /// </summary>
        /// <param name="startAddress">Start address of the range to remove</param>
        /// <param name="endAddress">End address of the range to remove</param>
        public void RemoveSelectedCells(Address startAddress, Address endAddress)
        {
            RemoveSelectedCells(new Range(startAddress, endAddress));
        }

        /// <summary>
        /// Removes the given address from the selected cell ranges of this worksheet, if existing
        /// </summary>
        /// <param name="address">Address of the range to remove</param>
        public void RemoveSelectedCells(Address address)
        {
            RemoveSelectedCells(new Range(address, address));
        }

        /// <summary>
        /// Sets or removes the password for worksheet protection. If set, UseSheetProtection will be also set to true
        /// </summary>
        /// <param name="password">Password (UTF-8) to protect the worksheet. If the password is null or empty, no password will be used</param>
        public void SetSheetProtectionPassword(string password)
        {
            if (string.IsNullOrEmpty(password))
            {
                sheetProtectionPassword.UnsetPassword();
                UseSheetProtection = false;
            }
            else
            {
                sheetProtectionPassword.SetPassword(password);
                UseSheetProtection = true;
            }
        }

        /// <summary>
        /// Sets the height of the passed row number (zero-based)
        /// </summary>
        /// <param name="rowNumber">Row number (zero-based, 0 to 1048575)</param>
        /// <param name="height">Height from 0 to 409.5</param>
        /// <exception cref="RangeException">Throws a RangeException:<br />a) If the passed row number is out of range<br />b) if the row height is out of range (0 - 409.5)</exception>
        public void SetRowHeight(int rowNumber, float height)
        {
            Cell.ValidateRowNumber(rowNumber);
            if (height < MIN_ROW_HEIGHT || height > MAX_ROW_HEIGHT)
            {
                throw new RangeException("The row height (" + height + ") is out of range. Range is from " + MIN_ROW_HEIGHT + " to " + MAX_ROW_HEIGHT + " (equals 546px).");
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
        /// <exception cref="RangeException">Throws a RangeException if the passed row number was out of range</exception>
        private void SetRowHiddenState(int rowNumber, bool state)
        {
            Cell.ValidateRowNumber(rowNumber);
            if (hiddenRows.ContainsKey(rowNumber))
            {
                if (state)
                {
                    hiddenRows[rowNumber] = true;
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
        }

        /// <summary>
        /// Validates and sets the worksheet name
        /// </summary>
        /// <param name="name">Name to set</param>
        /// <exception cref="NanoXLSX.Exceptions.FormatException">Throws a FormatException if the worksheet name is too long (max. 31) or contains illegal characters [  ]  * ? / \</exception>
        public void SetSheetName(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                throw new FormatException("the worksheet name must be between 1 and " + MAX_WORKSHEET_NAME_LENGTH + " characters");
            }
            if (name.Length > MAX_WORKSHEET_NAME_LENGTH)
            {
                throw new FormatException("the worksheet name must be between 1 and " + MAX_WORKSHEET_NAME_LENGTH + " characters");
            }
            Regex regex = new Regex(@"[\[\]\*\?/\\]");
            Match match = regex.Match(name);
            if (match.Captures.Count > 0)
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
            if (sanitize)
            {
                sheetName = ""; // Empty name (temporary) to prevent conflicts during sanitizing
                sheetName = SanitizeWorksheetName(name, workbookReference);
            }
            else
            {
                SetSheetName(name);
            }
        }

        /// <summary>
        /// Sets the horizontal split of the worksheet into two panes. The measurement in characters cannot be used to freeze panes
        /// </summary>
        /// <param name="topPaneHeight">Height (similar to row height) from top of the worksheet to the split line in characters</param>
        /// <param name="topLeftCell">Top Left cell address of the bottom right pane (if applicable). Only the row component is important in a horizontal split</param>
        /// <param name="activePane">Active pane in the split window.<br />The parameter is nullable</param>
        public void SetHorizontalSplit(float topPaneHeight, Address topLeftCell, WorksheetPane? activePane)
        {
            SetSplit(null, topPaneHeight, topLeftCell, activePane);
        }

        /// <summary>
        /// Sets the horizontal split of the worksheet into two panes. The measurement in rows can be used to split and freeze panes
        /// </summary>
        /// <param name="numberOfRowsFromTop">Number of rows from top of the worksheet to the split line. The particular row heights are considered</param>
        /// <param name="freeze">If true, all panes are frozen, otherwise remains movable</param>
        /// <param name="topLeftCell">Top Left cell address of the bottom right pane (if applicable). Only the row component is important in a horizontal split</param>
        /// <param name="activePane">Active pane in the split window.<br />The parameter is nullable</param>
        /// <exception cref="WorksheetException">WorksheetException Thrown if the row number of the top left cell is smaller the split panes number of rows from top, if freeze is applied</exception>
        public void SetHorizontalSplit(int numberOfRowsFromTop, bool freeze, Address topLeftCell, WorksheetPane? activePane)
        {
            SetSplit(null, numberOfRowsFromTop, freeze, topLeftCell, activePane);
        }

        /// <summary>
        /// Sets the vertical split of the worksheet into two panes. The measurement in characters cannot be used to freeze panes
        /// </summary>
        /// <param name="leftPaneWidth">Width (similar to column width) from left of the worksheet to the split line in characters</param>
        /// <param name="topLeftCell">Top Left cell address of the bottom right pane (if applicable). Only the column component is important in a vertical split</param>
        /// <param name="activePane">Active pane in the split window.<br />The parameter is nullable</param>
        public void SetVerticalSplit(float leftPaneWidth, Address topLeftCell, WorksheetPane? activePane)
        {
            SetSplit(leftPaneWidth, null, topLeftCell, activePane);
        }

        /// <summary>
        /// Sets the vertical split of the worksheet into two panes. The measurement in columns can be used to split and freeze panes
        /// </summary>
        /// <param name="numberOfColumnsFromLeft">Number of columns from left of the worksheet to the split line. The particular column widths are considered</param>
        /// <param name="freeze">If true, all panes are frozen, otherwise remains movable</param>
        /// <param name="topLeftCell">Top Left cell address of the bottom right pane (if applicable). Only the column component is important in a vertical split</param>
        /// <param name="activePane">Active pane in the split window.<br />The parameter is nullable</param>
        /// <exception cref="WorksheetException">WorksheetException Thrown if the column number of the top left cell is smaller the split panes number of columns from left, 
        /// if freeze is applied</exception>
        public void SetVerticalSplit(int numberOfColumnsFromLeft, bool freeze, Address topLeftCell, WorksheetPane? activePane)
        {
            SetSplit(numberOfColumnsFromLeft, null, freeze, topLeftCell, activePane);
        }

        /// <summary>
        /// Sets the horizontal and vertical split of the worksheet into four panes. The measurement in rows and columns can be used to split and freeze panes
        /// </summary>
        /// <param name="numberOfColumnsFromLeft">Number of columns from left of the worksheet to the split line. The particular column widths are considered.<br />
        /// The parameter is nullable. If left null, the method acts identical to <see cref="SetHorizontalSplit(int, bool, Address, WorksheetPane?)"/></param>
        /// <param name="numberOfRowsFromTop">Number of rows from top of the worksheet to the split line. The particular row heights are considered.<br />
        /// The parameter is nullable. If left null, the method acts identical to <see cref="SetVerticalSplit(int, bool, Address, WorksheetPane?)"/></param>
        /// <param name="freeze">If true, all panes are frozen, otherwise remains movable</param>
        /// <param name="topLeftCell">Top Left cell address of the bottom right pane (if applicable)</param>
        /// <param name="activePane">Active pane in the split window.<br />The parameter is nullable</param>
        /// <exception cref="WorksheetException">WorksheetException Thrown if the address of the top left cell is smaller the split panes address, if freeze is applied</exception>
        public void SetSplit(int? numberOfColumnsFromLeft, int? numberOfRowsFromTop, bool freeze, Address topLeftCell, WorksheetPane? activePane)
        {
            if (freeze)
            {
                if (numberOfColumnsFromLeft != null && topLeftCell.Column < numberOfColumnsFromLeft.Value)
                {
                    throw new WorksheetException("The column number " + topLeftCell.Column +
                        " is not valid for a frozen, vertical split with the split pane column number " + numberOfColumnsFromLeft.Value);
                }
                if (numberOfRowsFromTop != null && topLeftCell.Row < numberOfRowsFromTop.Value)
                {
                    throw new WorksheetException("The row number " + topLeftCell.Row +
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
        /// <param name="leftPaneWidth">Width (similar to column width) from left of the worksheet to the split line in characters.<br />
        /// The parameter is nullable. If left null, the method acts identical to <see cref="SetHorizontalSplit(float, Address, WorksheetPane?)"/></param>
        /// <param name="topPaneHeight">Height (similar to row height) from top of the worksheet to the split line in characters.<br />
        /// The parameter is nullable. If left null, the method acts identical to <see cref="SetVerticalSplit(float, Address, WorksheetPane?)"/></param>
        /// <param name="topLeftCell">Top Left cell address of the bottom right pane (if applicable)</param>
        /// <param name="activePane">Active pane in the split window.<br />The parameter is nullable</param>
        public void SetSplit(float? leftPaneWidth, float? topPaneHeight, Address topLeftCell, WorksheetPane? activePane)
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

        /// <summary>
        /// Creates a (dereferenced) deep copy of this worksheet
        /// </summary>
        /// \remark <remarks>Not considered in the copy are the internal ID, the worksheet name and the workbook reference. 
        /// Since styles are managed in a shared repository, no dereferencing is applied (Styles are not deep-copied). 
        /// Use <see cref="Workbook.CopyWorksheetTo(Worksheet, string, Workbook, bool)"/> or <see cref="Workbook.CopyWorksheetIntoThis(Worksheet, string, bool)"/> 
        /// to add a copy of worksheet to a workbook. These methods will set the internal ID, name and workbook reference.
        /// </remarks>
        /// <returns>Copy of this worksheet</returns>
        public Worksheet Copy()
        {
            Worksheet copy = new Worksheet();
            foreach (KeyValuePair<string, Cell> cell in this.cells)
            {
                copy.AddCell(cell.Value.Copy(), cell.Key);
            }
            copy.activePane = this.activePane;
            copy.activeStyle = this.activeStyle;
            if (this.autoFilterRange.HasValue)
            {
                copy.autoFilterRange = this.autoFilterRange.Value.Copy();
            }
            foreach (KeyValuePair<int, Column> column in this.columns)
            {
                copy.columns.Add(column.Key, column.Value.Copy());
            }
            copy.CurrentCellDirection = this.CurrentCellDirection;
            copy.currentColumnNumber = this.currentColumnNumber;
            copy.currentRowNumber = this.currentRowNumber;
            copy.defaultColumnWidth = this.defaultColumnWidth;
            copy.defaultRowHeight = this.defaultRowHeight;
            copy.freezeSplitPanes = this.freezeSplitPanes;
            copy.hidden = this.hidden;
            foreach (KeyValuePair<int, bool> row in this.hiddenRows)
            {
                copy.hiddenRows.Add(row.Key, row.Value);
            }
            foreach (KeyValuePair<string, Range> cell in this.mergedCells)
            {
                copy.mergedCells.Add(cell.Key, cell.Value.Copy());
            }
            if (this.paneSplitAddress.HasValue)
            {
                copy.paneSplitAddress = this.paneSplitAddress.Value.Copy();
            }
            copy.paneSplitLeftWidth = this.paneSplitLeftWidth;
            copy.paneSplitTopHeight = this.paneSplitTopHeight;
            if (this.paneSplitTopLeftCell.HasValue)
            {
                copy.paneSplitTopLeftCell = this.paneSplitTopLeftCell.Value.Copy();
            }
            foreach (KeyValuePair<int, float> row in this.rowHeights)
            {
                copy.rowHeights.Add(row.Key, row.Value);
            }
            foreach (Range range in selectedCells)
            {
                copy.AddSelectedCells(range);
            }
            copy.sheetProtectionPassword.CopyFrom(this.sheetProtectionPassword);
            foreach (SheetProtectionValue value in this.sheetProtectionValues)
            {
                copy.sheetProtectionValues.Add(value);
            }
            copy.useActiveStyle = this.useActiveStyle;
            copy.UseSheetProtection = this.UseSheetProtection;
            copy.ShowGridLines = this.ShowGridLines;
            copy.ShowRowColumnHeaders = this.ShowRowColumnHeaders;
            copy.ShowRuler = this.ShowRuler;
            copy.ViewType = this.ViewType;
            copy.zoomFactor.Clear();
            foreach (KeyValuePair<SheetViewType, int> zoomFactor in this.zoomFactor)
            {
                copy.SetZoomFactor(zoomFactor.Key, zoomFactor.Value);
            }
            return copy;
        }

        /// <summary>
        /// Sets a zoom factor for a given <see cref="SheetViewType"/>. If <see cref="AUTO_ZOOM_FACTOR"/>, the zoom factor is set to automatic
        /// </summary>
        /// <param name="sheetViewType">Sheet view type to apply the zoom factor on</param>
        /// <param name="zoomFactor">Zoom factor in percent</param>
        /// \remark <remarks>This factor is not the currently set factor. use the property <see cref="ZoomFactor"/> to set the factor for the current <see cref="ViewType"/></remarks>
        /// <exception cref="WorksheetException">Throws a WorksheetException if the zoom factor is not <see cref="AUTO_ZOOM_FACTOR"/> or below <see cref="MIN_ZOOM_FACTOR"/> or above <see cref="MAX_ZOOM_FACTOR"/></exception>
        public void SetZoomFactor(SheetViewType sheetViewType, int zoomFactor)
        {
            if (zoomFactor != AUTO_ZOOM_FACTOR && (zoomFactor < MIN_ZOOM_FACTOR || zoomFactor > MAX_ZOOM_FACTOR))
            {
                throw new WorksheetException("The zoom factor " + zoomFactor + " is not valid. Valid are values between " + MIN_ZOOM_FACTOR + " and " + MAX_ZOOM_FACTOR + ", or " + AUTO_ZOOM_FACTOR + " (automatic)");
            }
            if (this.zoomFactor.ContainsKey(sheetViewType))
            {
                this.zoomFactor[sheetViewType] = zoomFactor;
            }
            else
            {
                this.zoomFactor.Add(sheetViewType, zoomFactor);
            }
        }



        #region static_methods
        /// <summary>
        /// Sanitizes a worksheet name
        /// </summary>
        /// <param name="input">Name to sanitize</param>
        /// <param name="workbook">Workbook reference</param>
        /// <exception cref="WorksheetException">A WorksheetException is thrown if the workbook reference is null, since all worksheets have to be considered during sanitation</exception>
        /// <returns>Name of the sanitized worksheet</returns>
        public static string SanitizeWorksheetName(string input, Workbook workbook)
        {
            if (string.IsNullOrEmpty(input))
            {
                input = "Sheet1";
            }
            int len;
            if (input.Length > MAX_WORKSHEET_NAME_LENGTH)
            {
                len = MAX_WORKSHEET_NAME_LENGTH;
            }
            else
            {
                len = input.Length;
            }
            StringBuilder sb = new StringBuilder(MAX_WORKSHEET_NAME_LENGTH);
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
        /// Parses a string to a range. If the string is a single address, the range consists of this as start and end address
        /// </summary>
        /// <param name="rangeOrAddress">Range or address expression</param>
        /// <returns>Range or null if the </returns>
        private static Range? ParseRange(string rangeOrAddress)
        {
            if (string.IsNullOrEmpty(rangeOrAddress))
            {
                return null;
            }
            Range range;
            if (rangeOrAddress.Contains(":"))
            {
                range = Cell.ResolveCellRange(rangeOrAddress);
            }
            else
            {
                Address address = Cell.ResolveCellCoordinate(rangeOrAddress);
                range = new Range(address, address);
            }
            return range;
        }

        /// <summary>
        /// Determines the next unused worksheet name in the passed workbook
        /// </summary>
        /// <param name="name">Original name to start the check</param>
        /// <param name="workbook">Workbook to look for existing worksheets</param>
        /// <returns>Not yet used worksheet name</returns>
        /// <exception cref="WorksheetException">A WorksheetException is thrown if the workbook reference is null, since all worksheets have to be considered during sanitation</exception>
        /// \remark <remarks>The 'rare' case where 10^31 Worksheets exists (leads to a crash) is deliberately not handled, 
        /// since such a number of sheets would consume at least one quintillion bytes of RAM... what is vastly out of the 64 bit range</remarks>
        private static string GetUnusedWorksheetName(string name, Workbook workbook)
        {
            if (workbook == null)
            {
                throw new WorksheetException("The workbook reference is null");
            }
            if (!WorksheetExists(name, workbook))
            { return name; }
            Regex regex = new Regex(@"^(.*?)(\d{1,31})$");
            Match match = regex.Match(name);
            string prefix = name;
            int number = 1;
            if (match.Groups.Count > 1)
            {
                prefix = match.Groups[1].Value;
                int.TryParse(match.Groups[2].Value, out number);
                // if this failed, the start number is 0 (parsed number was >max. int32)
            }
            while (true)
            {
                string numberString = ParserUtils.ToString(number);
                if (numberString.Length + prefix.Length > MAX_WORKSHEET_NAME_LENGTH)
                {
                    int endIndex = prefix.Length - (numberString.Length + prefix.Length - MAX_WORKSHEET_NAME_LENGTH);
                    prefix = prefix.Substring(0, endIndex);
                }
                string newName = prefix + numberString;
                if (!WorksheetExists(newName, workbook))
                { return newName; }
                number++;
            }
        }

        /// <summary>
        /// Checks whether a worksheet with the given name exists
        /// </summary>
        /// <param name="name">Name to check</param>
        /// <param name="workbook">Workbook reference</param>
        /// <returns>True if the name exits, otherwise false</returns>
        private static bool WorksheetExists(string name, Workbook workbook)
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
