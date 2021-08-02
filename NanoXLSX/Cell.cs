/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2021
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using NanoXLSX.Exceptions;
using NanoXLSX.Styles;
using FormatException = NanoXLSX.Exceptions.FormatException;

namespace NanoXLSX
{
    /// <summary>
    /// Class representing a cell of a worksheet
    /// </summary>
    public class Cell : IComparable<Cell>
    {
        #region constants
        private const int ASCII_OFFSET = 64;
        #endregion


        #region enums
        /// <summary>
        /// Enum defines the basic data types of a cell
        /// </summary>
        public enum CellType
        {
            /// <summary>Type for single characters and strings</summary>
            STRING,
            /// <summary>Type for all numeric types (long, integer and float and double)</summary>
            NUMBER,
            /// <summary>Type for dates(Note: Dates before 1900-01-01 and after 9999-12-31 are not allowed)</summary>
            DATE,
            /// <summary>Type for times (Note: Internally handled as OAdate, represented by <see cref="TimeSpan"/>)</summary>
            TIME,
            /// <summary>Type for boolean</summary>
            BOOL,
            /// <summary>Type for Formulas (The cell will be handled differently)</summary>
            FORMULA,
            /// <summary>Type for empty cells. This type is only used for merged cells (all cells except the first of the cell range)</summary>
            EMPTY,
            /// <summary>Default Type, not specified</summary>
            DEFAULT
        }

        /// <summary>
        /// Enum for the referencing style of the address
        /// </summary>
        public enum AddressType
        {
            /// <summary>Default behavior (e.g. 'C3')</summary>
            Default,
            /// <summary>Row of the address is fixed (e.g. 'C$3')</summary>
            FixedRow,
            /// <summary>Column of the address is fixed (e.g. '$C3')</summary>
            FixedColumn,
            /// <summary>Row and column of the address is fixed (e.g. '$C$3')</summary>
            FixedRowAndColumn
        }

        /// <summary>
        /// Enum to define the scope of a passed address string (used in static context)
        /// </summary>
        public enum AddressScope
        {
            /// <summary>The address represents a single cell</summary>
            SingleAddress,
            /// <summary>The address represents a range of cells</summary>
            Range,
            /// <summary>The address expression is invalid</summary>
            Invalid
        }

        #endregion

        #region privateFileds
        private Style cellStyle;
        private int columnNumber;
        private int rowNumber;
        #endregion

        #region properties

        /// <summary>
        /// Gets or sets the combined cell Address as string in the format A1 - XFD1048576. The address may contain a <see cref="Cell.AddressType"/> modifier (e.g. C$50)
        /// </summary>
        public string CellAddress
        {
            get { return ResolveCellAddress(ColumnNumber, RowNumber, CellAddressType); }
            set {
                AddressType addressType;
                ResolveCellCoordinate(value, out columnNumber, out rowNumber, out addressType);
                CellAddressType = addressType;
            }
        }

        /// <summary>Gets or sets the combined cell Address as Address object</summary>
        public Address CellAddress2
        {
            get { return new Address(ColumnNumber, RowNumber, CellAddressType); }
            set
            {
                ColumnNumber = value.Column;
                RowNumber = value.Row;
                CellAddressType = value.Type;
            }
        }

        /// <summary>
        /// Gets the assigned style of the cell
        /// </summary>
        public Style CellStyle
        {
            get { return cellStyle; }
        }

        /// <summary>Gets or sets the number of the column (zero-based)</summary>  
        /// <exception cref="RangeException">Throws a RangeException if the column number is out of range</exception>
        public int ColumnNumber
        {
            get { return columnNumber; }
            set
            {
                ValidateColumnNumber(value);
                columnNumber = value;
            }
        }

        /// <summary>Gets or sets the type of the cell</summary>
        public CellType DataType { get; set; }


        /// <summary>Gets or sets the number of the row (zero-based)</summary>
        /// <exception cref="RangeException">Throws a RangeException if the row number is out of range</exception>
        public int RowNumber
        {
            get { return rowNumber; }
            set
            {
                ValidateRowNumber(value);
                rowNumber = value;
            }
        }

        /// <summary>
        /// Gets or sets the optional address type that can be part of the cell address. 
        /// </summary>
        /// <remarks>The type has no influence on the behavior of the cell, though. It is preserved to avoid losing information on the address object of the cell</remarks>
        public AddressType CellAddressType { get; set; }

        /// <summary>Gets or sets the value of the cell (generic object type)</summary>
        public object Value { get; set; }

        #endregion

        #region constructors
        /// <summary>Default constructor. Cells created with this constructor do not have a link to a worksheet initially</summary>
        public Cell()
        {
            DataType = CellType.DEFAULT;
        }

        /// <summary>
        /// Constructor with value and cell type. Cells created with this constructor do not have a link to a worksheet initially
        /// </summary>
        /// <param name="value">Value of the cell</param>
        /// <param name="type">Type of the cell</param>
        public Cell(object value, CellType type)
        {
            Value = value;
            DataType = type;
            if (type == CellType.DEFAULT)
            {
                ResolveCellType();
            }
        }

        /// <summary>
        /// Constructor with value, cell type and address as string. The worksheet reference is set to null and must be assigned later
        /// </summary>
        /// <param name="value">Value of the cell</param>
        /// <param name="type">Type of the cell</param>
        /// <param name="address">Address of the cell</param>
        public Cell(Object value, CellType type, string address)
        {
            DataType = type;
            Value = value;
            CellAddress = address;
            if (type == CellType.DEFAULT)
            {
                ResolveCellType();
            }
        }

        /// <summary>
        /// Constructor with value, cell type and address as struct. The worksheet reference is set to null and must be assigned later
        /// </summary>
        /// <param name="value">Value of the cell</param>
        /// <param name="type">Type of the cell</param>
        /// <param name="address">Address struct of the cell</param>
        public Cell(Object value, CellType type, Address address)
        {
            DataType = type;
            Value = value;
            columnNumber = address.Column;
            rowNumber = address.Row;
            CellAddressType = address.Type;
            if (type == CellType.DEFAULT)
            {
                ResolveCellType();
            }
        }

        /// <summary>
        /// Constructor with value, cell type, row number and column number
        /// </summary>
        /// <param name="value">Value of the cell</param>
        /// <param name="type">Type of the cell</param>
        /// <param name="column">Column number of the cell (zero-based)</param>
        /// <param name="row">Row number of the cell (zero-based)</param>
        public Cell(object value, CellType type, int column, int row) : this(value, type)
        {
            ColumnNumber = column;
            RowNumber = row;
            CellAddressType = AddressType.Default;
            if (type == CellType.DEFAULT)
            {
                ResolveCellType();
            }
        }
        #endregion

        #region methods
        /// <summary>
        /// Implemented CompareTo method
        /// </summary>
        /// <param name="other">Object to compare</param>
        /// <returns>0 if values are the same, -1 if this object is smaller, 1 if it is bigger</returns>
        /// <remarks>Note that this method only compares the row and column numbers, 
        /// since the values or styles may completely different types, and therefore hard to compare at all.<br/>
        /// The <see cref="Equals(object)"/> method considers values and style, though.</remarks>
        public int CompareTo(Cell other)
        {
            if (other == null)
            {
                return -1;
            }
            if (RowNumber == other.RowNumber)
            {
                return ColumnNumber.CompareTo(other.ColumnNumber);
            }

            return RowNumber.CompareTo(other.RowNumber);
        }

        public override bool Equals(object obj)
        {
            if (obj == null || obj.GetType() != typeof(Cell))
            {
                return false;
            }
            Cell other = (Cell)obj;
            if (!this.CellAddress2.Equals(other.CellAddress2))
            {
                return false;
            }
            if (this.cellStyle != null && other.CellStyle != null && !this.CellStyle.Equals(other.CellStyle))
            {
                return false;
            }
            if (this.DataType != other.DataType)
            {
                return false;
            }
            if (this.Value != null && other.Value != null && !this.Value.Equals(other.Value))
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// Removes the assigned style from the cell
        /// </summary>
        public void RemoveStyle()
        {
                cellStyle = null;
        }

        /// <summary>
        /// Method resets the Cell type and tries to find the actual type. This is used if a Cell was created with the CellType DEFAULT. CellTypes FORMULA and EMPTY will skip this method
        /// </summary>
        public void ResolveCellType()
        {
            if (Value == null)
            {
                DataType = CellType.EMPTY;
                Value = null;
                return;
            }
            if (DataType == CellType.FORMULA || DataType == CellType.EMPTY)
            { return; }
            Type t = Value.GetType();
            if (t == typeof(bool))
            { DataType = CellType.BOOL; }
            else if (t == typeof(byte) || t == typeof(sbyte))
            { DataType = CellType.NUMBER; }
            else if (t == typeof(decimal))
            { DataType = CellType.NUMBER; }
            else if (t == typeof(double))
            { DataType = CellType.NUMBER; }
            else if (t == typeof(float))
            { DataType = CellType.NUMBER; }
            else if (t == typeof(int) || t == typeof(uint))
            { DataType = CellType.NUMBER; }
            else if (t == typeof(long) || t == typeof(ulong))
            { DataType = CellType.NUMBER; }
            else if (t == typeof(short) || t == typeof(ushort))
            { DataType = CellType.NUMBER; }
            else if (t == typeof(DateTime)) // Not native but standard
            {
                DataType = CellType.DATE;
                SetStyle(BasicStyles.DateFormat);
            }
           
            else if (t == typeof(TimeSpan)) // Not native but standard
            { 
                DataType = CellType.TIME;
                SetStyle(BasicStyles.TimeFormat);
            } 
            else { DataType = CellType.STRING; } // Default (char, string, object)
        }

        /// <summary>
        /// Sets the lock state of the cell
        /// </summary>
        /// <param name="isLocked">If true, the cell will be locked if the worksheet is protected</param>
        /// <param name="isHidden">If true, the value of the cell will be invisible if the worksheet is protected</param>
        /// <exception cref="StyleException">Throws an UndefinedStyleException if the style used to lock cells cannot be referenced</exception>
        /// <remarks>The listed exception should never happen because the mentioned style is internally generated</remarks>
        public void SetCellLockedState(bool isLocked, bool isHidden)
        {
            Style lockStyle;
            if (cellStyle == null)
            {
                lockStyle = new Style();
            }
            else
            {
                lockStyle = cellStyle.CopyStyle();
            }
            lockStyle.CurrentCellXf.Locked = isLocked;
            lockStyle.CurrentCellXf.Hidden = isHidden;
            SetStyle(lockStyle);
        }

        /// <summary>
        /// Sets the style of the cell
        /// </summary>
        /// <param name="style">Style to assign</param>
        /// <param name="unmanaged">Internally used: If true, the style repository is not invoked and only the style object of the cell is updated. Do not use!</param>
        /// <returns>If the passed style already exists in the repository, the existing one will be returned, otherwise the passed one</returns>
        public Style SetStyle(Style style, bool unmanaged = false)
        {
            if (style == null)
            {
                throw new StyleException(StyleException.MISSING_REFERENCE, "No style to assign was defined");
            }
            cellStyle = unmanaged ? style : StyleRepository.Instance.AddStyle(style);
            return cellStyle;
        }


        #endregion

        #region staticMethods
        /// <summary>
        /// Converts a List of supported objects into a list of cells
        /// </summary>
        /// <typeparam name="T">Generic data type</typeparam>
        /// <param name="list">List of generic objects</param>
        /// <returns>List of cells</returns>
        public static IEnumerable<Cell> ConvertArray<T>(IEnumerable<T> list)
        {
            List<Cell> output = new List<Cell>();
            if (list == null)
            {
                return output;
            }
            Cell c;
            object o;
            Type t;
            foreach (T item in list)
            {
                if (item == null)
                { 
                    c = new Cell(null, CellType.EMPTY);
                    output.Add(c);
                    continue;
                }
                o = item; // intermediate object is necessary to cast the types below
                t = item.GetType();
                if (t == typeof(bool))
                { c = new Cell((bool)o, CellType.BOOL); }
                else if (t == typeof(byte))
                { c = new Cell((byte)o, CellType.NUMBER); }
                else if (t == typeof(sbyte))
                { c = new Cell((sbyte)o, CellType.NUMBER); }
                else if (t == typeof(decimal))
                { c = new Cell((decimal)o, CellType.NUMBER); }
                else if (t == typeof(double))
                { c = new Cell((double)o, CellType.NUMBER); }
                else if (t == typeof(float))
                { c = new Cell((float)o, CellType.NUMBER); }
                else if (t == typeof(int))
                { c = new Cell((int)o, CellType.NUMBER); }
                else if (t == typeof(uint))
                { c = new Cell((uint)o, CellType.NUMBER); }
                else if (t == typeof(long))
                { c = new Cell((long)o, CellType.NUMBER); }
                else if (t == typeof(ulong))
                { c = new Cell((ulong)o, CellType.NUMBER); }
                else if (t == typeof(short))
                { c = new Cell((short)o, CellType.NUMBER); }
                else if (t == typeof(ushort))
                { c = new Cell((ushort)o, CellType.NUMBER); }
                else if (t == typeof(DateTime))
                { c = new Cell((DateTime)o, CellType.DATE); }
                else if (t == typeof(TimeSpan))
                { c = new Cell((TimeSpan)o, CellType.TIME); }
                else if (t == typeof(string))
                { c = new Cell((string)o, CellType.STRING); }
                else // Default = unspecified object
                {
                     c = new Cell(o.ToString(), CellType.DEFAULT); 
                }
                output.Add(c);
            }
            return output;
        }

        /// <summary>
        /// Gets a list of cell addresses from a cell range (format A1:B3 or AAD556:AAD1000)
        /// </summary>
        /// <param name="range">Range to process</param>
        /// <returns>List of cell addresses</returns>
        /// <exception cref="Exceptions.FormatException">Throws a FormatException if a part of the passed range is malformed</exception>
        /// <exception cref="RangeException">Throws an RangeException if the range is out of range (A-XFD and 1 to 1048576) </exception>
        public static IEnumerable<Address> GetCellRange(string range)
        {
            Range range2 = ResolveCellRange(range);
            return GetCellRange(range2.StartAddress, range2.EndAddress);
        }

        /// <summary>
        /// Get a list of cell addresses from a cell range
        /// </summary>
        /// <param name="startAddress">Start address as string in the format A1 - XFD1048576</param>
        /// <param name="endAddress">End address as string in the format A1 - XFD1048576</param>
        /// <returns>List of cell addresses</returns>
        /// <exception cref="Exceptions.FormatException">Throws a FormatException if a part of the passed range is malformed</exception>
        /// <exception cref="RangeException">Throws an RangeException if the range is out of range (A-XFD and 1 to 1048576) </exception> 
        public static IEnumerable<Address> GetCellRange(string startAddress, string endAddress)
        {
            Address start = ResolveCellCoordinate(startAddress);
            Address end = ResolveCellCoordinate(endAddress);
            return GetCellRange(start, end);
        }

        /// <summary>
        /// Get a list of cell addresses from a cell range
        /// </summary>
        /// <param name="startColumn">Start column (zero based)</param>
        /// <param name="startRow">Start row (zero based)</param>
        /// <param name="endColumn">End column (zero based)</param>
        /// <param name="endRow">End row (zero based)</param>
        /// <returns>List of cell addresses</returns>
        /// <exception cref="RangeException">Throws an RangeException if the value of one passed address parts is out of range (A-XFD and 1 to 1048576) </exception>
        public static IEnumerable<Address> GetCellRange(int startColumn, int startRow, int endColumn, int endRow)
        {
            Address start = new Address(startColumn, startRow);
            Address end = new Address(endColumn, endRow);
            return GetCellRange(start, end);
        }

        /// <summary>
        /// Get a list of cell addresses from a cell range
        /// </summary>
        /// <param name="startAddress">Start address</param>
        /// <param name="endAddress">End address</param>
        /// <returns>List of cell addresses</returns>
        /// <exception cref="Exceptions.FormatException">Throws a FormatException if a part of the passed addresses is malformed</exception>
        /// <exception cref="RangeException">Throws an RangeException if the value of one passed address is out of range (A-XFD and 1 to 1048576) </exception>
        public static IEnumerable<Address> GetCellRange(Address startAddress, Address endAddress)
        {
            int startColumn;
            int endColumn;
            int startRow;
            int endRow;
            if (startAddress.Column < endAddress.Column)
            {
                startColumn = startAddress.Column;
                endColumn = endAddress.Column;
            }
            else
            {
                startColumn = endAddress.Column;
                endColumn = startAddress.Column;
            }
            if (startAddress.Row < endAddress.Row)
            {
                startRow = startAddress.Row;
                endRow = endAddress.Row;
            }
            else
            {
                startRow = endAddress.Row;
                endRow = startAddress.Row;
            }
            List<Address> output = new List<Address>();
            for (int column = startColumn; column <= endColumn; column++)
            {
                for (int row = startRow; row <= endRow; row++)
                {
                    output.Add(new Address(column, row));
                }
            }
            return output;
        }

        /// <summary>
        /// Gets the address of a cell by the column and row number (zero based)
        /// </summary>
        /// <param name="column">Column number of the cell (zero-based)</param>
        /// <param name="row">Row number of the cell (zero-based)</param>
        /// <param name="type">Optional referencing type of the address</param>
        /// <exception cref="RangeException">Throws an RangeException if the start or end address was out of range</exception>
        /// <returns>Cell Address as string in the format A1 - XFD1048576. Depending on the type, Addresses like '$A55', 'B$2' or '$A$5' are possible outputs</returns>
        public static string ResolveCellAddress(int column, int row, AddressType type = AddressType.Default)
        {
            ValidateColumnNumber(column);
            ValidateRowNumber(row);
            switch (type)
            {
                case AddressType.FixedRowAndColumn:
                    return "$" + ResolveColumnAddress(column) + "$" + (row + 1);
                case AddressType.FixedColumn:
                    return "$" + ResolveColumnAddress(column) + (row + 1);
                case AddressType.FixedRow:
                    return ResolveColumnAddress(column) + "$" + (row + 1);
                default:
                    return ResolveColumnAddress(column) + (row + 1);
            }
        }

        /// <summary>
        /// Gets the column and row number (zero based) of a cell by the address
        /// </summary>
        /// <param name="address">Address as string in the format A1 - XFD1048576</param>
        /// <returns>Struct with row and column</returns>
        /// <exception cref="Exceptions.FormatException">Throws a FormatException if the passed address is malformed</exception>
        /// <exception cref="RangeException">Throws an RangeException if the value of the passed address is out of range (A-XFD and 1 to 1048576) </exception>
        public static Address ResolveCellCoordinate(string address)
        {
            int row;
            int column;
            AddressType type;
            ResolveCellCoordinate(address, out column, out row, out type);
            return new Address(column, row, type);
        }

        /// <summary>
        /// Gets the column and row number (zero based) of a cell by the address
        /// </summary>
        /// <param name="address">Address as string in the format A1 - XFD1048576</param>
        /// <param name="column">Column number of the cell (zero-based) as out parameter</param>
        /// <param name="row">Row number of the cell (zero-based) as out parameter</param>
        /// <exception cref="Exceptions.FormatException">Throws a FormatException if the range address was malformed</exception>
        /// <exception cref="RangeException">Throws an RangeException if the row or column number was out of range</exception>
        public static void ResolveCellCoordinate(string address, out int column, out int row)
        {
            AddressType dummy;
            ResolveCellCoordinate(address, out column, out row, out dummy);
        }


        /// <summary>
        /// Gets the column and row number (zero based) of a cell by the address
        /// </summary>
        /// <param name="address">Address as string in the format A1 - XFD1048576</param>
        /// <param name="column">Column number of the cell (zero-based) as out parameter</param>
        /// <param name="row">Row number of the cell (zero-based) as out parameter</param>
        /// <param name="addressType">Address type of the cell (if defined as modifiers in the address string)</param>
        /// <exception cref="Exceptions.FormatException">Throws a FormatException if the range address was malformed</exception>
        /// <exception cref="RangeException">Throws an RangeException if the row or column number was out of range</exception>
        public static void ResolveCellCoordinate(string address, out int column, out int row, out AddressType addressType)
        {
            if (string.IsNullOrEmpty(address))
            {
                throw new FormatException("The cell address is null or empty and could not be resolved");
            }
            address = Utils.ToUpper(address);
            Regex pattern = new Regex("(^(\\$?)([A-Z]{1,3})(\\$?)([0-9]{1,7})$)");
            Match matcher = pattern.Match(address);
            if (matcher.Groups.Count != 6)
            {
                throw new FormatException("The format of the cell address (" + address + ") is malformed");
            }
            int digits = int.Parse(matcher.Groups[5].Value, CultureInfo.InvariantCulture);
            column = ResolveColumn(matcher.Groups[3].Value);
            row = digits - 1;
            ValidateRowNumber(row);
            if (!String.IsNullOrEmpty(matcher.Groups[2].Value) && !String.IsNullOrEmpty(matcher.Groups[4].Value))
            {
                addressType = AddressType.FixedRowAndColumn;
            }
            else if (!String.IsNullOrEmpty(matcher.Groups[2].Value) && String.IsNullOrEmpty(matcher.Groups[4].Value))
            {
                addressType = AddressType.FixedColumn;
            }
            else if (String.IsNullOrEmpty(matcher.Groups[2].Value) && !String.IsNullOrEmpty(matcher.Groups[4].Value))
            {
                addressType = AddressType.FixedRow;
            }
            else
            {
                addressType = AddressType.Default;
            }
        }

        /// <summary>
        /// Resolves a cell range from the format like A1:B3 or AAD556:AAD1000
        /// </summary>
        /// <param name="range">Range to process</param>
        /// <returns>Range object</returns>
        /// <exception cref="Exceptions.FormatException">Throws a FormatException if the start or end address was malformed</exception>
        /// <exception cref="RangeException">Throws an RangeException if the range is out of range (A-XFD and 1 to 1048576) </exception>
        public static Range ResolveCellRange(string range)
        {
            if (string.IsNullOrEmpty(range))
            {
                throw new FormatException("The cell range is null or empty and could not be resolved");
            }
            string[] split = range.Split(':');
            if (split.Length != 2)
            {
                throw new FormatException("The cell range (" + range + ") is malformed and could not be resolved");
            }
            return new Range(ResolveCellCoordinate(split[0]), ResolveCellCoordinate(split[1]));
        }

        /// <summary>
        /// Gets the column number from the column address (A - XFD)
        /// </summary>
        /// <param name="columnAddress">Column address (A - XFD)</param>
        /// <returns>Column number (zero-based)</returns>
        /// <exception cref="RangeException">Throws an RangeException if the passed address was out of range</exception>
        public static int ResolveColumn(string columnAddress)
        {
            if (String.IsNullOrEmpty(columnAddress))
            {
                throw new RangeException(RangeException.GENERAL, "The passed address was null or empty");
            }
            columnAddress = columnAddress.ToUpper();
            int chr;
            int result = 0;
            int multiplier = 1;
            for (int i = columnAddress.Length - 1; i >= 0; i--)
            {
                chr = columnAddress[i];
                chr -= ASCII_OFFSET;
                result += (chr * multiplier);
                multiplier *= 26;
            }
            ValidateColumnNumber(result - 1);
            return result - 1;
        }

        /// <summary>
        /// Gets the column address (A - XFD)
        /// </summary>
        /// <param name="columnNumber">Column number (zero-based)</param>
        /// <returns>Column address (A - XFD)</returns>
        /// <exception cref="RangeException">Throws an RangeException if the passed column number was out of range</exception>
        public static string ResolveColumnAddress(int columnNumber)
        {
            ValidateColumnNumber(columnNumber);
            // A - XFD
            int j = 0;
            int k = 0;
            int l = 0;
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i <= columnNumber; i++)
            {
                if (j > 25)
                {
                    k++;
                    j = 0;
                }
                if (k > 25)
                {
                    l++;
                    k = 0;
                }
                j++;
            }
            if (l > 0)
            { sb.Append((char)(l + ASCII_OFFSET)); }
            if (k > 0)
            { sb.Append((char)(k + ASCII_OFFSET)); }
            sb.Append((char)(j + ASCII_OFFSET));
            return sb.ToString();
        }

        /// <summary>
        /// Gets the scope of the passed address (string expression). Scope means either single cell address or range
        /// </summary>
        /// <param name="addressExpression">Address expression</param>
        /// <returns>Scope of the address expression</returns>
        public static AddressScope GetAddressScope(String addressExpression)
        {
            try
            {
                ResolveCellCoordinate(addressExpression);
                return AddressScope.SingleAddress;
            }
            catch
            {
                try
                {
                    ResolveCellRange(addressExpression);
                    return AddressScope.Range;
                }
                catch
                {
                    return AddressScope.Invalid;
                }
            }
            
        }

        /// <summary>
        /// Validates the passed (zero-based) column number. An exception will be thrown if the column is invalid
        /// </summary>
        /// <param name="column">Number to check</param>
        /// <exception cref="RangeException">Thrown if the passed column number is out of range</exception>
        public static void ValidateColumnNumber(int column)
        {
            if (column > Worksheet.MAX_COLUMN_NUMBER || column < Worksheet.MIN_COLUMN_NUMBER)
            {
                throw new RangeException(RangeException.GENERAL, "The column number (" + column + ") is out of range. Range is from " +
                    Worksheet.MIN_COLUMN_NUMBER + " to " + Worksheet.MAX_COLUMN_NUMBER + " (" + (Worksheet.MAX_COLUMN_NUMBER + 1) + " columns).");
            }
        }

        /// <summary>
        /// Validates the passed (zero-based) row number. An exception will be thrown if the row is invalid
        /// </summary>
        /// <param name="row">Number to check</param>
        /// <exception cref="RangeException">Thrown if the passed row number is out of range</exception>
        public static void ValidateRowNumber(int row)
        {
            if (row > Worksheet.MAX_ROW_NUMBER || row < Worksheet.MIN_ROW_NUMBER)
            {
                throw new RangeException(RangeException.GENERAL, "The row number (" + row + ") is out of range. Range is from " +
                    Worksheet.MIN_ROW_NUMBER + " to " + Worksheet.MAX_ROW_NUMBER + " (" + (Worksheet.MAX_ROW_NUMBER + 1) + " rows).");
            }
        }

        #endregion


    }
}
