/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2020
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Xml;
using NanoXLSX.Exceptions;
using NanoXLSX.Styles;
using IOException = NanoXLSX.Exceptions.IOException;

namespace NanoXLSX.LowLevel
{
    /// <summary>
    /// Class representing a reader for worksheets of XLSX files
    /// </summary>
    public class WorksheetReader
    {
        #region privateFields

        private SharedStringsReader sharedStrings;
        private ImportOptions importOptions;
        private List<string> dateStyles;
        private List<string> timeStyles;
        #endregion

        #region properties

        /// <summary>
        /// Gets the number of the worksheet
        /// </summary>
        /// <value>
        /// Number of the worksheet
        /// </value>
        public int WorksheetNumber { get; private set; }

        /// <summary>
        /// Gets the name of the worksheet
        /// </summary>
        /// <value>
        /// Name of the worksheet
        /// </value>
        public string Name { get; private set; }

        /// <summary>
        /// Gets the data of the worksheet as Dictionary of cell address-cell object tuples
        /// </summary>
        /// <value>
        /// Dictionary of cell address-cell object tuples
        /// </value>
        public Dictionary<string, Cell> Data { get; private set; }

        /// <summary>
        /// Gets the assignment of resolved styles to cell addresses
        /// </summary>
        /// <value>Dictionary of cell address-style number tuples</value>
        public Dictionary<string, string> StyleAssignment { get; private set; } = new Dictionary<string, string>();

        #endregion

        #region constructors

        /// <summary>
        /// Constructor with parameters
        /// </summary>
        /// <param name="sharedStrings">SharedStringsReader object</param>
        /// <param name="name">Worksheet name</param>
        /// <param name="number">Worksheet number</param>
        /// <param name="styleReaderContainer">resolved styles, used to determine dates or times</param>
        /// <param name="options">Import options to override the automatic approach of the reader. <see cref="ImportOptions"/> for information about import options.</param>
        public WorksheetReader(SharedStringsReader sharedStrings, string name, int number, StyleReaderContainer styleReaderContainer, ImportOptions options = null)
        {
            importOptions = options;
            Data = new Dictionary<string, Cell>();
            Name = name;
            WorksheetNumber = number;
            this.sharedStrings = sharedStrings;
            processStyles(styleReaderContainer);
        }

        #endregion

        #region functions

        /// <summary>
        /// Determine which of the resolved styles are either to define a time or a date. Stores also the styles into a dictionary 
        /// </summary>
        /// <param name="styleReaderContainer">Resolved styles from the style reader</param>
        private void processStyles(StyleReaderContainer styleReaderContainer)
        {
            dateStyles = new List<string>();
            timeStyles = new List<string>();
            for (int i = 0; i < styleReaderContainer.StyleCount; i++)
            {
                bool isDate, isTime;
                Style style = styleReaderContainer.GetStyle(i, out isDate, out isTime, true);
                if (isDate)
                {
                    dateStyles.Add(i.ToString("G", CultureInfo.InvariantCulture));
                }
                if (isTime)
                {
                    timeStyles.Add(i.ToString("G", CultureInfo.InvariantCulture));
                }
            }
        }

        /// <summary>
        /// Gets whether the specified column exists in the data
        /// </summary>
        /// <param name="columnAddress">Column address as string</param>
        /// <returns>
        ///   Column address as string
        /// </returns>
        public bool HasColumn(string columnAddress)
        {
            if (string.IsNullOrEmpty(columnAddress)) { return false; }
            int columnNumber = Cell.ResolveColumn(columnAddress);
            foreach (KeyValuePair<string, Cell> cell in Data)
            {
                if (cell.Value.ColumnNumber == columnNumber)
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Gets whether the passed row (represented as list of cell objects) contains the specified column numbers
        /// </summary>
        /// <param name="cells">List of cell objects to check.</param>
        /// <param name="columnNumbers">Array of column numbers</param>
        /// <returns>True if all column numbers were found, otherwise false</returns>
        public bool RowHasColumns(List<Cell> cells, int[] columnNumbers)
        {
            if (columnNumbers == null || cells == null) { return false; }
            int len = columnNumbers.Length;
            int len2 = cells.Count;
            int j;
            bool match;
            if (len < 1 || len2 < 1) { return false; }
            for (int i = 0; i < len; i++)
            {
                match = false;
                for (j = 0; j < len2; j++)
                {
                    if (cells[j].ColumnNumber == columnNumbers[i])
                    {
                        match = true;
                        break;
                    }
                }
                if (match == false) { return false; }
            }
            return true;
        }

        /// <summary>
        /// Gets the number of rows
        /// </summary>
        /// <returns>Number of rows</returns>
        public int GetRowCount()
        {
            int count = -1;
            foreach (KeyValuePair<string, Cell> cell in Data)
            {
                if (cell.Value.RowNumber > count)
                {
                    count = cell.Value.RowNumber;
                }
            }
            return count + 1;
        }

        /// <summary>
        /// Gets a row as list of cell objects
        /// </summary>
        /// <param name="rowNumber">Row number</param>
        /// <returns>List of cell objects</returns>
        public List<Cell> GetRow(int rowNumber)
        {
            List<Cell> list = new List<Cell>();
            foreach (KeyValuePair<string, Cell> cell in Data)
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
        /// Reads the XML file form the passed stream and processes the worksheet data
        /// </summary>
        /// <param name="stream">Stream of the XML file</param>
        /// <exception cref="Exceptions.IOException">Throws IOException in case of an error</exception>
        public void Read(MemoryStream stream)
        {
            try
            {
                using (stream) // Close after processing
                {
                    string type, styleNumber, address, value, formula;
                    XmlDocument xr = new XmlDocument();
                    xr.Load(stream);
                    XmlNodeList rows = xr.GetElementsByTagName("row");
                    foreach (XmlNode row in rows)
                    {
                        if (row.HasChildNodes)
                        {
                            foreach (XmlNode rowChild in row.ChildNodes)
                            {
                                type = "s";
                                styleNumber = "";
                                address = "A1";
                                value = "";
                                formula = null;
                                if (rowChild.LocalName.ToLower() == "c")
                                {
                                    address = ReaderUtils.GetAttribute("r", rowChild); // Mandatory
                                    type = ReaderUtils.GetAttribute("t", rowChild); // can be null if not existing
                                    styleNumber = ReaderUtils.GetAttribute("s", rowChild); // can be null
                                    if (rowChild.HasChildNodes)
                                    {
                                        foreach (XmlNode valueNode in rowChild.ChildNodes)
                                        {
                                            if (valueNode.LocalName.ToLower() == "v")
                                            {
                                                value = valueNode.InnerText;
                                            }
                                            if (valueNode.LocalName.ToLower() == "f")
                                            {
                                                formula = valueNode.InnerText;
                                            }
                                        }
                                    }
                                }
                                ResolveCellData(address, type, value, styleNumber, formula);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new IOException("XMLStreamException", "The XML entry could not be read from the input stream. Please see the inner exception:", ex);
            }
        }

        /// <summary>
        /// Resolves the data of a read cell either automatically or conditionally  (import options), transforms it into a cell object and adds it to the data
        /// </summary>
        /// <param name="addressString">Address of the cell</param>
        /// <param name="type">Expected data type</param>
        /// <param name="value">Raw value as string</param>
        /// <param name="styleNumber">Style number as string (can be null)</param>
        /// <param name="formula"> Formula as string (can be null; data type determines whether value or formula is used)</param>
        private void ResolveCellData(string addressString, string type, string value, string styleNumber, string formula)
        {
            Address address = new Address(addressString);
            string key = addressString.ToUpper();
            StyleAssignment[key] = styleNumber;
            if (importOptions == null)
            {
                Data.Add(key, AutoResolveCellData(address, type, value, styleNumber, formula));
            }
            else
            {
                Data.Add(key, ResolveCellDataConditionally(address, type, value, styleNumber, formula));
            }
        }

        /// <summary>
        /// Resolves the data of a read cell with conditions of import options, transforms it into a cell object
        /// </summary>
        /// <param name="address">Address of the cell</param>
        /// <param name="type">Expected data type</param>
        /// <param name="value">Raw value as string</param>
        /// <param name="styleNumber">Style number as string (can be null)</param>
        /// <param name="formula"> Formula as string (can be null; data type determines whether value or formula is used)</param>
        /// <returns>The resolved Cell</returns>
        private Cell ResolveCellDataConditionally(Address address, string type, string value, string styleNumber, string formula)
        {
            if (address.Row < importOptions.EnforcingStartRowNumber)
            {
                return AutoResolveCellData(address, type, value, styleNumber, formula); // Skip enforcing
            }
            if (importOptions.EnforcedColumnTypes.ContainsKey(address.Column))
            {
                ImportOptions.ColumnType importType = importOptions.EnforcedColumnTypes[address.Column];
                switch (importType)
                {
                    case ImportOptions.ColumnType.Bool:
                        return GetBooleanValue(value, address);
                    case ImportOptions.ColumnType.Date:
                        if (importOptions.EnforceDateTimesAsNumbers)
                        {
                            return GetNumericValue(value, address);
                        }
                        else
                        {
                            return GetDateTimeValue(value, address, Cell.CellType.DATE);
                        }
                    case ImportOptions.ColumnType.Time:
                        if (importOptions.EnforceDateTimesAsNumbers)
                        {
                            return GetNumericValue(value, address);
                        }
                        else
                        {
                            return GetDateTimeValue(value, address, Cell.CellType.TIME);
                        }
                    case ImportOptions.ColumnType.Numeric:
                        return GetNumericValue(value, address);
                    case ImportOptions.ColumnType.String:
                        return GetStringValue(value, address);
                    default:
                        return AutoResolveCellData(address, type, value, styleNumber, formula);
                }
            }
            else
            {
                return AutoResolveCellData(address, type, value, styleNumber, formula);
            }
        }

        /// <summary>
        /// Resolves the data of a read cell automatically, transforms it into a cell object
        /// </summary>
        /// <param name="address">Address of the cell</param>
        /// <param name="type">Expected data type</param>
        /// <param name="value">Raw value as string</param>
        /// <param name="styleNumber">Style number as string (can be null)</param>
        /// <param name="formula"> Formula as string (can be null; data type determines whether value or formula is used)</param>
        /// <returns>The resolved Cell</returns>
        private Cell AutoResolveCellData(Address address, string type, string value, string styleNumber, string formula)
        {
            if (type == "s") // string (declared)
            {
                return GetStringValue(value, address);
            }
            else if (type == "b") // boolean
            {
                return GetBooleanValue(value, address);
            }
            else if (dateStyles.Contains(styleNumber))  // date (priority)
            {
                return GetDateTimeValue(value, address, Cell.CellType.DATE);
            }
            else if (timeStyles.Contains(styleNumber)) // time
            {
                return GetDateTimeValue(value, address, Cell.CellType.TIME);
            }
            else if (type == null) // try numeric if not parsed as date or time, before numeric
            {
                return GetNumericValue(value, address);
            }
            else if (formula != null) // formula before string
            {
                return new Cell(formula, Cell.CellType.FORMULA, address);
            }
            else // fall back to sting
            {
                return GetStringValue(value, address);
            }
        }

        /// <summary>
        /// Parses the numeric (double) value of a raw cell
        /// </summary>
        /// <param name="raw">Raw value as string</param>
        /// <param name="address">Address of the cell</param>
        /// <returns>Cell of the type double or the defined fall-back type</returns>
        private Cell GetNumericValue(string raw, Address address)
        {
            double dValue;
            if (double.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out dValue))
            {
                return new Cell(dValue, Cell.CellType.NUMBER, address);
            }
            else
            {
                return new Cell(raw, Cell.CellType.STRING, address);
            }
        }

        /// <summary>
        /// Parses the string value of a raw cell. May take the value from the shared string table, if available
        /// </summary>
        /// <param name="raw">Raw value as string</param>
        /// <param name="address">Address of the cell</param>
        /// <returns>Cell of the type string</returns>
        private Cell GetStringValue(string raw, Address address)
        {
            int stringId;
            if (int.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out stringId))
            {
                string resolvedString = sharedStrings.GetString(stringId);
                if (resolvedString == null)
                {
                    return new Cell(raw, Cell.CellType.STRING, address);
                }
                else
                {
                    return new Cell(resolvedString, Cell.CellType.STRING, address);
                }
            }
            else
            {
                return new Cell(raw, Cell.CellType.STRING, address);
            }
        }

        /// <summary>
        /// Parses the boolean value of a raw cell
        /// </summary>
        /// <param name="raw">Raw value as string</param>
        /// <param name="address">Address of the cell</param>
        /// <returns>Cell of the type bool or the defined fall-back type</returns>
        private Cell GetBooleanValue(String raw, Address address)
        {
            if (raw == "0")
            {
                return new Cell(false, Cell.CellType.BOOL, address);
            }
            else if (raw == "1")
            {
                return new Cell(true, Cell.CellType.BOOL, address);
            }
            else
            {
                bool value;
                if (bool.TryParse(raw, out value))
                {
                    return new Cell(value, Cell.CellType.BOOL, address);
                }
                else
                {
                    return new Cell(raw, Cell.CellType.STRING, address);

                }
            }
        }

        /// <summary>
        /// Parses the date (DateTime) or time (TimeSpan) value of a raw cell. If the value is numeric, but out of range of a OAdate, a numeric value will be returned instead. If invalid, the string representation will be returned.
        /// </summary>
        /// <param name="raw">Raw value as string</param>
        /// <param name="address">Address of the cell</param>
        /// <param name="type">Type of the value zu be converted: Valid values are DATE and TIME</param>
        /// <returns>Cell of the type TimeSpan or the defined fall-back type</returns>
        private Cell GetDateTimeValue(String raw, Address address, Cell.CellType type)
        {
            double dValue;
            if (double.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out dValue))
            {
                if (dValue < XlsxWriter.MIN_OADATE_VALUE || dValue > XlsxWriter.MAX_OADATE_VALUE)
                {
                    return new Cell(dValue, Cell.CellType.NUMBER, address); // Invalid OAdate == plain number
                }
                else
                {
                    switch (type)
                    {
                        case Cell.CellType.DATE:
                            DateTime date = DateTime.FromOADate(dValue);
                            return new Cell(date, Cell.CellType.DATE, address);
                        case Cell.CellType.TIME:
                            TimeSpan time = TimeSpan.FromSeconds(dValue * 86400d);
                            return new Cell(time, Cell.CellType.TIME, address);
                        default:
                            throw new ArgumentException("The defined type is not supported to be uses as date or time");
                    }
                }
            }
            else
            {
                return new Cell(raw, Cell.CellType.STRING, address);
            }
        }

        #endregion

    }
}
