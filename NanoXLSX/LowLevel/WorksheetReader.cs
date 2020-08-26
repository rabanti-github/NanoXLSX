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

        #endregion

        #region constructors

        /// <summary>
        /// Constructor with parameters
        /// </summary>
        /// <param name="sharedStrings">SharedStringsReader object</param>
        /// <param name="name">Worksheet name</param>
        /// <param name="number">Worksheet number</param>
        /// <param name="options">Import options to override the automatic approach of the reader. <see cref="ImportOptions"/> for information about import options.</param>
        public WorksheetReader(SharedStringsReader sharedStrings, string name, int number, ImportOptions options = null)
        {
            importOptions = options;
            Data = new Dictionary<string, Cell>();
            Name = name;
            WorksheetNumber = number;
            this.sharedStrings = sharedStrings;
        }

        #endregion

        #region functions

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
                    string type, style, address, value, formula;
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
                                style = "";
                                address = "A1";
                                value = "";
                                formula = null;
                                if (rowChild.LocalName.ToLower() == "c")
                                {
                                    address = GetAttribute("r", rowChild, null); // Mandatory
                                    type = GetAttribute("t", rowChild, null); // can be null if not existing
                                    style = GetAttribute("s", rowChild, null); // can be null; if "1" then date
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
                                ResolveCellData(address, type, value, style, formula);
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
        /// Gets the attribute with the passed name.
        /// </summary>
        /// <param name="targetName">Name of the target attribute</param>
        /// <param name="node">XML node that contains the attribute</param>
        /// <param name="defaultValue">Default value if the attribute was not found</param>
        /// <returns>Attribute value as string or default value if not found (can be null)</returns>
        private string GetAttribute(string targetName, XmlNode node, string defaultValue)
        {
            if (node.Attributes == null || node.Attributes.Count == 0)
            {
                return defaultValue;
            }

            foreach (XmlAttribute attribute in node.Attributes)
            {
                if (attribute.Name == targetName)
                {
                    return attribute.Value;
                }
            }

            return defaultValue;
        }


        /// <summary>
        /// Resolves the data of a read cell either automatically or conditionally  (import options), transforms it into a cell object and adds it to the data
        /// </summary>
        /// <param name="addressString">Address of the cell</param>
        /// <param name="type">Expected data type</param>
        /// <param name="value">Raw value as string</param>
        /// <param name="style">Style definition as string (can be null)</param>
        /// <param name="formula"> Formula as string (can be null; data type determines whether value or formula is used)</param>
        private void ResolveCellData(string addressString, string type, string value, string style, string formula)
        {
            Address address = new Address(addressString);
            string key = addressString.ToUpper();
            if (importOptions == null)
            {
                Data.Add(key, AutoResolveCellData(address, type, value, style, formula));
            }
            else
            {
                Data.Add(key, AutoResolveCellDataConditionally(address, type, value, style, formula));
            }
        }

        /// <summary>
        /// Resolves the data of a read cell with conditions of import options, transforms it into a cell object
        /// </summary>
        /// <param name="address">Address of the cell</param>
        /// <param name="type">Expected data type</param>
        /// <param name="value">Raw value as string</param>
        /// <param name="style">Style definition as string (can be null)</param>
        /// <param name="formula"> Formula as string (can be null; data type determines whether value or formula is used)</param>
        /// <returns>The resolved Cell</returns>
        private Cell AutoResolveCellDataConditionally(Address address, string type, string value, string style, string formula)
        {
            if (address.Row < importOptions.EnforcingStartRowNumber)
            {
                return AutoResolveCellData(address, type, value, style, formula); // Skip enforcing
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
                            return GetDateValue(value, address, "1"); // TODO: Hack; This is a workaround until a style reader is implemented
                        }
                    case ImportOptions.ColumnType.Time:
                        if (importOptions.EnforceDateTimesAsNumbers)
                        {
                            return GetNumericValue(value, address);
                        }
                        else
                        {
                            return GetDateValue(value, address, "3"); // TODO: Hack; This is a workaround until a style reader is implemented
                        }
                    case ImportOptions.ColumnType.Numeric:
                        return GetNumericValue(value, address);
                    case ImportOptions.ColumnType.String:
                        return GetStringValue(value, address, sharedStrings);
                    default:
                        return AutoResolveCellData(address, type, value, style, formula);
                }
            }
            else
            {
                return AutoResolveCellData(address, type, value, style, formula);
            }
        }

        /// <summary>
        /// Resolves the data of a read cell automatically, transforms it into a cell object
        /// </summary>
        /// <param name="address">Address of the cell</param>
        /// <param name="type">Expected data type</param>
        /// <param name="value">Raw value as string</param>
        /// <param name="style">Style definition as string (can be null)</param>
        /// <param name="formula"> Formula as string (can be null; data type determines whether value or formula is used)</param>
        /// <returns>The resolved Cell</returns>
        private Cell AutoResolveCellData(Address address, string type, string value, string style, string formula)
        {
            if (type == "b") // boolean
            {
                return GetBooleanValue(value, address);
            }
            // TODO: Hack; This is a workaround until a style reader is implemented
            else if (style == "1" || style == "3")  // Try to resolve dates or times before numeric values (if a style is defined)
            {
                return GetDateValue(value, address, style); 
            }
            else if (type == null) // try numeric if not parsed as date or time, before numeric
            {
                return GetNumericValue(value, address);
            }
            else if (formula != null) // formula before string
            {
                return new Cell(formula, Cell.CellType.FORMULA, address);
            }
            else if (type == "s") // string (declared)
            {
                return GetStringValue(value, address, sharedStrings);
            }
            else // fall back to sting
            {
                return GetStringValue(value, address, sharedStrings);
            }
        }

        /// <summary>
        /// Parses the numeric (double) value of a raw cell
        /// </summary>
        /// <param name="raw">Raw value as string</param>
        /// <param name="address">Address of the cell</param>
        /// <returns>Cell of the type double or the defined fall-back type</returns>
        private static Cell GetNumericValue(string raw, Address address)
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
        /// Parses the string value of a raw cell. may be take the value from the shared string table
        /// </summary>
        /// <param name="raw">Raw value as string</param>
        /// <param name="address">Address of the cell</param>
        /// <param name="sharedStrings">Shared string table</param>
        /// <returns>Cell of the type string</returns>
        private static Cell GetStringValue(string raw, Address address, SharedStringsReader sharedStrings)
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
        private static Cell GetBooleanValue(String raw, Address address)
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
        /// Parses the date (DateTime) value of a raw cell. If the value is numeric, but out of range of a OAdate, a numeric value will be returned
        /// </summary>
        /// <param name="raw">Raw value as string</param>
        /// <param name="address">Address of the cell</param>
        /// <param name="styleNumber">Raw style number that may indicate whether the cell represents a date or time value</param>
        /// <returns>Cell of the type DateTime or the defined fall-back type</returns>
        private static Cell GetDateValue(String raw, Address address, string styleNumber)
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
                    DateTime date = DateTime.FromOADate(dValue);
                    switch (styleNumber)
                    {
                        case "1": // Possibly a Date
                            return new Cell(date, Cell.CellType.DATE, address);
                        case "3": // Possibly a Time
                            return new Cell(date.TimeOfDay, Cell.CellType.DATE, address); // TODO: Define TIME as type (must be implemented in the writer as well)
                        default:
                            return new Cell(date, Cell.CellType.DATE, address); // Currently duplicate of "1", as long no style reader is implemented
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
