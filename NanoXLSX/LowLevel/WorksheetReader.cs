/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2019
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
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
        public WorksheetReader(SharedStringsReader sharedStrings, string name, int number)
        {
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
        /// Resolves the data of a read cell, transforms it into a cell object and adds it to the data
        /// </summary>
        /// <param name="address">Address of the cell</param>
        /// <param name="type">Expected data type</param>
        /// <param name="value">Raw value as string</param>
        /// <param name="style">Style definition as string (can be null)</param>
        /// <param name="formula"> Formula as string (can be null; data type determines whether value or formula is used)</param>
        private void ResolveCellData(string address, string type, string value, string style, string formula)
        {
            address = address.ToUpper();
            string s;
            Cell cell;
            CellResolverTuple tuple;
            if (style != null && style == "1") // Date must come before numeric values
            {
                tuple = GetDateValue(value);
                if (tuple.IsValid)
                {
                    cell = new Cell(tuple.Data, Cell.CellType.DATE, address);
                }
                else
                {
                    cell = new Cell(value, Cell.CellType.STRING, address);
                }
            }
            else if (type == null) // try numeric
            {
                tuple = GetNumericValue(value);
                if (tuple.IsValid)
                {
                    cell = new Cell(tuple.Data, Cell.CellType.NUMBER, address);
                }
                else
                {
                    cell = new Cell(value, Cell.CellType.STRING, address);
                }
            }
            else if (type == "b")
            {
                tuple = GetBooleanValue(value);
                if (tuple.IsValid)
                {
                    cell = new Cell(tuple.Data, Cell.CellType.BOOL, address);
                }
                else
                {
                    cell = new Cell(value, Cell.CellType.STRING, address);
                }
            }
            else if (formula != null) // formula before string
            {
                cell = new Cell(formula, Cell.CellType.FORMULA, address);
            }
            else if (type == "s")
            {
                tuple = GetIntValue(value);
                if (tuple.IsValid == false)
                {
                    cell = new Cell(value, Cell.CellType.STRING, address);
                }
                else
                {
                    s = sharedStrings.GetString((int)tuple.Data);
                    if (s != null)
                    {
                        cell = new Cell(s, Cell.CellType.STRING, address);
                    }
                    else
                    {
                        cell = new Cell(value, Cell.CellType.STRING, address);
                    }
                }
            }
            else
            {
                cell = new Cell(value, Cell.CellType.STRING, address);
            }
            Data.Add(address, cell);
        }


        /// <summary>
        /// Parses the numeric (double) value of a raw cell
        /// </summary>
        /// <param name="raw">Raw value as string</param>
        /// <returns>CellResolverTuple with information about the validity and resolved data</returns>
        private static CellResolverTuple GetNumericValue(string raw)
        {
            double dValue;
            CellResolverTuple t;
            if (double.TryParse(raw, out dValue))
            {
                t = new CellResolverTuple(true, dValue, typeof(double));
            }
            else
            {
                t = new CellResolverTuple(false, 0, typeof(double));
            }
            return t;
        }

        /// <summary>
        /// Parses the integer value of a raw cell
        /// </summary>
        /// <param name="raw">Raw value as string</param>
        /// <returns>CellResolverTuple with information about the validity and resolved data</returns>
        private static CellResolverTuple GetIntValue(string raw)
        {
            int iValue;
            CellResolverTuple t;
            if (int.TryParse(raw, out iValue))
            {
                t = new CellResolverTuple(true, iValue, typeof(int));
            }
            else
            {
                t = new CellResolverTuple(false, 0, typeof(int));
            }
            return t;
        }

        /// <summary>
        /// Parses the boolean value of a raw cell
        /// </summary>
        /// <param name="raw">Raw value as string</param>
        /// <returns>CellResolverTuple with information about the validity and resolved data</returns>
        private static CellResolverTuple GetBooleanValue(String raw)
        {
            bool value;
            bool state;
            if (raw == "0")
            {
                value = false;
                state = true;
            }
            else if (raw == "1")
            {
                value = true;
                state = true;
            }
            else
            {
                if (bool.TryParse(raw, out value))
                {
                    state = true;
                }
                else
                {
                    state = false;
                    value = false;
                }
            }
            return new CellResolverTuple(state, value, typeof(bool));
        }

        /// <summary>
        /// Parses the date (DateTime) value of a raw cell
        /// </summary>
        /// <param name="raw">Raw value as string</param>
        /// <returns>CellResolverTuple with information about the validity and resolved data</returns>
        private static CellResolverTuple GetDateValue(String raw)
        {
            double dValue;
            CellResolverTuple t;
            if (double.TryParse(raw, out dValue))
            {
                DateTime date = DateTime.FromOADate(dValue);
                t = new CellResolverTuple(true, date, typeof(DateTime));
            }
            else
            {
                t = new CellResolverTuple(false, new DateTime(), typeof(DateTime));
            }
            return t;
        }

        #endregion

        #region subClasses

        /// <summary>
        /// Helper class representing a tuple of cell data and is state (valid or invalid). And additional type is also available
        /// </summary>
        class CellResolverTuple
        {
            /// <summary>
            /// Gets whether the cell is valid
            /// </summary>
            /// <value>
            ///   True if valid, otherwise false
            /// </value>
            public bool IsValid { get; private set; }

            /// <summary>
            /// Gets the data as object
            /// </summary>
            /// <value>
            /// Generic object
            /// </value>
            public object Data { get; private set; }

            /// <summary>
            /// Gets the type of the cell
            /// </summary>
            /// <value>
            /// Data type
            /// </value>
            public Type DataType { get; private set; }

            /// <summary>
            /// Default constructor with parameters
            /// </summary>
            /// <param name="isValid">If true, the resolved cell contains valid data</param>
            /// <param name="data">Data object.</param>
            /// <param name="type">Type of the cell</param>
            public CellResolverTuple(bool isValid, object data, Type type)
            {
                Data = data;
                IsValid = isValid;
                DataType = type;
            }

        }

        #endregion

    }
}
