/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2021
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
        /// <param name="styleReaderContainer">Resolved styles, used to determine dates or times</param>
        /// <param name="options">Import options to override the automatic approach of the reader. <see cref="ImportOptions"/> for information about import options.</param>
        public WorksheetReader(SharedStringsReader sharedStrings, string name, int number, StyleReaderContainer styleReaderContainer, ImportOptions options = null)
        {
            importOptions = options;
            Data = new Dictionary<string, Cell>();
            Name = name;
            WorksheetNumber = number;
            this.sharedStrings = sharedStrings;
            ProcessStyles(styleReaderContainer);
        }

        #endregion

        #region functions

        /// <summary>
        /// Determine which of the resolved styles are either to define a time or a date. Stores also the styles into a dictionary 
        /// </summary>
        /// <param name="styleReaderContainer">Resolved styles from the style reader</param>
        private void ProcessStyles(StyleReaderContainer styleReaderContainer)
        {
            dateStyles = new List<string>();
            timeStyles = new List<string>();
            for (int i = 0; i < styleReaderContainer.StyleCount; i++)
            {
                bool isDate;
                bool isTime;
                styleReaderContainer.GetStyle(i, out isDate, out isTime, true);
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
                    string type;
                    string styleNumber;
                    string address;
                    string value;
                    string formula;
                    XmlDocument xr = new XmlDocument();
                    xr.XmlResolver = null;
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
                                if (rowChild.LocalName.Equals("c", StringComparison.InvariantCultureIgnoreCase))
                                {
                                    address = ReaderUtils.GetAttribute("r", rowChild); // Mandatory
                                    type = ReaderUtils.GetAttribute("t", rowChild); // can be null if not existing
                                    styleNumber = ReaderUtils.GetAttribute("s", rowChild); // can be null
                                    if (rowChild.HasChildNodes)
                                    {
                                        foreach (XmlNode valueNode in rowChild.ChildNodes)
                                        {
                                            if (valueNode.LocalName.Equals("v", StringComparison.InvariantCultureIgnoreCase))
                                            {
                                                value = valueNode.InnerText;
                                            }
                                            if (valueNode.LocalName.Equals("f", StringComparison.InvariantCultureIgnoreCase))
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
                throw new IOException("The XML entry could not be read from the input stream. Please see the inner exception:", ex);
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
            string key = Utils.ToUpper(addressString);
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
        /// <param name="formula">Formula as string (can be null; data type determines whether value or formula is used)</param>
        /// <returns>The resolved Cell</returns>
        private Cell ResolveCellDataConditionally(Address address, string type, string value, string styleNumber, string formula)
        {
            if (address.Row < importOptions.EnforcingStartRowNumber)
            {
                return AutoResolveCellData(address, type, value, styleNumber, formula); // Skip enforcing
            }
            Cell tempCell = AutoResolveCellData(address, type, value, styleNumber, formula);
            IConvertible converter = tempCell.Value as IConvertible;
            switch (importOptions.GlobalEnforcingType)
            {
                case ImportOptions.GlobalType.AllNumbersToDouble:
                    if (tempCell.DataType == Cell.CellType.NUMBER)
                    {
                        return new Cell(converter.ToDouble(Utils.INVARIANT_CULTURE), tempCell.DataType, address);
                    }
                    else if (tempCell.DataType == Cell.CellType.BOOL)
                    {
                        double tempDouble = ((bool)tempCell.Value) ? 1d : 0d;
                        return new Cell(tempDouble, Cell.CellType.NUMBER, address);
                    }
                    else if (tempCell.DataType == Cell.CellType.DATE || tempCell.DataType == Cell.CellType.TIME)
                    {
                        converter = value as IConvertible;
                        double tempDouble = converter.ToDouble(Utils.INVARIANT_CULTURE);
                        return new Cell(tempDouble, Cell.CellType.NUMBER, address);
                    }
                    else if (tempCell.DataType == Cell.CellType.STRING)
                    {
                        double tempDouble;
                        if (double.TryParse(tempCell.Value.ToString(), out tempDouble)){
                            return new Cell(tempDouble, Cell.CellType.NUMBER, address);
                        }
                    }
                    return tempCell;
                case ImportOptions.GlobalType.AllNumbersToInt:
                    if (tempCell.DataType == Cell.CellType.NUMBER)
                    {
                        return new Cell(converter.ToInt32(Utils.INVARIANT_CULTURE), tempCell.DataType, address);
                    }
                    else if (tempCell.DataType == Cell.CellType.BOOL)
                    {
                        int tempint = ((bool)tempCell.Value) ? 1 : 0;
                        return new Cell(tempint, Cell.CellType.NUMBER, address);
                    }
                    else if (tempCell.DataType == Cell.CellType.DATE || tempCell.DataType == Cell.CellType.TIME)
                    {
                        converter = value as IConvertible;
                        double tempDouble = converter.ToDouble(Utils.INVARIANT_CULTURE);
                        return  new Cell((int)Math.Round(tempDouble, 0), Cell.CellType.NUMBER, address);
                    }
                    else if (tempCell.DataType == Cell.CellType.STRING)
                    {
                        int tempInt;
                        if (int.TryParse(tempCell.Value.ToString(), out tempInt))
                        {
                            return new Cell(tempInt, Cell.CellType.NUMBER, address);
                        }
                    }
                    return tempCell;
                case ImportOptions.GlobalType.EverythingToString:
                    return GetEnforcedStingValue(address, type, value, styleNumber, formula, importOptions);
            }
            
            if (string.IsNullOrEmpty(value) && string.IsNullOrEmpty(formula))
            {
                if (importOptions.EnforceEmptyValuesAsString)
                {
                    return new Cell("", Cell.CellType.STRING, address);
                }
                else
                {
                    return new Cell(null, Cell.CellType.EMPTY, address);
                }
            }
            if (importOptions.EnforcedColumnTypes.ContainsKey(address.Column))
            {
                ImportOptions.ColumnType importType = importOptions.EnforcedColumnTypes[address.Column];
                if (type == "s")
                {
                    // Resolve shared string first
                    value = ResolveSharedString(value);
                }
                switch (importType)
                {
                    case ImportOptions.ColumnType.Bool:
                         tempCell = GetBooleanValue(value, address);
                        if (tempCell == null)
                        {
                            return AutoResolveCellData(address, type, value, styleNumber, formula);
                        }
                        return tempCell;
                    case ImportOptions.ColumnType.Date:
                        if (!string.IsNullOrEmpty(formula))
                        {
                            return tempCell;
                        }
                        if (importOptions.EnforceDateTimesAsNumbers)
                        {
                            return GetNumericValue(value, address, styleNumber);
                        }
                        else
                        {
                            return GetDateTimeValue(value, address, Cell.CellType.DATE, type);
                        }
                    case ImportOptions.ColumnType.Time:
                        if (!string.IsNullOrEmpty(formula))
                        {
                            return tempCell;
                        }
                        if (importOptions.EnforceDateTimesAsNumbers)
                        {
                            return GetNumericValue(value, address, styleNumber);
                        }
                        else
                        {
                            return GetDateTimeValue(value, address, Cell.CellType.TIME, type);
                        }
                    case ImportOptions.ColumnType.Numeric:
                        return GetNumericValue(value, address);
                    case ImportOptions.ColumnType.Double:
                        return GetDoubleValue(value, address);
                    case ImportOptions.ColumnType.String:
                        return GetStringValue(value, address, type, styleNumber, importOptions);
                }
            }
            return AutoResolveCellData(address, type, value, styleNumber, formula);
        }

        /// <summary>
        /// Resolves the data of a read cell automatically, transforms it into a cell object
        /// </summary>
        /// <param name="address">Address of the cell</param>
        /// <param name="type">Expected data type</param>
        /// <param name="value">Raw value as string</param>
        /// <param name="styleNumber">Style number as string (can be null)</param>
        /// <param name="formula">Formula as string (can be null; data type determines whether value or formula is used)</param>
        /// <returns>Resolved Cell</returns>
        private Cell AutoResolveCellData(Address address, string type, string value, string styleNumber, string formula)
        {
            if (type != null && type == "s") // string (declared)
            {
                return GetStringValue(value, address, type);
            }
            else if (type == "b") // boolean
            {
                Cell tempCell = GetBooleanValue(value, address);
                if (tempCell == null)
                {
                    return AutoResolveCellData(address, null, value, styleNumber, formula);
                }
                return tempCell;
            }
            else if (dateStyles.Contains(styleNumber))  // date (priority)
            {
                return GetDateTimeValue(value, address, Cell.CellType.DATE);
            }
            else if (timeStyles.Contains(styleNumber)) // time
            {
                return GetDateTimeValue(value, address, Cell.CellType.TIME);
            }
            else if (type == null || type == "n") // try numeric if not parsed as date or time, before numeric
            {
                if (string.IsNullOrEmpty(value))
                {
                    return new Cell(null, Cell.CellType.EMPTY, address);
                }
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
        /// Handles the value of a raw cell as string. An appropriate formatting is applied to DateTime and TimeSpan values. Null values are left on type EMPTY 
        /// </summary>
        /// <param name="address">Address of the cell</param>
        /// <param name="type">Expected data type</param>
        /// <param name="value">Raw value as string</param>
        /// <param name="styleNumber">Style number as string (can be null)</param>
        /// <param name="formula">Formula as string (can be null; data type determines whether value or formula is used)</param>
        /// <param name="options">Options instance to determine appropriate formatting information</param>
        /// <returns>Cell of the type string</returns>
        private Cell GetEnforcedStingValue(Address address, string type, string value, string styleNumber, string formula , ImportOptions options)
        {
            Cell parsed = AutoResolveCellData(address, type, value, styleNumber, formula);
            if (parsed.DataType == Cell.CellType.EMPTY)
            {
                return parsed;
            }
            else if (parsed.DataType == Cell.CellType.DATE)
            {
                return GetStringValue(((DateTime)parsed.Value).ToString(options.DateTimeFormat), address);
            }
            else if (parsed.DataType == Cell.CellType.TIME)
            {
                return GetStringValue(((TimeSpan)parsed.Value).ToString(options.TimeSpanFormat), address);
            }
            else
            {
                return GetStringValue(parsed.Value.ToString(), address);
            }
        }

        /// <summary>
        /// Parses the numeric value of a raw cell. The order of possible number types are: ulong, long, uint, int, float or double. If nothing applies, a string is returned
        /// </summary>
        /// <param name="raw">Raw value as string</param>
        /// <param name="address">Address of the cell</param>
        /// <param name="styleNumber">Optional parameter to determine whether a double is enforced in case of a date or time style</param>
        /// <returns>Cell of the type int, float, double or string as fall-back type</returns>
        private Cell GetNumericValue(string raw, Address address, string styleNumber = null)
        {
            if (dateStyles.Contains(styleNumber) || timeStyles.Contains(styleNumber))
            {
                return GetDoubleValue(raw, address);
            }
            uint uiValue;
            int iValue;
            bool canBeUint = uint.TryParse(raw, out uiValue);
            bool canBeInt = int.TryParse(raw, out iValue);             
            if (canBeUint && !canBeInt)
            {
                return new Cell(uiValue, Cell.CellType.NUMBER, address);
            }
            else if (canBeInt)
            {
                return new Cell(iValue, Cell.CellType.NUMBER, address);
            }
            ulong ulValue;
            long lValue;
            bool canBeUlong = ulong.TryParse(raw, out ulValue);
            bool canBeLong = long.TryParse(raw, out lValue);
            if (canBeUlong && !canBeLong)
            {
                return new Cell(ulValue, Cell.CellType.NUMBER, address);
            }
            else if (canBeLong)
            {
                return new Cell(lValue, Cell.CellType.NUMBER, address);
            }

            float fValue;

            if (float.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out fValue))
            {
                if (!float.IsInfinity(fValue))
                {
                    return new Cell(fValue, Cell.CellType.NUMBER, address);
                }
            }
            return GetDoubleValue(raw, address);
        }

        /// <summary>
        /// Parses a raw value as double 
        /// </summary>
        /// <param name="raw">Raw value as string</param>
        /// <param name="address">Address of the cell</param>
        /// <returns>Cell of the type double or string as fall-back type</returns>
        private Cell GetDoubleValue(string raw, Address address)
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
        /// <param name="type">Optional parameter to check whether the raw value is already a resolved string</param>
        /// <param name="options">Optional import options to determine the date and time formatting information</param>
        /// <returns>Cell of the type string</returns>
        private Cell GetStringValue(string raw, Address address, string type = null, string styleNumber = null, ImportOptions options = null)
        {
            if (type != null && type == "s")
            {
                return new Cell(ResolveSharedString(raw), Cell.CellType.STRING, address);
            }
            else if (type != null && type == "b")
            {
                Cell tempCell = GetBooleanValue(raw, address);
                if (tempCell != null)
                {
                    return new Cell(tempCell.Value.ToString(), Cell.CellType.STRING, address);
                }
            }
            else if (styleNumber != null)
            {
                Cell tempCell = null;
                if (dateStyles.Contains(styleNumber))  // date (priority)
                {
                    tempCell = GetDateTimeValue(raw, address, Cell.CellType.DATE);
                }
                else if (timeStyles.Contains(styleNumber)) // time
                {
                    tempCell = GetDateTimeValue(raw, address, Cell.CellType.TIME);
                }
                if (tempCell != null && tempCell.DataType == Cell.CellType.DATE)
                {
                    return new Cell(((DateTime)tempCell.Value).ToString(options.DateTimeFormat), Cell.CellType.STRING, address);
                }
                else if (tempCell != null && tempCell.DataType == Cell.CellType.TIME)
                {
                    return new Cell(((TimeSpan)tempCell.Value).ToString(options.TimeSpanFormat), Cell.CellType.STRING, address);
                }
            }
            return new Cell(raw, Cell.CellType.STRING, address);
                                 
        }

        /// <summary>
        /// Tries to resolve a shared string from its ID
        /// </summary>
        /// <param name="raw">Raw value that can be either an ID of a shared string or an actual string value</param>
        /// <returns>Resolved string or the raw value if no shared string could be determined</returns>
        private string ResolveSharedString(string raw)
        {
            int stringId;
            if (int.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out stringId))
            {
                string resolvedString = sharedStrings.GetString(stringId);
                if (resolvedString == null)
                {
                    return raw;
                }
                else
                {
                    return resolvedString;
                }
            }
            return raw;
        }

        /// <summary>
        /// Parses the boolean value of a raw cell
        /// </summary>
        /// <param name="raw">Raw value as string</param>
        /// <param name="address">Address of the cell</param>
        /// <returns>Cell of the type bool or null if not able to parse</returns>
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
                    return null;
                }
            }
        }

        /// <summary>
        /// Parses the date (DateTime) or time (TimeSpan) value of a raw cell. If the value is numeric, but out of range of a OAdate, a numeric value will be returned instead. 
        /// If invalid, the string representation will be returned.
        /// </summary>
        /// <param name="raw">Raw value as string</param>
        /// <param name="address">Address of the cell</param>
        /// <param name="valueType">Type of the value to be converted: Valid values are DATE and TIME</param>
        /// <param name="type">Optional parameter to check whether the raw value should be tried to be parsed as date or time</param>
        /// <returns>Cell of the type TimeSpan or the defined fall-back type</returns>
        private Cell GetDateTimeValue(String raw, Address address, Cell.CellType valueType, string type = null)
        {
            if (type != null && type == "b")
            {
                return GetBooleanValue(raw, address);
            }
            if (type != null && type == "s")
            {
                DateTime? tempDate = TryParseDate(raw);
                if (tempDate != null && valueType == Cell.CellType.DATE)
                {
                    return GetTemporalCell(tempDate.Value, address);
                }
                else if (tempDate != null && valueType == Cell.CellType.TIME)
                {
                    return GetTemporalCell(new TimeSpan(tempDate.Value.Hour, tempDate.Value.Minute, tempDate.Value.Second), address);
                }
                TimeSpan? tempTime = TryParseTime(raw);
                if (tempTime != null && valueType == Cell.CellType.TIME)
                {
                    return GetTemporalCell(tempTime.Value, address);
                }
            }
            double dValue;
            bool isDouble = false;
            if (importOptions == null || string.IsNullOrEmpty(importOptions.DateTimeFormat) || importOptions.TemporalCultureInfo == null)
            {
                isDouble = double.TryParse(raw, out dValue);
            }
            else
            {
                isDouble = double.TryParse(raw, NumberStyles.Any, importOptions.TemporalCultureInfo, out dValue);
            }
            if (!isDouble)
            {
                return new Cell(raw, Cell.CellType.STRING, address);
            }
            if (dValue < Utils.MIN_OADATE_VALUE || dValue > Utils.MAX_OADATE_VALUE || (importOptions != null && importOptions.EnforceDateTimesAsNumbers))
            {
                return new Cell(dValue, Cell.CellType.NUMBER, address); // Invalid OAdate / enforced number == plain number
            }
            else if (valueType == Cell.CellType.DATE)
            {
                DateTime date = Utils.GetDateFromOA(dValue);
                if (date >= Utils.FIRST_ALLOWED_EXCEL_DATE)
                {
                    return GetTemporalCell(date, address);
                }
                else
                {
                    // Prevent to import 00.01.1900, since it will lead to trouble when exporting / writing
                    return new Cell(dValue, Cell.CellType.NUMBER, address);
                }
            }
            else if (valueType == Cell.CellType.TIME)
            {
                TimeSpan time = TimeSpan.FromSeconds(dValue * 86400d);
                return GetTemporalCell(time, address);
            }
            return new Cell(raw, Cell.CellType.STRING, address);
        }

        /// <summary>
        /// Gets a cell either as DateTime, TimeSpan or as double if enforced by import options
        /// </summary>
        /// <param name="dateTimeValue">Value of the cell</param>
        /// <param name="address">Address of the cell</param>
        /// <returns>Casted cell</returns>
        private Cell GetTemporalCell(Object dateTimeValue, Address address)
        {
            if (importOptions != null && importOptions.EnforceDateTimesAsNumbers)
            {
                if (dateTimeValue is DateTime)
                {
                    return new Cell(Utils.GetOADateTime((DateTime)dateTimeValue), Cell.CellType.NUMBER, address);
                }
                else
                {
                    return new Cell(Utils.GetOATime((TimeSpan)dateTimeValue), Cell.CellType.NUMBER, address);
                }
            }
            if (dateTimeValue is DateTime)
            {
                return new Cell((DateTime)dateTimeValue, Cell.CellType.DATE, address);
            }
            else
            {
                return new Cell((TimeSpan)dateTimeValue, Cell.CellType.TIME, address);
            }
        }

        /// <summary>
        /// Tris to parse a DateTime instance from a string
        /// </summary>
        /// <param name="raw">String to parse</param>
        /// <returns>DateTime instance or null if not possible to parse</returns>
        private DateTime? TryParseDate(string raw)
        {
            DateTime dateTime;
            bool isDateTime = false;
            if (importOptions == null || string.IsNullOrEmpty(importOptions.DateTimeFormat) || importOptions.TemporalCultureInfo == null)
            {
                isDateTime = DateTime.TryParse(raw, out dateTime);
            }
            else
            {
                isDateTime = DateTime.TryParseExact(raw, importOptions.DateTimeFormat, importOptions.TemporalCultureInfo, DateTimeStyles.None, out dateTime);
            }
            if (isDateTime && dateTime >= Utils.FIRST_ALLOWED_EXCEL_DATE && dateTime <= Utils.LAST_ALLOWED_EXCEL_DATE)
            {
                return dateTime;
            }
            return null;
        }

        /// <summary>
        /// Tris to parse a TimeSpan instance from a string
        /// </summary>
        /// <param name="raw">String to parse</param>
        /// <returns>TimeSpan instance or null if not possible to parse</returns>
        private TimeSpan? TryParseTime(string raw)
        {
            TimeSpan timeSpan;
            bool isTimeSpan = false;
            if (importOptions == null || string.IsNullOrEmpty(importOptions.TimeSpanFormat) || importOptions.TemporalCultureInfo == null)
            {
                isTimeSpan = TimeSpan.TryParse(raw, out timeSpan);
            }
            else
            {
                isTimeSpan = TimeSpan.TryParseExact(raw, importOptions.TimeSpanFormat, importOptions.TemporalCultureInfo, out timeSpan);
            }
            if (isTimeSpan)
            {
                return timeSpan;
            }
            return null;
        }

        #endregion

    }
}
