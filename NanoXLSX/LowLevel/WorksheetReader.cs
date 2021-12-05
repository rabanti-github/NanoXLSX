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
        private Dictionary<string, Style> resolvedStyles;
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
            resolvedStyles = new Dictionary<string, Style>();
            for (int i = 0; i < styleReaderContainer.StyleCount; i++)
            {
                bool isDate;
                bool isTime;
                string index = i.ToString("G", CultureInfo.InvariantCulture);
                Style style = styleReaderContainer.GetStyle(i, out isDate, out isTime, true);
                if (isDate)
                {
                    dateStyles.Add(index);
                }
                if (isTime)
                {
                    timeStyles.Add(index);
                }
                resolvedStyles.Add(index, style);
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
                                ReadCell(rowChild);
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
        /// Reads one cell in a worksheet
        /// </summary>
        /// <param name="rowChild">Current child row as XmlNode</param>
        private void ReadCell(XmlNode rowChild)
        {
            string type = "s";
            string styleNumber = "";
            string address = "A1";
            string value = "";
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
                            value = valueNode.InnerText;
                        }
                    }
                }
            }
            string key = Utils.ToUpper(address);
            StyleAssignment[key] = styleNumber;
            Data.Add(key, ResolveCellData(value, type, styleNumber, address));
        }

        /// <summary>
        /// Resolves the data of a read cell either automatically or conditionally  (import options), transforms it into a cell object and adds it to the data
        /// </summary>
        /// <param name="raw">Raw value as string</param>
        /// <param name="type">Expected data type</param>
        /// <param name="styleNumber">>Style number as string (can be null)</param>
        /// <param name="address">Address of the cell</param>
        /// <returns>Cell object with either the originally loaded or modified (by import options) value</returns>
        private Cell ResolveCellData(string raw, string type, string styleNumber, string address)
        {
            Cell.CellType importedType = Cell.CellType.DEFAULT;
            object rawValue;
            if (type == "b")
            {
                rawValue = TryParseBool(raw);
                if (rawValue != null)
                {
                    importedType = Cell.CellType.BOOL;
                }
                else
                {
                    rawValue = GetNumericValue(raw);
                    if (rawValue != null)
                    {
                        importedType = Cell.CellType.NUMBER;
                    }
                }
            }
            else if (type == "s")
            {
                importedType = Cell.CellType.STRING;
                rawValue = ResolveSharedString(raw);
            }
            else if (type == "str")
            {
                importedType = Cell.CellType.FORMULA;
                rawValue = raw;
            }
            else if (dateStyles.Contains(styleNumber) && (type == null || type == "" || type == "n"))
            {
                rawValue = GetDateTimeValue(raw, Cell.CellType.DATE, out importedType);
            }
            else if (timeStyles.Contains(styleNumber) && (type == null || type == "" || type == "n"))
            {
                rawValue = GetDateTimeValue(raw, Cell.CellType.TIME, out importedType);
            }
            else
            {
                importedType = Cell.CellType.NUMBER;
                rawValue = GetNumericValue(raw);
            }
            if (rawValue == null && raw == "")
            {
                importedType = Cell.CellType.EMPTY;
                rawValue = null;
            }
            else if (rawValue == null && raw.Length > 0)
            {
                importedType = Cell.CellType.STRING;
                rawValue = raw;
            }
            Address cellAddress = new Address(address);
            if (importOptions != null)
            {
                if (importOptions.EnforcedColumnTypes.Count > 0)
                {
                    rawValue = GetEnforcedColumnValue(rawValue, importedType, cellAddress);
                }
                rawValue = GetGloballyEnforcedValue(rawValue, cellAddress);
                rawValue = GetGloballyEnforcedFlagValues(rawValue, cellAddress);
                importedType = ResolveType(rawValue, importedType);
                if (importedType == Cell.CellType.DATE && rawValue is DateTime && (DateTime)rawValue < Utils.FIRST_ALLOWED_EXCEL_DATE)
                {
                    // Fix conversion from time to date, where time has no days
                    rawValue = ((DateTime)rawValue).AddDays(1);
                }
            }
            return CreateCell(rawValue, importedType, cellAddress, styleNumber);
        }

        /// <summary>
        /// Resolves the final cell type after a possible conversion by import options
        /// </summary>
        /// <param name="value">Value of the cell</param>
        /// <param name="defaultType">Originally resolved type. If a formula, the method immediately returns</param>
        /// <returns>Resolved cell type</returns>
        private Cell.CellType ResolveType(object value, Cell.CellType defaultType)
        {
            if (defaultType == Cell.CellType.FORMULA)
            {
                return defaultType;
            }
            if (value == null)
            {
                return Cell.CellType.EMPTY;
            }
            switch (value)
            {
                case uint _:
                case long _:
                case ulong _:
                case short _:
                case ushort _:
                case float _:
                case byte _:
                case sbyte _:
                case int _:
                    return Cell.CellType.NUMBER;
                case string _:
                    return Cell.CellType.STRING;
                case DateTime _:
                    return Cell.CellType.DATE;
                case TimeSpan _:
                    return Cell.CellType.TIME;
                case bool _:
                    return Cell.CellType.BOOL;
            }
            return defaultType;
        }

        /// <summary>
        /// Modifies certain values globally by import options (e.g. empty as string or dates as numbers)
        /// </summary>
        /// <param name="data">Cell data</param>
        /// <param name="address">Cell address (conversion is skipped if start row is not reached)</param>
        /// <returns>Modified value</returns>
        private object GetGloballyEnforcedFlagValues(object data, Address address)
        {
            if (address.Row < importOptions.EnforcingStartRowNumber)
            {
                return data;
            }
            if (importOptions.EnforceDateTimesAsNumbers)
            {
                if (data is DateTime)
                {
                    data = Utils.GetOADateTime((DateTime)data, true);
                }
                else if (data is TimeSpan)
                {
                    data = Utils.GetOATime((TimeSpan)data);
                }
            }
            if (importOptions.EnforceEmptyValuesAsString)
            {
                if (data == null)
                {
                    return "";
                }
            }
            return data;
        }

        /// <summary>
        /// Converts the cell values globally, based on import options (e.g. everything to string or all numbers to double)
        /// </summary>
        /// <param name="data">Cell data</param>
        /// <param name="address">>Cell address (conversion is skipped if start row is not reached)</param>
        /// <returns>Converted value</returns>
        private object GetGloballyEnforcedValue(object data, Address address)
        {
            if (address.Row < importOptions.EnforcingStartRowNumber)
            {
                return data;
            }
            if (importOptions.GlobalEnforcingType == ImportOptions.GlobalType.AllNumbersToDouble)
            {
                object tempDouble = ConvertToDouble(data);
                if (tempDouble != null)
                {
                    return tempDouble;
                }
            }
            else if (importOptions.GlobalEnforcingType == ImportOptions.GlobalType.AllNumbersToInt)
            {
                object tempInt;
                if (ConvertToInt(data, out tempInt))
                {
                    return tempInt;
                }
            }
            else if (importOptions.GlobalEnforcingType == ImportOptions.GlobalType.EverythingToString)
            {
                return ConvertToString(data);
            }
            return data;
        }

        /// <summary>
        /// Converts the cell values of defined rows, based on import options (e.g. everything to string or all values to double)
        /// </summary>
        /// <param name="data"></param>
        /// <param name="importedTyp"></param>
        /// <param name="address"></param>
        /// <returns></returns>
        private object GetEnforcedColumnValue(object data, Cell.CellType importedTyp, Address address)
        {
            if (address.Row < importOptions.EnforcingStartRowNumber)
            {
                return data;
            }
            if (!importOptions.EnforcedColumnTypes.ContainsKey(address.Column))
            {
                return data;
            }
            if (importedTyp == Cell.CellType.FORMULA)
            {
                return data;
            }
            switch (importOptions.EnforcedColumnTypes[address.Column])
            {
                case ImportOptions.ColumnType.Numeric:
                    return GetNumericValue(data, importedTyp);
                case ImportOptions.ColumnType.Double:
                        return ConvertToDouble(data);
                case ImportOptions.ColumnType.Date:
                    return ConvertToDate(data);
                case ImportOptions.ColumnType.Time:
                    return ConvertToTime(data);
                case ImportOptions.ColumnType.Bool:
                    return ConvertToBool(data);
                default:
                    return ConvertToString(data);
            }
        }

        /// <summary>
        /// Tries to convert a value to a bool
        /// </summary>
        /// <param name="data">Raw data</param>
        /// <returns>Bool value or original value if not possible to convert</returns>
        private object ConvertToBool(object data)
        {
            switch (data)
            {
                case bool _:
                    return data;
                case uint _:
                case long _:
                case ulong _:
                case short _:
                case ushort _:
                case float _:
                case byte _:
                case sbyte _:
                case int _:
                    object tempObject = ConvertToDouble(data);
                    if (tempObject is double)
                    {
                        double tempDouble = (double)tempObject;
                        if (double.Equals(tempDouble, 0d))
                        {
                            return false;
                        }
                        else if (double.Equals(tempDouble, 1d))
                        {
                            return true;
                        }
                    }
                    break;
                case string _:
                    bool tempBool;
                    string tempString = (string)data;
                    if (bool.TryParse(tempString, out tempBool)){
                        return tempBool;
                    }
                    else if (tempString == "1")
                    {
                        return true;
                    }
                    else if (tempString == "0")
                    {
                        return false;
                    }
                    break;
            }
            return data;
        }

        /// <summary>
        /// Tries to convert a value to a double
        /// </summary>
        /// <param name="data">Raw data</param>
        /// <returns>Double value or original value if not possible to convert</returns>
        private object ConvertToDouble(object data)
        {
            switch (data)
            {
                case double _:
                    return data;
                case uint _:
                case long _:
                case ulong _:
                case short _:
                case ushort _:
                case float _:
                case byte _:
                case sbyte _:
                case int _:
                    IConvertible converter = data as IConvertible;
                    return converter.ToDouble(Utils.INVARIANT_CULTURE);
                case bool _:
                    if ((bool)data)
                    {
                        return 1d;
                    }
                    else
                    {
                        return 0d;
                    }
                case DateTime _:
                    return Utils.GetOADateTime((DateTime)data);
                case TimeSpan _:
                    return Utils.GetOATime((TimeSpan)data);
                case string _:
                    double dValue;
                    string tempString = (string)data;
                    if (double.TryParse(tempString, out dValue))
                    {
                        return dValue;
                    }
                    DateTime? tempDate = TryParseDate(tempString);
                    if (tempDate != null)
                    {
                        return Utils.GetOADateTime(tempDate.Value);
                    }
                    TimeSpan? tempTime = TryParseTime(tempString);
                    if (tempTime != null)
                    {
                        return Utils.GetOATime(tempTime.Value);
                    }
                    break;
            }
            return data;
        }

        /// <summary>
        /// Tries to convert a value to an integer
        /// </summary>
        /// <param name="data">Raw data</param>
        /// <returns>Integer value or original value if not possible to convert</returns>
        private bool ConvertToInt(object data, out object converted)
        {
            object tempValue;
            double tempDouble;
            switch (data)
            {
                case uint _:
                case long _:
                case ulong _:
                    break;
                case DateTime _:
                    tempDouble = Utils.GetOADateTime((DateTime)data, true);
                    converted = ConvertDoubleToInt(tempDouble);
                    return true;
                case TimeSpan _:
                    tempDouble = Utils.GetOATime((TimeSpan)data);
                    converted = ConvertDoubleToInt(tempDouble);
                    return true;
                case float _:
                case double _:
                    if (TryConvertDoubleToInt(data, out tempValue))
                    {
                        converted = tempValue;
                        return true;
                    }
                    break;
                case bool _:
                    if ((bool)data)
                    {
                        converted = 1;
                    }
                    else
                    {
                        converted = 0;
                    }
                    return true;
                case string _:
                    int tempInt;
                    if (int.TryParse((string)data, out tempInt))
                    {
                        converted = tempInt;
                        return true;
                    }
                    break;
            }
            converted = null;
            return false;
        }

        /// <summary>
        /// Tries to convert a value to a Date (DateTime)
        /// </summary>
        /// <param name="data">Raw data</param>
        /// <returns>DateTime value or original value if not possible to convert</returns>
        private object ConvertToDate(object data)
        {
            switch (data)
            {
                case DateTime _:
                    return data;
                case TimeSpan _:
                    DateTime root = Utils.FIRST_ALLOWED_EXCEL_DATE;
                    TimeSpan time = (TimeSpan)data;
                    root = root.AddDays(-1); // Fix offset of 1
                    root = root.AddHours(time.Hours);
                    root = root.AddMinutes(time.Minutes);
                    root = root.AddSeconds(time.Seconds);
                    return root;
                case double _:
                case uint _:
                case long _:
                case ulong _:
                case short _:
                case ushort _:
                case float _:
                case byte _:
                case sbyte _:
                case int _:
                    return ConvertDateFromDouble(data);
                case string _:
                    DateTime date2;
                    if(DateTime.TryParseExact((string)data, importOptions.DateTimeFormat, importOptions.TemporalCultureInfo, DateTimeStyles.None, out date2))
                    {
                        if (date2 >= Utils.FIRST_ALLOWED_EXCEL_DATE && date2 <= Utils.LAST_ALLOWED_EXCEL_DATE)
                        {
                            return date2;
                        }
                    }
                    return ConvertDateFromDouble(data);
            }
            return data;
        }

        /// <summary>
        /// Tries to convert a value to a Time (TimeSpan)
        /// </summary>
        /// <param name="data">Raw data</param>
        /// <returns>TimeSpan value or original value if not possible to convert</returns>
        private object ConvertToTime(object data)
        {
            switch (data)
            {
                case DateTime _:
                    return ConvertTimeFromDouble(data);
                case TimeSpan _:
                    return data;
                case double _:
                case uint _:
                case long _:
                case ulong _:
                case short _:
                case ushort _:
                case float _:
                case byte _:
                case sbyte _:
                case int _:
                    return ConvertTimeFromDouble(data);
                case string _:
                    TimeSpan time;
                    if(TimeSpan.TryParseExact((string)data, importOptions.TimeSpanFormat, importOptions.TemporalCultureInfo, out time))
                    {
                        return time;
                    }
                    return ConvertTimeFromDouble(data);
            }
            return data;
        }

        /// <summary>
        /// Tries to convert a date (DateTime) from a double
        /// </summary>
        /// <param name="data">Raw data (may not be a double)</param>
        /// <returns>DateTime value or original value if not possible to convert</returns>
        private object ConvertDateFromDouble(object data)
        {
            object oaDate = ConvertToDouble(data);
            if (oaDate is double && (double)oaDate < Utils.MAX_OADATE_VALUE)
            {
                DateTime date = Utils.GetDateFromOA((double)oaDate);
                if (date >= Utils.FIRST_ALLOWED_EXCEL_DATE && date <= Utils.LAST_ALLOWED_EXCEL_DATE)
                {
                    return date;
                }
            }
            return data;
        }

        /// <summary>
        /// Tries to convert a time (TimeSpan) from a double
        /// </summary>
        /// <param name="data">Raw data (my not be a double)</param>
        /// <returns>TimeSpan value or original value if not possible to convert</returns>
        private object ConvertTimeFromDouble(object data)
        {
            object oaDate = ConvertToDouble(data);
            if (oaDate is double)
            { double d = (double)oaDate;
                if (d >= Utils.MIN_OADATE_VALUE && d <= Utils.MAX_OADATE_VALUE)
                {
                    DateTime date = Utils.GetDateFromOA(d);
                    return new TimeSpan((int)d, date.Hour, date.Minute, date.Second);
                }
            }
            return data;
        }

        /// <summary>
        /// Tries to convert a double to an integer
        /// </summary>
        /// <param name="data">Numeric value (possibly integer)</param>
        /// <param name="converted">>Converted values as out parameter. If not possible to convert, the original value is returned</param>
        /// <returns>True if possible to convert, otherwise false</returns>
        private bool TryConvertDoubleToInt(object data, out object converted)
        {
            IConvertible converter = data as IConvertible;
            double dValue = converter.ToDouble(ImportOptions.DEFAULT_CULTURE_INFO);
            if (dValue > int.MinValue && dValue < int.MaxValue)
            {
                converted = converter.ToInt32(ImportOptions.DEFAULT_CULTURE_INFO);
                return true;
            }
            converted = data;
            return false;
        }

        /// <summary>
        /// Converts a double to an integer without checks
        /// </summary>
        /// <param name="data">Numeric value</param>
        /// <returns>Converted Value</returns>
        public object ConvertDoubleToInt(object data)
        {
            IConvertible converter = data as IConvertible;
            return converter.ToInt32(ImportOptions.DEFAULT_CULTURE_INFO);
        }

        /// <summary>
        /// Converts an arbitrary value to string 
        /// </summary>
        /// <param name="data">Raw data</param>
        /// <returns>Converted string or null in case of null as input</returns>
        private string ConvertToString(object data)
        {
            switch (data)
            {
                case int _:
                    return ((int)data).ToString(ImportOptions.DEFAULT_CULTURE_INFO);
                case uint _:
                    return ((uint)data).ToString(ImportOptions.DEFAULT_CULTURE_INFO);
                case long _:
                    return ((long)data).ToString(ImportOptions.DEFAULT_CULTURE_INFO);
                case ulong _:
                    return ((ulong)data).ToString(ImportOptions.DEFAULT_CULTURE_INFO);
                case float _:
                    return ((float)data).ToString(ImportOptions.DEFAULT_CULTURE_INFO);
                case double _:
                    return ((double)data).ToString(ImportOptions.DEFAULT_CULTURE_INFO);
                case bool _:
                    return ((bool)data).ToString(ImportOptions.DEFAULT_CULTURE_INFO);
                case DateTime _:
                    return ((DateTime)data).ToString(importOptions.DateTimeFormat);
                case TimeSpan _:
                    return ((TimeSpan)data).ToString(importOptions.TimeSpanFormat);
                default:
                    if (data == null)
                    {
                        return null;
                    }
                    return data.ToString();
            }
        }

        /// <summary>
        /// Tries to parse a numeric value with an appropriate type
        /// </summary>
        /// <param name="raw">Raw value</param>
        /// <param name="importedType">Originally resolved cell type</param>
        /// <returns>Converted value or the raw value if not possible to convert</returns>
        private object GetNumericValue(object raw, Cell.CellType importedType)
        {
            if (raw == null)
            {
                return null;
            }
            object tempObject;
            switch (importedType)
            {
                case Cell.CellType.STRING:
                    string tempString = raw.ToString();
                    tempObject = GetNumericValue(tempString);
                    if (tempObject != null)
                    {
                        return tempObject;
                    }
                    DateTime? tempDate = TryParseDate(tempString);
                    if (tempDate != null)
                    {
                        return Utils.GetOADateTime(tempDate.Value);
                    }
                    TimeSpan? tempTime = TryParseTime(tempString);
                    if (tempTime != null)
                    {
                        return Utils.GetOATime(tempTime.Value);
                    }
                    tempObject = ConvertToBool(raw);
                    if (tempObject is bool)
                    {
                        return (bool)tempObject ? 1 : (object)0;
                    }
                    break;
                case Cell.CellType.NUMBER:
                    return raw;
                case Cell.CellType.DATE:
                    return Utils.GetOADateTime((DateTime)raw);
                case Cell.CellType.TIME:
                    return Utils.GetOATime((TimeSpan)raw);
                case Cell.CellType.BOOL:
                    if ((bool)raw){
                        return 1;
                    }
                    return 0;
            }
            return raw;
        }


        /// <summary>
        /// Parses the numeric value of a raw cell. The order of possible number types are: ulong, long, uint, int, float or double. If nothing applies, null is returned
        /// </summary>
        /// <param name="raw">Raw value as string</param>
        /// <returns>Value of the type int, float, double or null as fall-back</returns>
        private object GetNumericValue(string raw)
        {
            // integer section
            uint uiValue;
            int iValue;
            bool canBeUint = uint.TryParse(raw, out uiValue);
            bool canBeInt = int.TryParse(raw, out iValue);
            if (canBeUint && !canBeInt)
            {
                return uiValue;
            }
            else if (canBeInt)
            {
                return iValue;
            }
            ulong ulValue;
            long lValue;
            bool canBeUlong = ulong.TryParse(raw, out ulValue);
            bool canBeLong = long.TryParse(raw, out lValue);
            if (canBeUlong && !canBeLong)
            {
                return  ulValue;
            }
            else if (canBeLong)
            {
                return lValue;
            }
            decimal dcValue;
            double dValue;
            float fValue;
            // float section
            if (decimal.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out dcValue))
            {
                int decimals = BitConverter.GetBytes(decimal.GetBits(dcValue)[3])[2];
                if (decimals < 7)
                {
                    return decimal.ToSingle(dcValue);
                }
                else
                {
                    return decimal.ToDouble(dcValue);
                }
            }
            // High range float section
            else if (float.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out fValue) && fValue >= float.MinValue && fValue <= float.MaxValue && !float.IsInfinity(fValue))
            {
                return fValue;
            }
            if (double.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out dValue))
            {
                    return dValue;
            }
            return null;
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
        /// <returns>Object of the type bool or null if not able to parse</returns>
        private bool? TryParseBool(string raw)
        {
            if (raw == "0")
            {
                return false;
            }
            else if (raw == "1")
            {
                return true;
            }
            else
            {
                bool value;
                if (bool.TryParse(raw, out value))
                {
                    return value;
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
        /// <param name="valueType">Type of the value to be converted: Valid values are DATE and TIME</param>
        /// <returns>Object of the type TimeSpan or null if not possible to parse</returns>
        private object GetDateTimeValue(string raw, Cell.CellType valueType, out Cell.CellType resolvedType)
        {
            double dValue;
            if (!double.TryParse(raw, out dValue))
            {
                resolvedType = Cell.CellType.STRING;
                return raw;
            }
            if ((valueType == Cell.CellType.DATE && ( dValue < Utils.MIN_OADATE_VALUE || dValue > Utils.MAX_OADATE_VALUE)) || (valueType == Cell.CellType.TIME && (dValue < 0.0 || dValue >Utils.MAX_OADATE_VALUE)))
            {
                // fallback to number (cannot be anything else)
                resolvedType = Cell.CellType.NUMBER;
                return GetNumericValue(raw);
            }
            DateTime tempDate = Utils.GetDateFromOA(dValue);
            if (dValue < 1.0)
            {
                tempDate = tempDate.AddDays(1); // Modify wrong 1st date when < 1
            }
            if (valueType == Cell.CellType.DATE)
            {
                resolvedType = Cell.CellType.DATE;
                return tempDate;
            }
            else 
            {
                resolvedType = Cell.CellType.TIME;
                return new TimeSpan((int)dValue, tempDate.Hour, tempDate.Minute, tempDate.Second);
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

        /// <summary>
        /// Creates a generic cell with optional style information
        /// </summary>
        /// <param name="value">Value of the cell</param>
        /// <param name="type">Cell type</param>
        /// <param name="address">Cell address</param>
        /// <param name="styleNumber">Optional style number of the cell</param>
        /// <returns>Resolved cell</returns>
        private Cell CreateCell(object value, Cell.CellType type, Address address, string styleNumber = null)
        {
            Cell cell = new Cell(value, type, address);
            if (styleNumber != null && resolvedStyles.ContainsKey(styleNumber))
            {
                cell.SetStyle(resolvedStyles[styleNumber]);
            }
            return cell;
        }

        #endregion

    }
}
