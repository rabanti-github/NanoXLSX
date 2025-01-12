/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Styles;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Xml;
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

        /// <summary>
        /// Gets the auto filter range. If null, no auto filters were defined
        /// </summary>
        public Range? AutoFilterRange { get; private set; }
        /// <summary>
        /// Gets a list of defined Columns
        /// </summary>
        public List<Column> Columns { get; private set; } = new List<Column>();

        /// <summary>
        /// Gets the default column width if defined, otherwise null
        /// </summary>
        public float? DefaultColumnWidth { get; private set; } = null;

        /// <summary>
        /// Gets the default row height if defined, otherwise null
        /// </summary>
        public float? DefaultRowHeight { get; private set; } = null;

        /// <summary>
        /// Gets a dictionary of internal Row definitions
        /// </summary>
        public Dictionary<int, RowDefinition> Rows { get; private set; } = new Dictionary<int, RowDefinition>();
        /// <summary>
        /// Gets a list of merged cells
        /// </summary>
        public List<Range> MergedCells { get; private set; } = new List<Range>();

        /// <summary>
        /// Gets the selected cells (panes are currently not considered)
        /// </summary>
        public List<Range> SelectedCells { get; private set; } = new List<Range>();

        /// <summary>
        /// Gets the applicable worksheet protection values
        /// </summary>
        public Dictionary<Worksheet.SheetProtectionValue, int> WorksheetProtection { get; private set; } = new Dictionary<Worksheet.SheetProtectionValue, int>();

        /// <summary>
        /// Gets the (legacy) password hash of a worksheet if protection values are applied with a password
        /// </summary>
        public string WorksheetProtectionHash { get; private set; }

        /// <summary>
        /// Gets the definition of pane split-related information 
        /// </summary>
        public PaneDefinition PaneSplitValue { get; private set; }

        /// <summary>
        /// Gets whether grid lines are shown
        /// </summary>
        public bool ShowGridLines { get; private set; } = true; // default

        /// <summary>
        /// Gets whether column and row headers are shown
        /// </summary>
        public bool ShowRowColHeaders { get; private set; } = true; // default

        /// <summary>
        /// Gets whether rulers are shown in view type: pageLayout
        /// </summary>
        public bool ShowRuler { get; private set; } = true; // default

        /// <summary>
        /// Gets the sheet view type of the current worksheet
        /// </summary>
        public Worksheet.SheetViewType ViewType { get; private set; } = Worksheet.SheetViewType.normal; // default
        /// <summary>
        /// Gets the zoom factor of the current view type
        /// </summary>
        public int CurrentZoomScale { get; private set; } = 100; // default
        /// <summary>
        /// Gets all preserved zoom factors of the worksheet
        /// </summary>
        public Dictionary<Worksheet.SheetViewType, int> ZoomFactors { get; private set; } = new Dictionary<Worksheet.SheetViewType, int>();

        #endregion

        #region constructors

        /// <summary>
        /// Constructor with parameters
        /// </summary>
        /// <param name="sharedStrings">SharedStringsReader object</param>
        /// <param name="styleReaderContainer">Resolved styles, used to determine dates or times</param>
        /// <param name="options">Import options to override the automatic approach of the reader. <see cref="ImportOptions"/> for information about import options.</param>
        public WorksheetReader(SharedStringsReader sharedStrings, StyleReaderContainer styleReaderContainer, ImportOptions options = null)
        {
            importOptions = options;
            Data = new Dictionary<string, Cell>();
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
                Style style = styleReaderContainer.GetStyle(i, out isDate, out isTime);
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
                        string rowAttribute = ReaderUtils.GetAttribute(row, "r");
                        if (rowAttribute != null)
                        {
                            string hiddenAttribute = ReaderUtils.GetAttribute(row, "hidden");
                            RowDefinition.AddRowDefinition(Rows, rowAttribute, null, hiddenAttribute);
                            string heightAttribute = ReaderUtils.GetAttribute(row, "ht");
                            RowDefinition.AddRowDefinition(Rows, rowAttribute, heightAttribute, null);
                        }
                        if (row.HasChildNodes)
                        {
                            foreach (XmlNode rowChild in row.ChildNodes)
                            {
                                ReadCell(rowChild);
                            }
                        }
                    }
                    GetSheetView(xr);
                    GetMergedCells(xr);
                    GetSheetFormats(xr);
                    GetAutoFilters(xr);
                    GetColumns(xr);
                    GetSheetProtection(xr);
                }
            }
            catch (Exception ex)
            {
                throw new IOException("The XML entry could not be read from the input stream. Please see the inner exception:", ex);
            }
        }

        /// <summary>
        /// Gets the selected cells of the current worksheet
        /// </summary>
        /// <param name="xmlDocument">XML document of the current worksheet</param>
        private void GetSheetView(XmlDocument xmlDocument)
        {
            XmlNodeList sheetViewsNodes = xmlDocument.GetElementsByTagName("sheetViews");
            if (sheetViewsNodes != null && sheetViewsNodes.Count > 0)
            {
                XmlNodeList sheetViewNodes = sheetViewsNodes[0].ChildNodes;
                string attribute;
                // Go through all possible views
                foreach (XmlNode sheetView in sheetViewNodes)
                {
                    attribute = ReaderUtils.GetAttribute(sheetView, "view");
                    if (attribute != null)
                    {
                        Worksheet.SheetViewType viewType;
                        if (Enum.TryParse<Worksheet.SheetViewType>(attribute, out viewType))
                        {
                            ViewType = viewType;
                        }
                    }
                    attribute = ReaderUtils.GetAttribute(sheetView, "zoomScale");
                    if (attribute != null)
                    {
                        CurrentZoomScale = ReaderUtils.ParseInt(attribute);
                    }
                    attribute = ReaderUtils.GetAttribute(sheetView, "zoomScaleNormal");
                    if (attribute != null)
                    {
                        int scale = ReaderUtils.ParseInt(attribute);
                        ZoomFactors.Add(Worksheet.SheetViewType.normal, scale);
                    }
                    attribute = ReaderUtils.GetAttribute(sheetView, "zoomScalePageLayoutView");
                    if (attribute != null)
                    {
                        int scale = ReaderUtils.ParseInt(attribute);
                        ZoomFactors.Add(Worksheet.SheetViewType.pageLayout, scale);
                    }
                    attribute = ReaderUtils.GetAttribute(sheetView, "zoomScaleSheetLayoutView");
                    if (attribute != null)
                    {
                        int scale = ReaderUtils.ParseInt(attribute);
                        ZoomFactors.Add(Worksheet.SheetViewType.pageBreakPreview, scale);
                    }
                    attribute = ReaderUtils.GetAttribute(sheetView, "showGridLines");
                    if (attribute != null)
                    {
                        ShowGridLines = ReaderUtils.ParseBinaryBool(attribute) == 1;
                    }
                    attribute = ReaderUtils.GetAttribute(sheetView, "showRowColHeaders");
                    if (attribute != null)
                    {
                        ShowRowColHeaders = ReaderUtils.ParseBinaryBool(attribute) == 1;
                    }
                    attribute = ReaderUtils.GetAttribute(sheetView, "showRuler");
                    if (attribute != null)
                    {
                        ShowRuler = ReaderUtils.ParseBinaryBool(attribute) == 1;
                    }
                    if (sheetView.LocalName.Equals("sheetView", StringComparison.InvariantCultureIgnoreCase))
                    {
                        XmlNodeList selectionNodes = sheetView.ChildNodes;
                        if (selectionNodes != null && selectionNodes.Count > 0)
                        {
                            foreach(XmlNode selectionNode in selectionNodes)
                            {
                                attribute = ReaderUtils.GetAttribute(selectionNode, "sqref");
                                if (attribute != null)
                                {
                                    if (attribute.Contains(" "))
                                    {
                                        // Multiple ranges
                                        string[] ranges = attribute.Split(' ');
                                        foreach (string range in ranges)
                                        {
                                            CollectSelectedCells(range);
                                        }
                                    }
                                    else
                                    {
                                        CollectSelectedCells(attribute);
                                    }
                                    
                                }
                            }
                        }
                        XmlNode paneNode = ReaderUtils.GetChildNode(sheetView, "pane");
                        if (paneNode != null)
                        {
                            attribute = ReaderUtils.GetAttribute(paneNode, "state");
                            bool useNumbers = false;
                            this.PaneSplitValue = new PaneDefinition();
                            if (attribute != null)
                            {
                                this.PaneSplitValue.SetFrozenState(attribute);
                                useNumbers = this.PaneSplitValue.FrozenState;
                            }
                            attribute = ReaderUtils.GetAttribute(paneNode, "ySplit");
                            if (attribute != null)
                            {
                                this.PaneSplitValue.YSplitDefined = true;
                                if (useNumbers)
                                {
                                    this.PaneSplitValue.PaneSplitRowIndex = ReaderUtils.ParseInt(attribute);
                                }
                                else
                                {
                                    this.PaneSplitValue.PaneSplitHeight = Utils.GetPaneSplitHeight(ReaderUtils.ParseFloat(attribute));
                                }
                            }
                            attribute = ReaderUtils.GetAttribute(paneNode, "xSplit");
                            if (attribute != null)
                            {
                                this.PaneSplitValue.XSplitDefined = true;
                                if (useNumbers)
                                {
                                    this.PaneSplitValue.PaneSplitColumnIndex = ReaderUtils.ParseInt(attribute);
                                }
                                else
                                {
                                    this.PaneSplitValue.PaneSplitWidth = Utils.GetPaneSplitWidth(ReaderUtils.ParseFloat(attribute));
                                }
                            }
                            attribute = ReaderUtils.GetAttribute(paneNode, "topLeftCell");
                            if (attribute != null)
                            {
                                this.PaneSplitValue.TopLeftCell = new Address(attribute);
                            }
                            attribute = ReaderUtils.GetAttribute(paneNode, "activePane");
                            if (attribute != null)
                            {
                                this.PaneSplitValue.SetActivePane(attribute);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Resolves the selected cells of a range or a single cell
        /// </summary>
        /// <param name="attribute">Raw range/cell as string</param>
        private void CollectSelectedCells(string attribute)
        {
            if (attribute.Contains(":"))
            {
                // One range
                this.SelectedCells.Add(new Range(attribute));
            }
            else
            {
                // One cell
                this.SelectedCells.Add(new Range(attribute + ":" + attribute));
            }
        }

        /// <summary>
        /// Gets the sheet protection values of the current worksheets
        /// </summary>
        /// <param name="xmlDocument">XML document of the current worksheet</param>
        private void GetSheetProtection(XmlDocument xmlDocument)
        {
            XmlNodeList sheetProtectionNodes = xmlDocument.GetElementsByTagName("sheetProtection");
            if (sheetProtectionNodes != null && sheetProtectionNodes.Count > 0)
            {
                XmlNode sheetProtectionNode = sheetProtectionNodes[0];
                ManageSheetProtection(sheetProtectionNode, Worksheet.SheetProtectionValue.autoFilter);
                ManageSheetProtection(sheetProtectionNode, Worksheet.SheetProtectionValue.deleteColumns);
                ManageSheetProtection(sheetProtectionNode, Worksheet.SheetProtectionValue.deleteRows);
                ManageSheetProtection(sheetProtectionNode, Worksheet.SheetProtectionValue.formatCells);
                ManageSheetProtection(sheetProtectionNode, Worksheet.SheetProtectionValue.formatColumns);
                ManageSheetProtection(sheetProtectionNode, Worksheet.SheetProtectionValue.formatRows);
                ManageSheetProtection(sheetProtectionNode, Worksheet.SheetProtectionValue.insertColumns);
                ManageSheetProtection(sheetProtectionNode, Worksheet.SheetProtectionValue.insertHyperlinks);
                ManageSheetProtection(sheetProtectionNode, Worksheet.SheetProtectionValue.insertRows);
                ManageSheetProtection(sheetProtectionNode, Worksheet.SheetProtectionValue.objects);
                ManageSheetProtection(sheetProtectionNode, Worksheet.SheetProtectionValue.pivotTables);
                ManageSheetProtection(sheetProtectionNode, Worksheet.SheetProtectionValue.scenarios);
                ManageSheetProtection(sheetProtectionNode, Worksheet.SheetProtectionValue.selectLockedCells);
                ManageSheetProtection(sheetProtectionNode, Worksheet.SheetProtectionValue.selectUnlockedCells);
                ManageSheetProtection(sheetProtectionNode, Worksheet.SheetProtectionValue.sort);
                string legacyPasswordHash = ReaderUtils.GetAttribute(sheetProtectionNode, "password");
                if (legacyPasswordHash != null)
                {
                    this.WorksheetProtectionHash = legacyPasswordHash;
                }
            }
        }

        /// <summary>
        /// Manages particular sheet protection values if defined
        /// </summary>
        /// <param name="node">Sheet protection node</param>
        /// <param name="sheetProtectionValue">Value to check and maintain (if defined)</param>
        private void ManageSheetProtection(XmlNode node, Worksheet.SheetProtectionValue sheetProtectionValue)
        {
            string attributeName = Enum.GetName(typeof(Worksheet.SheetProtectionValue), sheetProtectionValue);
            string attribute = ReaderUtils.GetAttribute(node, attributeName);
            if (attribute != null)
            {
                int value = ReaderUtils.ParseBinaryBool(attribute);
                WorksheetProtection.Add(sheetProtectionValue, value);
            }
        }

        /// <summary>
        /// Gets the merged cells of the current worksheet
        /// </summary>
        /// <param name="xmlDocument">XML document of the current worksheet</param>
        private void GetMergedCells(XmlDocument xmlDocument)
        {
            XmlNodeList mergedCellsNodes = xmlDocument.GetElementsByTagName("mergeCells");
            if (mergedCellsNodes != null && mergedCellsNodes.Count > 0)
            {
                XmlNodeList mergedCellNodes = mergedCellsNodes[0].ChildNodes;
                if (mergedCellNodes != null && mergedCellNodes.Count > 0)
                {
                    foreach(XmlNode mergedCells in mergedCellNodes)
                    {
                        string attribute = ReaderUtils.GetAttribute(mergedCells, "ref");
                        if (attribute != null)
                        {
                            this.MergedCells.Add(new Range(attribute));
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Gets the sheet format information of the current worksheet
        /// </summary>
        /// <param name="xmlDocument">XML document of the current worksheet</param>
        private void GetSheetFormats(XmlDocument xmlDocument)
        {
            XmlNodeList formatNodes = xmlDocument.GetElementsByTagName("sheetFormatPr");
            if (formatNodes != null && formatNodes.Count > 0)
            {
                string attribute = ReaderUtils.GetAttribute(formatNodes[0], "defaultColWidth");
                if (attribute != null)
                {
                    this.DefaultColumnWidth = ReaderUtils.ParseFloat(attribute);
                }
                attribute = ReaderUtils.GetAttribute(formatNodes[0], "defaultRowHeight");
                if (attribute != null)
                {
                    this.DefaultRowHeight = ReaderUtils.ParseFloat(attribute);
                }
            }
        }

        /// <summary>
        /// Gets the auto filters of the current worksheet
        /// </summary>
        /// <param name="xmlDocument">XML document of the current worksheet</param>
        private void GetAutoFilters(XmlDocument xmlDocument)
        {
            XmlNodeList autoFilterNodes = xmlDocument.GetElementsByTagName("autoFilter");
            if (autoFilterNodes != null && autoFilterNodes.Count > 0)
            {
                string autoFilterRef = ReaderUtils.GetAttribute(autoFilterNodes[0], "ref");
                if (autoFilterRef != null)
                {
                    if (autoFilterRef.Contains(":"))
                    {
                        this.AutoFilterRange = new Range(autoFilterRef);
                    }
                    else
                    {
                        Address address = new Address(autoFilterRef);
                        this.AutoFilterRange = new Range(address, address);
                    }
                }
            }
        }

        /// <summary>
        /// Gets the columns of the current worksheet
        /// </summary>
        /// <param name="xmlDocument">XML document of the current worksheet</param>
        private void GetColumns(XmlDocument xmlDocument)
        {
            XmlNodeList columnNodes = xmlDocument.GetElementsByTagName("col");
            foreach (XmlNode columnNode in columnNodes)
            {
                int? min = null;
                int? max = null;
                List<int> indices = new List<int>();
                string attribute = ReaderUtils.GetAttribute(columnNode, "min");
                if (attribute != null)
                {
                    min = ReaderUtils.ParseInt(attribute);
                    max = min;
                    indices.Add(min.Value);
                }
                attribute = ReaderUtils.GetAttribute(columnNode, "max");
                if (attribute != null)
                {
                    max = ReaderUtils.ParseInt(attribute);
                }
                if (min != null && max.Value != min.Value)
                {
                    for (int i = min.Value; i <= max.Value; i++)
                    {
                        indices.Add(i);
                    }
                }
                attribute = ReaderUtils.GetAttribute(columnNode, "width");
                float width = Worksheet.DEFAULT_COLUMN_WIDTH;
                if (attribute != null)
                {
                    width = ReaderUtils.ParseFloat(attribute);
                }
                attribute = ReaderUtils.GetAttribute(columnNode, "hidden");
                bool hidden = false;
                if (attribute != null)
                {
                    int value = ReaderUtils.ParseBinaryBool(attribute);
                    if (value == 1)
                    {
                        hidden = true;
                    }
                }
                attribute = ReaderUtils.GetAttribute(columnNode, "style");
                Style defaultStyle = null;
                if (attribute != null)
				{
                    if (resolvedStyles.ContainsKey(attribute))
					{
                        defaultStyle = resolvedStyles[attribute];
					}
				}
                foreach (int index in indices)
                {
                    Column column = new Column(index - 1); // Transform to zero-based
                    column.Width = width;
                    column.IsHidden = hidden;
                    if (defaultStyle != null)
					{
                        column.SetDefaultColumnStyle(defaultStyle);
					}
                    this.Columns.Add(column);
                }
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
                address = ReaderUtils.GetAttribute(rowChild, "r"); // Mandatory
                type = ReaderUtils.GetAttribute(rowChild, "t"); // can be null if not existing
                styleNumber = ReaderUtils.GetAttribute(rowChild, "s"); // can be null
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
                case double _:
                case byte _:
                case sbyte _:
                case int _:
                    return Cell.CellType.NUMBER;
                case DateTime _:
                    return Cell.CellType.DATE;
                case TimeSpan _:
                    return Cell.CellType.TIME;
                case bool _:
                    return Cell.CellType.BOOL;
                default:
                    return Cell.CellType.STRING;
            }
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
            if (importOptions.EnforceEmptyValuesAsString && data == null)
            {
                return "";
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
            else if (importOptions.GlobalEnforcingType == ImportOptions.GlobalType.AllNumbersToDecimal)
            {
                object tempDecimal = ConvertToDecimal(data);
                if (tempDecimal != null)
                {
                    return tempDecimal;
                }
            }
            else if (importOptions.GlobalEnforcingType == ImportOptions.GlobalType.AllNumbersToInt)
            {
                object tempInt = ConvertToInt(data);
                if (tempInt != null)
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
                case ImportOptions.ColumnType.Decimal:
                    return ConvertToDecimal(data);
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
                    
                    string tempString = (string)data;
                    bool? tempBool = TryParseBool(tempString);
                    if (tempBool != null)
                    {
                        return tempBool.Value;
                    }
                    break;
            }
            return data;
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
        /// Tries to convert a value to a double
        /// </summary>
        /// <param name="data">Raw data</param>
        /// <returns>Double value or original value if not possible to convert</returns>
        private object ConvertToDouble(object data)
        {
            object value = ConvertToDecimal(data);
            if (value is decimal)
            {
                return Decimal.ToDouble((decimal)value);
            }
            else if (value is float)
            {
                return Convert.ToDouble((float)value);
            }
            return value;
        }

        /// <summary>
        /// Tries to convert a value to a decimal
        /// </summary>
        /// <param name="data">Raw data</param>
        /// <returns>Decimal value or original value if not possible to convert</returns>
        private object ConvertToDecimal(object data)
        {
            IConvertible converter;
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
                    converter = data as IConvertible;
                    double tempDouble = converter.ToDouble(Utils.INVARIANT_CULTURE);
                    if (tempDouble > (double)decimal.MaxValue || tempDouble < (double)decimal.MinValue)
                    {
                        return data;
                    }
                    else
                    {
                        return converter.ToDecimal(Utils.INVARIANT_CULTURE);
                    }
                case bool _:
                    if ((bool)data)
                    {
                        return decimal.One;
                    }
                    else
                    {
                        return decimal.Zero;
                    }
                case DateTime _:
                    return new decimal(Utils.GetOADateTime((DateTime)data));
                case TimeSpan _:
                    return  new decimal(Utils.GetOATime((TimeSpan)data));
                case string _:
                    decimal dValue;
                    string tempString = (string)data;
                    if (decimal.TryParse(tempString, NumberStyles.Float, CultureInfo.InvariantCulture, out dValue))
                    {
                        return dValue;
                    }
                    DateTime? tempDate = TryParseDate(tempString);
                    if (tempDate != null)
                    {
                        return new decimal(Utils.GetOADateTime(tempDate.Value));
                    }
                    TimeSpan? tempTime = TryParseTime(tempString);
                    if (tempTime != null)
                    {
                        return new decimal(Utils.GetOATime(tempTime.Value));
                    }
                    break;
            }
            return data;
        }

        /// <summary>
        /// Tries to convert a value to an integer
        /// </summary>
        /// <param name="data">Raw data</param>
        /// <returns>Integer value or null if not possible to convert</returns>
        private object ConvertToInt(object data)
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
                    return ConvertDoubleToInt(tempDouble);
                case TimeSpan _:
                    tempDouble = Utils.GetOATime((TimeSpan)data);
                    return ConvertDoubleToInt(tempDouble);
                case float _:
                case double _:
                    object tempInt = TryConvertDoubleToInt(data);
                    if (tempInt != null)
                    {
                        return tempInt;
                    }
                    break;
                case bool _:
                    return (bool)data ? 1 : 0;
                case string _:
                    int tempInt2;
                    if (ReaderUtils.TryParseInt((string)data, out tempInt2))
                    {
                        return tempInt2;
                    }
                    break;
            }
            return null;
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
                    DateTime? date2 = TryParseDate((string)data);
                    if(date2 != null)
                    {
                        return date2.Value;
                    }
                    return ConvertDateFromDouble(data);
            }
            return data;
        }

        /// <summary>
        /// Tris to parse a DateTime instance from a string
        /// </summary>
        /// <param name="raw">String to parse</param>
        /// <returns>DateTime instance or null if not possible to parse</returns>
        private DateTime? TryParseDate(string raw)
        {
            DateTime dateTime;
            bool isDateTime;
            if (importOptions == null || string.IsNullOrEmpty(importOptions.DateTimeFormat) || importOptions.TemporalCultureInfo == null)
            {
                isDateTime = DateTime.TryParse(raw, ImportOptions.DEFAULT_CULTURE_INFO, DateTimeStyles.None, out dateTime);
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
                    TimeSpan? time = TryParseTime((string)data);
                    if(time != null)
                    {
                        return time;
                    }
                    return ConvertTimeFromDouble(data);
            }
            return data;
        }

        /// <summary>
        /// Tris to parse a TimeSpan instance from a string
        /// </summary>
        /// <param name="raw">String to parse</param>
        /// <returns>TimeSpan instance or null if not possible to parse</returns>
        private TimeSpan? TryParseTime(string raw)
        {
            TimeSpan timeSpan;
            bool isTimeSpan;
            if (importOptions == null || string.IsNullOrEmpty(importOptions.TimeSpanFormat) || importOptions.TemporalCultureInfo == null)
            {
                isTimeSpan = TimeSpan.TryParse(raw, ImportOptions.DEFAULT_CULTURE_INFO,  out timeSpan);
            }
            else
            {
                isTimeSpan = TimeSpan.TryParseExact(raw, importOptions.TimeSpanFormat, importOptions.TemporalCultureInfo, out timeSpan);
            }
            if (isTimeSpan && timeSpan.Days >= 0 && timeSpan.Days < Utils.MAX_OADATE_VALUE)
            {
                return timeSpan;
            }
            return null;
        }

        /// <summary>
        /// Parses the date (DateTime) or time (TimeSpan) value of a raw cell. If the value is numeric, but out of range of a OAdate, a numeric value will be returned instead. 
        /// If invalid, the string representation will be returned.
        /// </summary>
        /// <param name="raw">Raw value as string</param>
        /// <param name="valueType">Type of the value to be converted: Valid values are DATE and TIME</param>
        /// <param name="resolvedType">Out parameter for the determined value type</param>
        /// <returns>Object of the type TimeSpan or null if not possible to parse</returns>
        private object GetDateTimeValue(string raw, Cell.CellType valueType, out Cell.CellType resolvedType)
        {
            double dValue;
            if (!double.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out dValue))
            {
                resolvedType = Cell.CellType.STRING;
                return raw;
            }
            if ((valueType == Cell.CellType.DATE && (dValue < Utils.MIN_OADATE_VALUE || dValue > Utils.MAX_OADATE_VALUE)) || (valueType == Cell.CellType.TIME && (dValue < 0.0 || dValue > Utils.MAX_OADATE_VALUE)))
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
        /// <returns>Converted value if possible to convert, otherwise null</returns>
        private object TryConvertDoubleToInt(object data)
        {
            IConvertible converter = data as IConvertible;
            double dValue = converter.ToDouble(ImportOptions.DEFAULT_CULTURE_INFO);
            if (dValue > int.MinValue && dValue < int.MaxValue)
            {
                return converter.ToInt32(ImportOptions.DEFAULT_CULTURE_INFO);
            }
            return null;
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
                        return (bool)tempObject ? 1 : 0;
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
            bool canBeUint = uint.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out uiValue);
            bool canBeInt = ReaderUtils.TryParseInt(raw, out iValue);
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
            bool canBeUlong = ulong.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out ulValue);
            bool canBeLong = long.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out lValue);
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
                if (importOptions?.GlobalEnforcingType == ImportOptions.GlobalType.AllSingleToDecimal)
                    return dcValue;

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
            if (ReaderUtils.TryParseInt(raw, out stringId))
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

        #region subClasses

        /// <summary>
        /// Class representing a pane in the applications window
        /// </summary>
        public class PaneDefinition
        {
            /// <summary>
            /// Gets or sets the pane split height of a worksheet split
            /// </summary>
            public float? PaneSplitHeight { get; set; }
            /// <summary>
            /// Gets or sets the pane split width of a worksheet split
            /// </summary>
            public float? PaneSplitWidth { get; set; }
            /// <summary>
            /// Gets or sets the pane split row index of a worksheet split
            /// </summary>
            public int? PaneSplitRowIndex { get; set; }
            /// <summary>
            /// Gets or sets the pane split column index of a worksheet split
            /// </summary>
            public int? PaneSplitColumnIndex { get; set; }
            /// <summary>
            /// Top Left cell address of the bottom right pane
            /// </summary>
            public Address TopLeftCell { get; set; }
            /// <summary>
            /// Active pane in the split window
            /// </summary>
            public Worksheet.WorksheetPane? ActivePane { get; private set; }
            /// <summary>
            /// Frozen state of the split window
            /// </summary>
            public bool FrozenState { get; private set; }

            /// <summary>
            /// Gets whether an Y split was defined
            /// </summary>
            public bool YSplitDefined { get; set; }

            /// <summary>
            /// Gets whether an X split was defined
            /// </summary>
            public bool XSplitDefined { get; set; }

            /// <summary>
            /// Default constructor, with no active pane and the top left cell at A1
            /// </summary>
            public PaneDefinition()
            {
                ActivePane = null;
                TopLeftCell = new Address(0, 0);
            }

            /// <summary>
            /// Parses and sets the active pane from a string value
            /// </summary>
            /// <param name="value">raw enum value as string</param>
            public void SetActivePane(string value)
            {
                this.ActivePane = (Worksheet.WorksheetPane)Enum.Parse(typeof(Worksheet.WorksheetPane), value);
            }

            /// <summary>
            /// Sets the frozen state of the split window if defined
            /// </summary>
            /// <param name="value">raw attribute value</param>
            public void SetFrozenState(string value)
            {
                if (value.ToLower() == "frozen" || value.ToLower() == "frozensplit")
                {
                    this.FrozenState = true;
                }
            }

        }

        /// <summary>
        /// Internal class to represent a row
        /// </summary>
        public class RowDefinition
        {
            /// <summary>
            /// Indicates whether the row is hidden
            /// </summary>
            public bool Hidden { get; set; }
            /// <summary>
            /// Non-standard row height
            /// </summary>
            public float? Height { get; set; } = null;

            /// <summary>
            /// Adds a row definition or changes it, when a non-standard row height and/or hidden state is defined
            /// </summary>
            /// <param name="rows">Row dictionary</param>
            /// <param name="rowNumber">Row number as string (directly resolved from the corresponding XML attribute)</param>
            /// <param name="heightProperty">Row height as string (directly resolved from the corresponding XML attribute)</param>
            /// <param name="hiddenProperty">Hidden definition as string (directly resolved from the corresponding XML attribute)</param>
            public static void AddRowDefinition(Dictionary<int, RowDefinition> rows, string rowNumber, string heightProperty, string hiddenProperty)
            {
                int row = ReaderUtils.ParseInt(rowNumber) - 1; // Transform to zero-based
                if (!rows.ContainsKey(row))
                {
                    rows.Add(row, new RowDefinition());
                }
                if (heightProperty != null)
                {
                    rows[row].Height = ReaderUtils.ParseFloat(heightProperty);
                }
                if (hiddenProperty != null)
                {
                    int value = ReaderUtils.ParseBinaryBool(hiddenProperty);
                    rows[row].Hidden = value == 1;
                    
                }
            }
        }
        #endregion
    }
}
