/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml;
using NanoXLSX.Exceptions;
using NanoXLSX.Interfaces.Plugin;
using NanoXLSX.Interfaces.Reader;
using NanoXLSX.Registry;
using NanoXLSX.Styles;
using NanoXLSX.Utils;
using static NanoXLSX.Internal.Enums.ReaderPassword;
using IOException = NanoXLSX.Exceptions.IOException;

namespace NanoXLSX.Internal.Readers
{
    /// <summary>
    /// Class representing a reader for worksheets of XLSX files
    /// </summary>
    public class WorksheetReader : IWorksheetReader
    {
        #region privateFields
        private MemoryStream stream;
        private ReaderOptions readerOptions;
        private List<string> dateStyles = null;
        private List<string> timeStyles = null;
        private Dictionary<string, Style> resolvedStyles = null;
        private IPasswordReader passwordReader = null;
        #endregion

        #region properties

        /// <summary>
        /// Workbook reference where read data is stored (should not be null)
        /// </summary>
        public Workbook Workbook { get; set; }
        /// <summary>
        /// Gets or sets the (r)ID of the current worksheet
        /// </summary>
        public int CurrentWorksheetID { get; set; }

        /// <summary>
        /// Gets or Sets the list of the shared strings. The index of the list corresponds to the index, defined in cell values
        /// </summary>
        public List<String> SharedStrings { get; set; }
        #endregion

        #region constructors
        /// <summary>
        /// Default constructor - Must be defined for instantiation of the plug-ins
        /// </summary>
        public WorksheetReader()
        {
        }
        #endregion

        #region functions
        /// <summary>
        /// Initialization method (interface implementation)
        /// </summary>
        /// <param name="stream">MemoryStream to be read</param>
        /// <param name="workbook">Workbook reference</param>
        /// <param name="readerOptions">Reader options</param>
        public void Init(MemoryStream stream, Workbook workbook, IOptions readerOptions)
        {
            this.stream = stream;
            this.Workbook = workbook;
            this.readerOptions = readerOptions as ReaderOptions;
            if (dateStyles == null || timeStyles == null || this.resolvedStyles == null)
            {
                StyleReaderContainer styleReaderContainer = workbook.AuxiliaryData.GetData<StyleReaderContainer>(PlugInUUID.STYLE_READER, PlugInUUID.STYLES_ENTITY);
                ProcessStyles(styleReaderContainer);
            }
            if (this.passwordReader == null)
            {
                this.passwordReader = PlugInLoader.GetPlugIn<IPasswordReader>(PlugInUUID.PASSWORD_READER, new LegacyPasswordReader());
                this.passwordReader.Init(PasswordType.WORKSHEET_PROTECTION, this.readerOptions);
            }
        }

        /// <summary>
        /// Method to execute the main logic of the plug-in (interface implementation)
        /// </summary>
        /// <exception cref="Exceptions.IOException">Throws an IOException in case of a error during reading</exception>
        public void Execute()
        {
            try
            {
                WorksheetDefinition worksheetDefinition = Workbook.AuxiliaryData.GetData<WorksheetDefinition>(PlugInUUID.WORKBOOK_READER, PlugInUUID.WORKSHEET_DEFINITION_ENTITY, CurrentWorksheetID);
                Worksheet worksheet = new Worksheet(worksheetDefinition.WorksheetName, CurrentWorksheetID, Workbook);
                worksheet.Hidden = worksheetDefinition.Hidden;
                using (stream) // Close after processing
                {
                    XmlDocument document = new XmlDocument();
                    document.XmlResolver = null;
                    document.Load(stream);
                    GetRows(document, worksheet);
                    GetSheetView(document, worksheet);
                    GetMergedCells(document, worksheet);
                    GetSheetFormats(document, worksheet);
                    GetAutoFilters(document, worksheet);
                    GetColumns(document, worksheet);
                    GetSheetProtection(document, worksheet);
                    SetWorkbookRelation(worksheet);
                    RederPlugInHandler.HandleInlineQueuePlugins(ref stream, Workbook, PlugInUUID.WORKSHEET_INLINE_READER, CurrentWorksheetID);
                }


            }
            catch (NotSupportedContentException ex)
            {
                throw ex; // rethrow
            }
            catch (Exception ex)
            {
                throw new IOException("The XML entry could not be read from the input stream. Please see the inner exception:", ex);
            }
        }

        /// <summary>
        /// Sets all relation details of the worksheet to its parent workbook
        /// </summary>
        /// <param name="worksheet">Worksheet to process</param>
        private void SetWorkbookRelation(Worksheet worksheet)
        {
            Workbook.AddWorksheet(worksheet);
            int selectedWorksheetId = Workbook.AuxiliaryData.GetData<int>(PlugInUUID.WORKBOOK_READER, PlugInUUID.SELECTED_WORKSHEET_ENTITY);
            if (selectedWorksheetId + 1 == CurrentWorksheetID) // selectedWorksheetId is 0-based
            {
                Workbook.SetSelectedWorksheet(worksheet);
            }
        }

        /// <summary>
        /// Determine which of the resolved styles are either to define a time or a date. Stores also the styles into a dictionary 
        /// </summary>
        /// <param name="styleReaderContainer">Resolved styles from the style reader</param>
        private void ProcessStyles(StyleReaderContainer styleReaderContainer)
        {
            this.dateStyles = new List<string>();
            this.timeStyles = new List<string>();
            this.resolvedStyles = new Dictionary<string, Style>();
            for (int i = 0; i < styleReaderContainer.StyleCount; i++)
            {
                bool isDate;
                bool isTime;
                string index = ParserUtils.ToString(i);
                Style style = styleReaderContainer.GetStyle(i, out isDate, out isTime);
                if (isDate)
                {
                    this.dateStyles.Add(index);
                }
                if (isTime)
                {
                    this.timeStyles.Add(index);
                }
                this.resolvedStyles.Add(index, style);
            }
        }

        /// <summary>
        /// Gets the row definitions of the current worksheet
        /// </summary>
        /// <param name="document">XML document of the current worksheet</param>
        /// <param name="worksheet">Currently processed worksheet</param>
        private void GetRows(XmlDocument document, Worksheet worksheet)
        {
            XmlNodeList rows = document.GetElementsByTagName("row");
            foreach (XmlNode row in rows)
            {
                string rowAttribute = ReaderUtils.GetAttribute(row, "r");
                if (rowAttribute != null)
                {
                    int rowNumber = ParserUtils.ParseInt(rowAttribute) - 1; // Transform to zero-based
                    string hiddenAttribute = ReaderUtils.GetAttribute(row, "hidden");
                    if (hiddenAttribute != null)
                    {
                        int value = ParserUtils.ParseBinaryBool(hiddenAttribute);
                        if (value == 1)
                        {
                            worksheet.AddHiddenRow(rowNumber);
                        }
                    }
                    string heightAttribute = ReaderUtils.GetAttribute(row, "ht");
                    if (heightAttribute != null)
                    {
                        worksheet.RowHeights.Add(rowNumber, ParserUtils.ParseFloat(heightAttribute));
                    }
                }
                if (row.HasChildNodes)
                {
                    foreach (XmlNode rowChild in row.ChildNodes)
                    {
                        ReadCell(rowChild, worksheet);
                    }
                }
            }
        }

        /// <summary>
        /// Gets the selected cells of the current worksheet
        /// </summary>
        /// <param name="xmlDocument">XML document of the current worksheet</param>
        /// <param name="worksheet">Currently processed worksheet</param>
        private void GetSheetView(XmlDocument xmlDocument, Worksheet worksheet)
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
                            worksheet.ViewType = viewType;
                        }
                    }
                    attribute = ReaderUtils.GetAttribute(sheetView, "zoomScale");
                    if (attribute != null)
                    {
                        worksheet.ZoomFactor = ParserUtils.ParseInt(attribute);
                    }
                    attribute = ReaderUtils.GetAttribute(sheetView, "zoomScaleNormal");
                    if (attribute != null)
                    {
                        int scale = ParserUtils.ParseInt(attribute);
                        worksheet.ZoomFactors[Worksheet.SheetViewType.normal] = scale;
                    }
                    attribute = ReaderUtils.GetAttribute(sheetView, "zoomScalePageLayoutView");
                    if (attribute != null)
                    {
                        int scale = ParserUtils.ParseInt(attribute);
                        worksheet.ZoomFactors[Worksheet.SheetViewType.pageLayout] = scale;
                    }
                    attribute = ReaderUtils.GetAttribute(sheetView, "zoomScaleSheetLayoutView");
                    if (attribute != null)
                    {
                        int scale = ParserUtils.ParseInt(attribute);
                        worksheet.ZoomFactors[Worksheet.SheetViewType.pageBreakPreview] = scale;
                    }
                    attribute = ReaderUtils.GetAttribute(sheetView, "showGridLines");
                    if (attribute != null)
                    {
                        worksheet.ShowGridLines = ParserUtils.ParseBinaryBool(attribute) == 1;
                    }
                    attribute = ReaderUtils.GetAttribute(sheetView, "showRowColHeaders");
                    if (attribute != null)
                    {
                        worksheet.ShowRowColumnHeaders = ParserUtils.ParseBinaryBool(attribute) == 1;
                    }
                    attribute = ReaderUtils.GetAttribute(sheetView, "showRuler");
                    if (attribute != null)
                    {
                        worksheet.ShowRuler = ParserUtils.ParseBinaryBool(attribute) == 1;
                    }
                    if (sheetView.LocalName.Equals("sheetView", StringComparison.InvariantCultureIgnoreCase))
                    {
                        XmlNodeList selectionNodes = sheetView.ChildNodes;
                        if (selectionNodes != null && selectionNodes.Count > 0)
                        {
                            foreach (XmlNode selectionNode in selectionNodes)
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
                                            CollectSelectedCells(range, worksheet);
                                        }
                                    }
                                    else
                                    {
                                        CollectSelectedCells(attribute, worksheet);
                                    }

                                }
                            }
                        }
                        XmlNode paneNode = ReaderUtils.GetChildNode(sheetView, "pane");
                        if (paneNode != null)
                        {
                            SetPaneSplit(paneNode, worksheet);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Resolves the selected cells of a range or a single cell
        /// </summary>
        /// <param name="attribute">Raw range/cell as string</param>
        /// <param name="worksheet">Currently processed worksheet</param>
        private void CollectSelectedCells(string attribute, Worksheet worksheet)
        {
            if (attribute.Contains(":"))
            {
                // One range
                worksheet.AddSelectedCells(new Range(attribute));
            }
            else
            {
                // One cell
                worksheet.AddSelectedCells(new Range(attribute + ":" + attribute));
            }
        }

        /// <summary>
        /// Sets the pane split values of the current worksheet
        /// </summary>
        /// <param name="paneNode">XML node</param>
        /// <param name="worksheet">Currently processed worksheet</param>
        private void SetPaneSplit(XmlNode paneNode, Worksheet worksheet)
        {
            string attribute = ReaderUtils.GetAttribute(paneNode, "state");
            bool useNumbers = false;
            bool frozenState = false;
            bool ySplitDefined = false;
            bool xSplitDefined = false;
            int? paneSplitRowIndex = null;
            int? paneSplitColumnIndex = null;
            float? paneSplitHeight = null;
            float? paneSplitWidth = null;
            Address topLeftCell = new Address(0, 0); // default value
            Worksheet.WorksheetPane? activePane = null;
            if (attribute != null)
            {
                if (attribute.ToLower() == "frozen" || attribute.ToLower() == "frozensplit")
                {
                    frozenState = true;
                }
                useNumbers = frozenState;
            }
            attribute = ReaderUtils.GetAttribute(paneNode, "ySplit");
            if (attribute != null)
            {
                ySplitDefined = true;
                if (useNumbers)
                {
                    paneSplitRowIndex = ParserUtils.ParseInt(attribute);
                }
                else
                {
                    paneSplitHeight = DataUtils.GetPaneSplitHeight(ParserUtils.ParseFloat(attribute));
                }
            }
            attribute = ReaderUtils.GetAttribute(paneNode, "xSplit");
            if (attribute != null)
            {
                xSplitDefined = true;
                if (useNumbers)
                {
                    paneSplitColumnIndex = ParserUtils.ParseInt(attribute);
                }
                else
                {
                    paneSplitWidth = DataUtils.GetPaneSplitWidth(ParserUtils.ParseFloat(attribute));
                }
            }
            attribute = ReaderUtils.GetAttribute(paneNode, "topLeftCell");
            if (attribute != null)
            {
                topLeftCell = new Address(attribute);
            }
            attribute = ReaderUtils.GetAttribute(paneNode, "activePane");
            if (attribute != null)
            {
                activePane = (Worksheet.WorksheetPane)Enum.Parse(typeof(Worksheet.WorksheetPane), attribute);
            }
            // assign to worksheet
            if (frozenState)
            {
                if (ySplitDefined && !xSplitDefined)
                {
                    worksheet.SetHorizontalSplit(paneSplitRowIndex.Value, frozenState, topLeftCell, activePane);
                }
                if (!ySplitDefined && xSplitDefined)
                {
                    worksheet.SetVerticalSplit(paneSplitColumnIndex.Value, frozenState, topLeftCell, activePane);
                }
                else if (ySplitDefined && xSplitDefined)
                {
                    worksheet.SetSplit(paneSplitColumnIndex.Value, paneSplitRowIndex.Value, frozenState, topLeftCell, activePane);
                }
            }
            else
            {
                if (ySplitDefined && !xSplitDefined)
                {
                    worksheet.SetHorizontalSplit(paneSplitHeight.Value, topLeftCell, activePane);
                }
                if (!ySplitDefined && xSplitDefined)
                {
                    worksheet.SetVerticalSplit(paneSplitWidth.Value, topLeftCell, activePane);
                }
                else if (ySplitDefined && xSplitDefined)
                {
                    worksheet.SetSplit(paneSplitWidth, paneSplitHeight, topLeftCell, activePane);
                }
            }
        }

        /// <summary>
        /// Gets the sheet protection values of the current worksheets
        /// </summary>
        /// <param name="xmlDocument">XML document of the current worksheet</param>
        /// <param name="worksheet">Currently processed worksheet</param>
        private void GetSheetProtection(XmlDocument xmlDocument, Worksheet worksheet)
        {
            XmlNodeList sheetProtectionNodes = xmlDocument.GetElementsByTagName("sheetProtection");
            if (sheetProtectionNodes != null && sheetProtectionNodes.Count > 0)
            {
                int hasProtection = 0;
                XmlNode sheetProtectionNode = sheetProtectionNodes[0];
                hasProtection += ManageSheetProtection(sheetProtectionNode, Worksheet.SheetProtectionValue.autoFilter, worksheet);
                hasProtection += ManageSheetProtection(sheetProtectionNode, Worksheet.SheetProtectionValue.deleteColumns, worksheet);
                hasProtection += ManageSheetProtection(sheetProtectionNode, Worksheet.SheetProtectionValue.deleteRows, worksheet);
                hasProtection += ManageSheetProtection(sheetProtectionNode, Worksheet.SheetProtectionValue.formatCells, worksheet);
                hasProtection += ManageSheetProtection(sheetProtectionNode, Worksheet.SheetProtectionValue.formatColumns, worksheet);
                hasProtection += ManageSheetProtection(sheetProtectionNode, Worksheet.SheetProtectionValue.formatRows, worksheet);
                hasProtection += ManageSheetProtection(sheetProtectionNode, Worksheet.SheetProtectionValue.insertColumns, worksheet);
                hasProtection += ManageSheetProtection(sheetProtectionNode, Worksheet.SheetProtectionValue.insertHyperlinks, worksheet);
                hasProtection += ManageSheetProtection(sheetProtectionNode, Worksheet.SheetProtectionValue.insertRows, worksheet);
                hasProtection += ManageSheetProtection(sheetProtectionNode, Worksheet.SheetProtectionValue.objects, worksheet);
                hasProtection += ManageSheetProtection(sheetProtectionNode, Worksheet.SheetProtectionValue.pivotTables, worksheet);
                hasProtection += ManageSheetProtection(sheetProtectionNode, Worksheet.SheetProtectionValue.scenarios, worksheet);
                hasProtection += ManageSheetProtection(sheetProtectionNode, Worksheet.SheetProtectionValue.selectLockedCells, worksheet);
                hasProtection += ManageSheetProtection(sheetProtectionNode, Worksheet.SheetProtectionValue.selectUnlockedCells, worksheet);
                hasProtection += ManageSheetProtection(sheetProtectionNode, Worksheet.SheetProtectionValue.sort, worksheet);
                if (hasProtection > 0)
                {
                    worksheet.UseSheetProtection = true;
                }
                this.passwordReader.ReadXmlAttributes(sheetProtectionNode);
                if (this.passwordReader.PasswordIsSet())
                {
                    if (this.passwordReader is LegacyPasswordReader && (this.passwordReader as LegacyPasswordReader).ContemporaryAlgorithmDetected && !readerOptions.IgnoreNotSupportedPasswordAlgorithms)
                    {
                        throw new NotSupportedContentException("A not supported, contemporary password algorithm for the worksheet protection was detected. Check possible packages to add support to NanoXLSX, or ignore this error by a reader option");
                    }
                    worksheet.SheetProtectionPassword.CopyFrom(this.passwordReader);
                }
            }
        }

        /// <summary>
        /// Manages particular sheet protection values if defined
        /// </summary>
        /// <param name="node">Sheet protection node</param>
        /// <param name="sheetProtectionValue">Value to check and maintain (if defined)</param>
        /// <param name="worksheet">Currently processed worksheet</param>
        private int ManageSheetProtection(XmlNode node, Worksheet.SheetProtectionValue sheetProtectionValue, Worksheet worksheet)
        {
            int hasProtection = 0;
            string attributeName = Enum.GetName(typeof(Worksheet.SheetProtectionValue), sheetProtectionValue);
            string attribute = ReaderUtils.GetAttribute(node, attributeName);
            if (attribute != null)
            {
                hasProtection = 1;
                // hasProtection = ParserUtils.ParseBinaryBool(attribute) != 0;
                worksheet.SheetProtectionValues.Add(sheetProtectionValue);
                //WorksheetProtection.Add(sheetProtectionValue, value);
            }
            return hasProtection;
        }

        /// <summary>
        /// Gets the merged cells of the current worksheet
        /// </summary>
        /// <param name="xmlDocument">XML document of the current worksheet</param>
        /// <param name="worksheet">Currently processed worksheet</param>
        private void GetMergedCells(XmlDocument xmlDocument, Worksheet worksheet)
        {
            XmlNodeList mergedCellsNodes = xmlDocument.GetElementsByTagName("mergeCells");
            if (mergedCellsNodes != null && mergedCellsNodes.Count > 0)
            {
                XmlNodeList mergedCellNodes = mergedCellsNodes[0].ChildNodes;
                if (mergedCellNodes != null && mergedCellNodes.Count > 0)
                {
                    foreach (XmlNode mergedCells in mergedCellNodes)
                    {
                        string attribute = ReaderUtils.GetAttribute(mergedCells, "ref");
                        if (attribute != null)
                        {
                            worksheet.MergeCells(new Range(attribute));
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Gets the sheet format information of the current worksheet
        /// </summary>
        /// <param name="xmlDocument">XML document of the current worksheet</param>
        /// <param name="worksheet">Currently processed worksheet</param></param>
        private void GetSheetFormats(XmlDocument xmlDocument, Worksheet worksheet)
        {
            XmlNodeList formatNodes = xmlDocument.GetElementsByTagName("sheetFormatPr");
            if (formatNodes != null && formatNodes.Count > 0)
            {
                string attribute = ReaderUtils.GetAttribute(formatNodes[0], "defaultColWidth");
                if (attribute != null)
                {
                    worksheet.DefaultColumnWidth = ParserUtils.ParseFloat(attribute);
                }
                attribute = ReaderUtils.GetAttribute(formatNodes[0], "defaultRowHeight");
                if (attribute != null)
                {
                    worksheet.DefaultRowHeight = ParserUtils.ParseFloat(attribute);
                }
            }
        }

        /// <summary>
        /// Gets the auto filters of the current worksheet
        /// </summary>
        /// <param name="xmlDocument">XML document of the current worksheet</param>
        /// <param name="worksheet">Currently processed worksheet</param>
        private void GetAutoFilters(XmlDocument xmlDocument, Worksheet worksheet)
        {
            XmlNodeList autoFilterNodes = xmlDocument.GetElementsByTagName("autoFilter");
            if (autoFilterNodes != null && autoFilterNodes.Count > 0)
            {
                string autoFilterRef = ReaderUtils.GetAttribute(autoFilterNodes[0], "ref");
                if (autoFilterRef != null)
                {
                    Range range = new Range(autoFilterRef);
                    worksheet.SetAutoFilter(range.StartAddress.Column, range.EndAddress.Column);
                }
            }
        }

        /// <summary>
        /// Gets the columns of the current worksheet
        /// </summary>
        /// <param name="xmlDocument">XML document of the current worksheet</param>
        /// <param name="worksheet">Currently processed worksheet</param>
        private void GetColumns(XmlDocument xmlDocument, Worksheet worksheet)
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
                    min = ParserUtils.ParseInt(attribute);
                    max = min;
                    indices.Add(min.Value);
                }
                attribute = ReaderUtils.GetAttribute(columnNode, "max");
                if (attribute != null)
                {
                    max = ParserUtils.ParseInt(attribute);
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
                    width = ParserUtils.ParseFloat(attribute);
                }
                attribute = ReaderUtils.GetAttribute(columnNode, "hidden");
                bool hidden = false;
                if (attribute != null)
                {
                    int value = ParserUtils.ParseBinaryBool(attribute);
                    if (value == 1)
                    {
                        hidden = true;
                    }
                }
                attribute = ReaderUtils.GetAttribute(columnNode, "style");
                Style defaultStyle = null;
                if (attribute != null && resolvedStyles.ContainsKey(attribute))
                {
                    defaultStyle = resolvedStyles[attribute];
                }
                foreach (int index in indices)
                {
                    string columnAddress = Cell.ResolveColumnAddress(index - 1); // Transform to zero-based     
                    if (defaultStyle != null)
                    {
                        worksheet.SetColumnDefaultStyle(columnAddress, defaultStyle);
                    }

                    if (width != Worksheet.DEFAULT_COLUMN_WIDTH)
                    {
                        worksheet.SetColumnWidth(columnAddress, width);
                    }
                    if (hidden)
                    {
                        worksheet.AddHiddenColumn(columnAddress);
                    }
                }
            }
        }

        /// <summary>
        /// Reads one cell in a worksheet
        /// </summary>
        /// <param name="rowChild">Current child row as XmlNode</param>
        /// <param name="worksheet">Currently processed worksheet</param>
        private void ReadCell(XmlNode rowChild, Worksheet worksheet)
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
            string key = ParserUtils.ToUpper(address);
            Cell cell = ResolveCellData(value, type, styleNumber, address);
            worksheet.AddCell(cell, address);
            if (styleNumber != null)
            {
                Style style = null;
                this.resolvedStyles.TryGetValue(styleNumber, out style);
                if (style != null)
                {
                    worksheet.Cells[address].SetStyle(style);
                }
            }
        }

        /// <summary>
        /// Resolves the data of a read cell either automatically or conditionally  (import options), transforms it into a cell object and adds it to the data
        /// </summary>
        /// <param name="raw">Raw value as string</param>
        /// <param name="type">Expected data type</param>
        /// <param name="styleNumber">Style number as string (can be null)</param>
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
            if (readerOptions != null)
            {
                if (readerOptions.EnforcedColumnTypes.Count > 0)
                {
                    rawValue = GetEnforcedColumnValue(rawValue, importedType, cellAddress);
                }
                rawValue = GetGloballyEnforcedValue(rawValue, cellAddress);
                rawValue = GetGloballyEnforcedFlagValues(rawValue, cellAddress);
                importedType = ResolveType(rawValue, importedType);
                if (importedType == Cell.CellType.DATE && rawValue is DateTime && (DateTime)rawValue < DataUtils.FIRST_ALLOWED_EXCEL_DATE)
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
            if (address.Row < readerOptions.EnforcingStartRowNumber)
            {
                return data;
            }
            if (readerOptions.EnforceDateTimesAsNumbers)
            {
                if (data is DateTime)
                {
                    data = DataUtils.GetOADateTime((DateTime)data, true);
                }
                else if (data is TimeSpan)
                {
                    data = DataUtils.GetOATime((TimeSpan)data);
                }
            }
            if (readerOptions.EnforceEmptyValuesAsString && data == null)
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
            if (address.Row < readerOptions.EnforcingStartRowNumber)
            {
                return data;
            }
            if (readerOptions.GlobalEnforcingType == ReaderOptions.GlobalType.AllNumbersToDouble)
            {
                object tempDouble = ConvertToDouble(data);
                if (tempDouble != null)
                {
                    return tempDouble;
                }
            }
            else if (readerOptions.GlobalEnforcingType == ReaderOptions.GlobalType.AllNumbersToDecimal)
            {
                object tempDecimal = ConvertToDecimal(data);
                if (tempDecimal != null)
                {
                    return tempDecimal;
                }
            }
            else if (readerOptions.GlobalEnforcingType == ReaderOptions.GlobalType.AllNumbersToInt)
            {
                object tempInt = ConvertToInt(data);
                if (tempInt != null)
                {
                    return tempInt;
                }
            }
            else if (readerOptions.GlobalEnforcingType == ReaderOptions.GlobalType.EverythingToString)
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
            if (address.Row < readerOptions.EnforcingStartRowNumber)
            {
                return data;
            }
            if (!readerOptions.EnforcedColumnTypes.ContainsKey(address.Column))
            {
                return data;
            }
            if (importedTyp == Cell.CellType.FORMULA)
            {
                return data;
            }
            switch (readerOptions.EnforcedColumnTypes[address.Column])
            {
                case ReaderOptions.ColumnType.Numeric:
                    return GetNumericValue(data, importedTyp);
                case ReaderOptions.ColumnType.Decimal:
                    return ConvertToDecimal(data);
                case ReaderOptions.ColumnType.Double:
                    return ConvertToDouble(data);
                case ReaderOptions.ColumnType.Date:
                    return ConvertToDate(data);
                case ReaderOptions.ColumnType.Time:
                    return ConvertToTime(data);
                case ReaderOptions.ColumnType.Bool:
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
                    double tempDouble = converter.ToDouble(DataUtils.INVARIANT_CULTURE);
                    if (tempDouble > (double)decimal.MaxValue || tempDouble < (double)decimal.MinValue)
                    {
                        return data;
                    }
                    else
                    {
                        return converter.ToDecimal(DataUtils.INVARIANT_CULTURE);
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
                    return new decimal(DataUtils.GetOADateTime((DateTime)data));
                case TimeSpan _:
                    return new decimal(DataUtils.GetOATime((TimeSpan)data));
                case string _:
                    decimal dValue;
                    string tempString = (string)data;
                    if (ParserUtils.TryParseDecimal(tempString, out dValue))
                    {
                        return dValue;
                    }
                    DateTime? tempDate = TryParseDate(tempString);
                    if (tempDate != null)
                    {
                        return new decimal(DataUtils.GetOADateTime(tempDate.Value));
                    }
                    TimeSpan? tempTime = TryParseTime(tempString);
                    if (tempTime != null)
                    {
                        return new decimal(DataUtils.GetOATime(tempTime.Value));
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
            double tempDouble;
            switch (data)
            {
                case uint _:
                case long _:
                case ulong _:
                    break;
                case DateTime _:
                    tempDouble = DataUtils.GetOADateTime((DateTime)data, true);
                    return ConvertDoubleToInt(tempDouble);
                case TimeSpan _:
                    tempDouble = DataUtils.GetOATime((TimeSpan)data);
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
                    if (ParserUtils.TryParseInt((string)data, out tempInt2))
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
                    DateTime root = DataUtils.FIRST_ALLOWED_EXCEL_DATE;
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
                    if (date2 != null)
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
            if (readerOptions == null || string.IsNullOrEmpty(readerOptions.DateTimeFormat) || readerOptions.TemporalCultureInfo == null)
            {
                isDateTime = DateTime.TryParse(raw, ReaderOptions.DEFAULT_CULTURE_INFO, DateTimeStyles.None, out dateTime);
            }
            else
            {
                isDateTime = DateTime.TryParseExact(raw, readerOptions.DateTimeFormat, readerOptions.TemporalCultureInfo, DateTimeStyles.None, out dateTime);
            }
            if (isDateTime && dateTime >= DataUtils.FIRST_ALLOWED_EXCEL_DATE && dateTime <= DataUtils.LAST_ALLOWED_EXCEL_DATE)
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
                    if (time != null)
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
            if (readerOptions == null || string.IsNullOrEmpty(readerOptions.TimeSpanFormat) || readerOptions.TemporalCultureInfo == null)
            {
                isTimeSpan = TimeSpan.TryParse(raw, ReaderOptions.DEFAULT_CULTURE_INFO, out timeSpan);
            }
            else
            {
                isTimeSpan = TimeSpan.TryParseExact(raw, readerOptions.TimeSpanFormat, readerOptions.TemporalCultureInfo, out timeSpan);
            }
            if (isTimeSpan && timeSpan.Days >= 0 && timeSpan.Days < DataUtils.MAX_OADATE_VALUE)
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
            if (!ParserUtils.TryParseDouble(raw, out dValue))
            {
                resolvedType = Cell.CellType.STRING;
                return raw;
            }
            if ((valueType == Cell.CellType.DATE && (dValue < DataUtils.MIN_OADATE_VALUE || dValue > DataUtils.MAX_OADATE_VALUE)) || (valueType == Cell.CellType.TIME && (dValue < 0.0 || dValue > DataUtils.MAX_OADATE_VALUE)))
            {
                // fallback to number (cannot be anything else)
                resolvedType = Cell.CellType.NUMBER;
                return GetNumericValue(raw);
            }
            DateTime tempDate = DataUtils.GetDateFromOA(dValue);
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
            if (oaDate is double && (double)oaDate < DataUtils.MAX_OADATE_VALUE)
            {
                DateTime date = DataUtils.GetDateFromOA((double)oaDate);
                if (date >= DataUtils.FIRST_ALLOWED_EXCEL_DATE && date <= DataUtils.LAST_ALLOWED_EXCEL_DATE)
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
            {
                double d = (double)oaDate;
                if (d >= DataUtils.MIN_OADATE_VALUE && d <= DataUtils.MAX_OADATE_VALUE)
                {
                    DateTime date = DataUtils.GetDateFromOA(d);
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
            double dValue = converter.ToDouble(ReaderOptions.DEFAULT_CULTURE_INFO);
            if (dValue > int.MinValue && dValue < int.MaxValue)
            {
                return converter.ToInt32(ReaderOptions.DEFAULT_CULTURE_INFO);
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
            return converter.ToInt32(ReaderOptions.DEFAULT_CULTURE_INFO);
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
                    return ((int)data).ToString(ReaderOptions.DEFAULT_CULTURE_INFO);
                case uint _:
                    return ((uint)data).ToString(ReaderOptions.DEFAULT_CULTURE_INFO);
                case long _:
                    return ((long)data).ToString(ReaderOptions.DEFAULT_CULTURE_INFO);
                case ulong _:
                    return ((ulong)data).ToString(ReaderOptions.DEFAULT_CULTURE_INFO);
                case float _:
                    return ((float)data).ToString(ReaderOptions.DEFAULT_CULTURE_INFO);
                case double _:
                    return ((double)data).ToString(ReaderOptions.DEFAULT_CULTURE_INFO);
                case bool _:
                    return ((bool)data).ToString(ReaderOptions.DEFAULT_CULTURE_INFO);
                case DateTime _:
                    return ((DateTime)data).ToString(readerOptions.DateTimeFormat);
                case TimeSpan _:
                    return ((TimeSpan)data).ToString(readerOptions.TimeSpanFormat);
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
                        return DataUtils.GetOADateTime(tempDate.Value);
                    }
                    TimeSpan? tempTime = TryParseTime(tempString);
                    if (tempTime != null)
                    {
                        return DataUtils.GetOATime(tempTime.Value);
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
                    return DataUtils.GetOADateTime((DateTime)raw);
                case Cell.CellType.TIME:
                    return DataUtils.GetOATime((TimeSpan)raw);
                case Cell.CellType.BOOL:
                    if ((bool)raw)
                    {
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
            bool canBeUint = ParserUtils.TryParseUint(raw, out uiValue);
            bool canBeInt = ParserUtils.TryParseInt(raw, out iValue);
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
            bool canBeUlong = ParserUtils.TryParseUlong(raw, out ulValue);
            bool canBeLong = ParserUtils.TryParseLong(raw, out lValue);
            if (canBeUlong && !canBeLong)
            {
                return ulValue;
            }
            else if (canBeLong)
            {
                return lValue;
            }
            decimal dcValue;
            double dValue;
            float fValue;
            // float section
            if (ParserUtils.TryParseDecimal(raw, out dcValue))
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
            else if (ParserUtils.TryParseFloat(raw, out fValue) && fValue >= float.MinValue && fValue <= float.MaxValue && !float.IsInfinity(fValue))
            {
                return fValue;
            }
            if (ParserUtils.TryParseDouble(raw, out dValue))
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
            if (ParserUtils.TryParseInt(raw, out stringId))
            {
                string resolvedString = SharedStrings.ElementAtOrDefault(stringId);
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
    }
}
