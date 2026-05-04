/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using NanoXLSX.Exceptions;
using NanoXLSX.Interfaces;
using NanoXLSX.Interfaces.Reader;
using NanoXLSX.Registry;
using NanoXLSX.Styles;
using NanoXLSX.Utils;
using NanoXLSX.Utils.Xml;
using static NanoXLSX.Enums.Password;
using IOException = NanoXLSX.Exceptions.IOException;

namespace NanoXLSX.Internal.Readers
{
    /// <summary>
    /// Class representing a reader for worksheets of XLSX files
    /// </summary>
    public class WorksheetReader : IWorksheetReader
    {
        #region privateFields
        private Stream stream;
        private HashSet<string> dateStyles;
        private HashSet<string> timeStyles;
        private Dictionary<string, Style> resolvedStyles;
        private IPasswordReader passwordReader;
        private ReaderOptions readerOptions;
        #endregion

        #region properties

        /// <summary>
        /// Workbook reference where read data is stored (should not be null)
        /// </summary>
        public Workbook Workbook { get; set; }
        /// <summary>
        /// Reader options
        /// </summary>
        public IOptions Options { get; set; }
        /// <summary>
        /// Reference to the <see cref="ReaderPlugInHandler"/>, to be used for post operations in the <see cref="Execute"/> method
        /// </summary>
        public Action<Stream, Workbook, string, IOptions, int?> InlinePluginHandler { get; set; }
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
        /// <param name="inlinePluginHandler">Reference to the a handler action, to be used for post operations in reader methods</param>
        public void Init(Stream stream, Workbook workbook, IOptions readerOptions, Action<Stream, Workbook, string, IOptions, int?> inlinePluginHandler)
        {
            this.stream = stream;
            this.Workbook = workbook;
            this.Options = readerOptions;
            this.readerOptions = readerOptions as ReaderOptions;
            this.InlinePluginHandler = inlinePluginHandler;
            if (dateStyles == null || timeStyles == null || this.resolvedStyles == null)
            {
                StyleReaderContainer styleReaderContainer = workbook.AuxiliaryData.GetData<StyleReaderContainer>(PlugInUUID.StyleReader, PlugInUUID.StyleEntity);
                ProcessStyles(styleReaderContainer);
            }
            if (this.passwordReader == null)
            {
                this.passwordReader = PlugInLoader.GetPlugIn<IPasswordReader>(PlugInUUID.PasswordReader, new LegacyPasswordReader());
                this.passwordReader.Init(PasswordType.WorksheetProtection, (ReaderOptions)readerOptions);
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
                WorksheetDefinition worksheetDefinition = Workbook.AuxiliaryData.GetData<WorksheetDefinition>(PlugInUUID.WorkbookReader, PlugInUUID.WorksheetDefinitionEntity, CurrentWorksheetID);
                Worksheet worksheet = new Worksheet(worksheetDefinition.WorksheetName, worksheetDefinition.SheetID, Workbook)
                {
                    Hidden = worksheetDefinition.Hidden
                };
                using (stream) // Close after processing
                {
                    StringBuilder sb = new StringBuilder();
                    using (XmlReader reader = XmlReader.Create(stream, XmlStreamUtils.CreateSettings()))
                    {
                        while (reader.Read())
                        {
                            if (reader.NodeType != XmlNodeType.Element)
                            {
                                continue;
                            }
                            switch (reader.LocalName.ToLowerInvariant())
                            {
                                case "sheetviews":
                                    GetSheetView(reader, worksheet);
                                    break;
                                case "sheetformatpr":
                                    GetSheetFormats(reader, worksheet);
                                    break;
                                case "cols":
                                    GetColumns(reader, worksheet, readerOptions);
                                    break;
                                case "sheetdata":
                                    GetRows(reader, worksheet, readerOptions, sb);
                                    break;
                                case "sheetprotection":
                                    GetSheetProtection(reader, worksheet);
                                    break;
                                case "mergecells":
                                    GetMergedCells(reader, worksheet);
                                    break;
                                case "autofilter":
                                    GetAutoFilters(reader, worksheet);
                                    break;
                            }
                        }
                        SetWorkbookRelation(worksheet);
                        InlinePluginHandler?.Invoke(stream, Workbook, PlugInUUID.WorksheetInlineReader, Options, CurrentWorksheetID);
                    }
                }
            }
            catch (NotSupportedContentException)
            {
                throw; // rethrow
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
            int selectedWorksheetId = Workbook.AuxiliaryData.GetData<int>(PlugInUUID.WorkbookReader, PlugInUUID.SelectedWorksheetEntity);
            if (selectedWorksheetId == CurrentWorksheetID)
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
            this.dateStyles = new HashSet<string>();
            this.timeStyles = new HashSet<string>();
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
        /// <param name="reader">XmlReader positioned on the sheetData start element</param>
        /// <param name="worksheet">Currently processed worksheet</param>
        /// <param name="readerOptions">Reader options</param>
        /// <param name="sb">Reusable StringBuilder to avoid per-cell allocations</param>
        private void GetRows(XmlReader reader, Worksheet worksheet, ReaderOptions readerOptions, StringBuilder sb)
        {
            using (XmlReader sheetDataReader = reader.ReadSubtree())
            {
                sheetDataReader.Read(); // consume the sheetData open tag
                while (sheetDataReader.Read())
                {
                    if (!XmlStreamUtils.IsElement(sheetDataReader, "row"))
                    {
                        continue;
                    }
                    string rowAttribute = sheetDataReader.GetAttribute("r");
                    if (rowAttribute != null)
                    {
                        int rowNumber = ParserUtils.ParseInt(rowAttribute) - 1; // Transform to zero-based
                        string hiddenAttribute = sheetDataReader.GetAttribute("hidden");
                        if (hiddenAttribute != null && ParserUtils.ParseBinaryBool(hiddenAttribute) == 1)
                        {
                            worksheet.AddHiddenRow(rowNumber);
                        }
                        string heightAttribute = sheetDataReader.GetAttribute("ht");
                        if (heightAttribute != null)
                        {
                            worksheet.RowHeights.Add(rowNumber, GetValidatedHeight(ParserUtils.ParseFloat(heightAttribute), readerOptions));
                        }
                    }
                    if (!sheetDataReader.IsEmptyElement)
                    {
                        using (XmlReader rowReader = sheetDataReader.ReadSubtree())
                        {
                            rowReader.Read(); // consume the row open tag
                            while (rowReader.Read())
                            {
                                if (!XmlStreamUtils.IsElement(rowReader, "c"))
                                {
                                    continue;
                                }
                                ReadCell(rowReader, worksheet, sb);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Gets the selected cells of the current worksheet
        /// </summary>
        /// <param name="reader">XmlReader positioned on the sheetViews start element</param>
        /// <param name="worksheet">Currently processed worksheet</param>
        private static void GetSheetView(XmlReader reader, Worksheet worksheet)
        {
            using (XmlReader subtree = reader.ReadSubtree())
            {
                subtree.Read(); // consume sheetViews
                while (subtree.Read())
                {
                    if (!XmlStreamUtils.IsElement(subtree, "sheetView"))
                    {
                        continue;
                    }
                    string attribute = subtree.GetAttribute("view") ?? string.Empty;
                    worksheet.ViewType = Worksheet.GetSheetViewTypeEnum(attribute);
                    attribute = subtree.GetAttribute("zoomScale");
                    if (attribute != null)
                    {
                        worksheet.ZoomFactor = ParserUtils.ParseInt(attribute);
                    }
                    attribute = subtree.GetAttribute("zoomScaleNormal");
                    if (attribute != null)
                    {
                        worksheet.ZoomFactors[Worksheet.SheetViewType.Normal] = ParserUtils.ParseInt(attribute);
                    }
                    attribute = subtree.GetAttribute("zoomScalePageLayoutView");
                    if (attribute != null)
                    {
                        worksheet.ZoomFactors[Worksheet.SheetViewType.PageLayout] = ParserUtils.ParseInt(attribute);
                    }
                    attribute = subtree.GetAttribute("zoomScaleSheetLayoutView");
                    if (attribute != null)
                    {
                        worksheet.ZoomFactors[Worksheet.SheetViewType.PageBreakPreview] = ParserUtils.ParseInt(attribute);
                    }
                    attribute = subtree.GetAttribute("showGridLines");
                    if (attribute != null)
                    {
                        worksheet.ShowGridLines = ParserUtils.ParseBinaryBool(attribute) == 1;
                    }
                    attribute = subtree.GetAttribute("showRowColHeaders");
                    if (attribute != null)
                    {
                        worksheet.ShowRowColumnHeaders = ParserUtils.ParseBinaryBool(attribute) == 1;
                    }
                    attribute = subtree.GetAttribute("showRuler");
                    if (attribute != null)
                    {
                        worksheet.ShowRuler = ParserUtils.ParseBinaryBool(attribute) == 1;
                    }
                    using (XmlReader sheetViewReader = subtree.ReadSubtree())
                    {
                        sheetViewReader.Read(); // consume sheetView
                        while (sheetViewReader.Read())
                        {
                            if (sheetViewReader.NodeType != XmlNodeType.Element)
                            {
                                continue;
                            }
                            if (XmlStreamUtils.IsElement(sheetViewReader, "selection"))
                            {
                                attribute = sheetViewReader.GetAttribute("sqref");
                                if (attribute != null)
                                {
                                    if (attribute.Contains(" "))
                                    {
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
                            else if (XmlStreamUtils.IsElement(sheetViewReader, "pane"))
                            {
                                SetPaneSplit(sheetViewReader, worksheet);
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
        /// <param name="worksheet">Currently processed worksheet</param>
        private static void CollectSelectedCells(string attribute, Worksheet worksheet)
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
        /// <param name="reader">XmlReader positioned on the pane start element</param>
        /// <param name="worksheet">Currently processed worksheet</param>
        private static void SetPaneSplit(XmlReader reader, Worksheet worksheet)
        {
            string attribute = reader.GetAttribute("state");
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
                if (ParserUtils.ToLower(attribute) == "frozen" || ParserUtils.ToLower(attribute) == "frozensplit")
                {
                    frozenState = true;
                }
                useNumbers = frozenState;
            }
            attribute = reader.GetAttribute("ySplit");
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
            attribute = reader.GetAttribute("xSplit");
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
            attribute = reader.GetAttribute("topLeftCell");
            if (attribute != null)
            {
                topLeftCell = new Address(attribute);
            }
            attribute = reader.GetAttribute("activePane") ?? string.Empty;
            activePane = Worksheet.GetWorksheetPaneEnum(attribute);
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
        /// Gets the sheet protection values of the current worksheet.
        /// Protection attributes are read directly from the start element; the element is then
        /// captured via ReadOuterXml and loaded into a minimal XmlDocument solely to satisfy the
        /// IPasswordReader.ReadXmlAttributes(XmlNode) interface contract.
        /// </summary>
        /// <param name="reader">XmlReader positioned on the sheetProtection start element</param>
        /// <param name="worksheet">Currently processed worksheet</param>
        private void GetSheetProtection(XmlReader reader, Worksheet worksheet)
        {
            int hasProtection = 0;
            hasProtection += ReadSheetProtectionAttribute(reader, Worksheet.SheetProtectionValue.AutoFilter, worksheet);
            hasProtection += ReadSheetProtectionAttribute(reader, Worksheet.SheetProtectionValue.DeleteColumns, worksheet);
            hasProtection += ReadSheetProtectionAttribute(reader, Worksheet.SheetProtectionValue.DeleteRows, worksheet);
            hasProtection += ReadSheetProtectionAttribute(reader, Worksheet.SheetProtectionValue.FormatCells, worksheet);
            hasProtection += ReadSheetProtectionAttribute(reader, Worksheet.SheetProtectionValue.FormatColumns, worksheet);
            hasProtection += ReadSheetProtectionAttribute(reader, Worksheet.SheetProtectionValue.FormatRows, worksheet);
            hasProtection += ReadSheetProtectionAttribute(reader, Worksheet.SheetProtectionValue.InsertColumns, worksheet);
            hasProtection += ReadSheetProtectionAttribute(reader, Worksheet.SheetProtectionValue.InsertHyperlinks, worksheet);
            hasProtection += ReadSheetProtectionAttribute(reader, Worksheet.SheetProtectionValue.InsertRows, worksheet);
            hasProtection += ReadSheetProtectionAttribute(reader, Worksheet.SheetProtectionValue.Objects, worksheet);
            hasProtection += ReadSheetProtectionAttribute(reader, Worksheet.SheetProtectionValue.PivotTables, worksheet);
            hasProtection += ReadSheetProtectionAttribute(reader, Worksheet.SheetProtectionValue.Scenarios, worksheet);
            hasProtection += ReadSheetProtectionAttribute(reader, Worksheet.SheetProtectionValue.SelectLockedCells, worksheet);
            hasProtection += ReadSheetProtectionAttribute(reader, Worksheet.SheetProtectionValue.SelectUnlockedCells, worksheet);
            hasProtection += ReadSheetProtectionAttribute(reader, Worksheet.SheetProtectionValue.Sort, worksheet);
            if (hasProtection > 0)
            {
                worksheet.UseSheetProtection = true;
            }
            string outerXml;
            using (XmlReader subtree = reader.ReadSubtree())
            {
                subtree.MoveToContent();
                outerXml = subtree.ReadOuterXml();
            }
            XmlDocument miniDoc = new XmlDocument { XmlResolver = null };
            miniDoc.LoadXml(outerXml);
            this.passwordReader.ReadXmlAttributes(miniDoc.DocumentElement);
            if (this.passwordReader.PasswordIsSet())
            {
                if (this.passwordReader is LegacyPasswordReader && (this.passwordReader as LegacyPasswordReader).ContemporaryAlgorithmDetected && (readerOptions == null || !readerOptions.IgnoreNotSupportedPasswordAlgorithms))
                {
                    throw new NotSupportedContentException("A not supported, contemporary password algorithm for the worksheet protection was detected. Check possible packages to add support to NanoXLSX, or ignore this error by a reader option");
                }
                worksheet.SheetProtectionPassword.CopyFrom(this.passwordReader);
            }
        }

        /// <summary>
        /// Reads a single sheet protection attribute from the current reader position and registers it on the worksheet if present
        /// </summary>
        /// <param name="reader">XmlReader positioned on the sheetProtection start element</param>
        /// <param name="sheetProtectionValue">Protection value to check</param>
        /// <param name="worksheet">Currently processed worksheet</param>
        /// <returns>1 if the attribute was present, 0 otherwise</returns>
        private static int ReadSheetProtectionAttribute(XmlReader reader, Worksheet.SheetProtectionValue sheetProtectionValue, Worksheet worksheet)
        {
            string attrName = Worksheet.GetSheetProtectionName(sheetProtectionValue);
            if (reader.GetAttribute(attrName) != null)
            {
                worksheet.SheetProtectionValues.Add(sheetProtectionValue);
                return 1;
            }
            return 0;
        }

        /// <summary>
        /// Gets the merged cells of the current worksheet
        /// </summary>
        /// <param name="reader">XmlReader positioned on the mergeCells start element</param>
        /// <param name="worksheet">Currently processed worksheet</param>
        private static void GetMergedCells(XmlReader reader, Worksheet worksheet)
        {
            using (XmlReader subtree = reader.ReadSubtree())
            {
                subtree.Read(); // consume the mergeCells open tag
                while (subtree.Read())
                {
                    if (!XmlStreamUtils.IsElement(subtree, "mergeCell"))
                    {
                        continue;
                    }
                    string attribute = subtree.GetAttribute("ref");
                    if (attribute != null)
                    {
                        worksheet.MergeCells(new Range(attribute));
                    }
                }
            }
        }

        /// <summary>
        /// Gets the sheet format information of the current worksheet
        /// </summary>
        /// <param name="reader">XmlReader positioned on the sheetFormatPr start element</param>
        /// <param name="worksheet">Currently processed worksheet</param>
        private static void GetSheetFormats(XmlReader reader, Worksheet worksheet)
        {
            string attribute = reader.GetAttribute("defaultColWidth");
            if (attribute != null)
            {
                worksheet.DefaultColumnWidth = ParserUtils.ParseFloat(attribute);
            }
            attribute = reader.GetAttribute("defaultRowHeight");
            if (attribute != null)
            {
                worksheet.DefaultRowHeight = ParserUtils.ParseFloat(attribute);
            }
        }

        /// <summary>
        /// Gets the auto filters of the current worksheet
        /// </summary>
        /// <param name="reader">XmlReader positioned on the autoFilter start element</param>
        /// <param name="worksheet">Currently processed worksheet</param>
        private static void GetAutoFilters(XmlReader reader, Worksheet worksheet)
        {
            string autoFilterRef = reader.GetAttribute("ref");
            if (autoFilterRef != null)
            {
                Range range = new Range(autoFilterRef);
                worksheet.SetAutoFilter(range.StartAddress.Column, range.EndAddress.Column);
            }
        }

        /// <summary>
        /// Gets the columns of the current worksheet
        /// </summary>
        /// <param name="reader">XmlReader positioned on the cols start element</param>
        /// <param name="worksheet">Currently processed worksheet</param>
        /// <param name="readerOptions">Reader options</param>
        private void GetColumns(XmlReader reader, Worksheet worksheet, ReaderOptions readerOptions)
        {
            using (XmlReader subtree = reader.ReadSubtree())
            {
                subtree.Read(); // consume the cols open tag
                while (subtree.Read())
                {
                    if (!XmlStreamUtils.IsElement(subtree, "col"))
                    {
                        continue;
                    }
                    int? min = null;
                    int? max = null;
                    List<int> indices = new List<int>();
                    string attribute = subtree.GetAttribute("min");
                    if (attribute != null)
                    {
                        min = ParserUtils.ParseInt(attribute);
                        max = min;
                        indices.Add(min.Value);
                    }
                    attribute = subtree.GetAttribute("max");
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
                    attribute = subtree.GetAttribute("width");
                    float width = Worksheet.DefaultWorksheetColumnWidth;
                    if (attribute != null)
                    {
                        width = ParserUtils.ParseFloat(attribute);
                    }
                    attribute = subtree.GetAttribute("hidden");
                    bool hidden = false;
                    if (attribute != null && ParserUtils.ParseBinaryBool(attribute) == 1)
                    {
                        hidden = true;
                    }
                    attribute = subtree.GetAttribute("style");
                    Style defaultStyle = null;
                    if (attribute != null && resolvedStyles.TryGetValue(attribute, out var attributeValue))
                    {
                        defaultStyle = attributeValue;
                    }
                    foreach (int index in indices)
                    {
                        string columnAddress = Cell.ResolveColumnAddress(index - 1); // Transform to zero-based
                        if (defaultStyle != null)
                        {
                            worksheet.SetColumnDefaultStyle(columnAddress, defaultStyle);
                        }
                        if (width != Worksheet.DefaultWorksheetColumnWidth)
                        {
                            worksheet.SetColumnWidth(columnAddress, GetValidatedWidth(width, readerOptions));
                        }
                        if (hidden)
                        {
                            worksheet.AddHiddenColumn(columnAddress);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Reads one cell in a worksheet
        /// </summary>
        /// <param name="rowReader">XmlReader positioned on the c start element</param>
        /// <param name="worksheet">Currently processed worksheet</param>
        /// <param name="sb">Reusable StringBuilder for inline string content</param>
        private void ReadCell(XmlReader rowReader, Worksheet worksheet, StringBuilder sb)
        {
            string address = rowReader.GetAttribute("r"); // Mandatory
            string type = rowReader.GetAttribute("t");    // can be null
            string styleNumber = rowReader.GetAttribute("s"); // can be null
            string value = "";
            if (!rowReader.IsEmptyElement)
            {
                using (XmlReader cellReader = rowReader.ReadSubtree())
                {
                    cellReader.Read(); // consume <c>
                    while (cellReader.Read())
                    {
                        if (cellReader.NodeType != XmlNodeType.Element)
                        {
                            continue;
                        }
                        if (cellReader.LocalName.Equals("v", StringComparison.OrdinalIgnoreCase) ||
                            cellReader.LocalName.Equals("f", StringComparison.OrdinalIgnoreCase))
                        {
                            value = cellReader.ReadElementContentAsString();
                        }
                        else if (cellReader.LocalName.Equals("is", StringComparison.OrdinalIgnoreCase))
                        {
                            sb.Clear();
                            using (XmlReader isReader = cellReader.ReadSubtree())
                            {
                                isReader.Read(); // consume <is>
                                while (isReader.Read())
                                {
                                    if (isReader.NodeType == XmlNodeType.Element &&
                                        isReader.LocalName.Equals("t", StringComparison.OrdinalIgnoreCase))
                                    {
                                        sb.Append(isReader.ReadElementContentAsString());
                                    }
                                }
                            }
                            value = sb.ToString();
                        }
                    }
                }
            }
            Cell cell = ResolveCellData(value, type, styleNumber, address);
            worksheet.AddCell(cell, address);
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
            Cell.CellType importedType = Cell.CellType.Default;
            object rawValue;
            if (type == "b")
            {
                rawValue = TryParseBool(raw);
                if (rawValue != null)
                {
                    importedType = Cell.CellType.Bool;
                }
                else
                {
                    rawValue = GetNumericValue(raw);
                    if (rawValue != null)
                    {
                        importedType = Cell.CellType.Number;
                    }
                }
            }
            else if (type == "s")
            {
                importedType = Cell.CellType.String;
                rawValue = ResolveSharedString(raw);
            }
            else if (type == "str")
            {
                importedType = Cell.CellType.Formula;
                rawValue = raw;
            }
            else if (type == "inlineStr")
            {
                importedType = Cell.CellType.String;
                rawValue = raw;
            }
            else if (dateStyles.Contains(styleNumber) && (type == null || type == "" || type == "n"))
            {
                rawValue = GetDateTimeValue(raw, Cell.CellType.Date, out importedType);
            }
            else if (timeStyles.Contains(styleNumber) && (type == null || type == "" || type == "n"))
            {
                rawValue = GetDateTimeValue(raw, Cell.CellType.Time, out importedType);
            }
            else
            {
                importedType = Cell.CellType.Number;
                rawValue = GetNumericValue(raw);
            }
            if (rawValue == null && raw == "")
            {
                importedType = Cell.CellType.Empty;
                rawValue = null;
            }
            else if (rawValue == null && raw.Length > 0)
            {
                importedType = Cell.CellType.String;
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
                if (importedType == Cell.CellType.Date && rawValue is DateTime && (DateTime)rawValue < DataUtils.FirstAllowedExcelDate)
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
        private static Cell.CellType ResolveType(object value, Cell.CellType defaultType)
        {
            if (defaultType == Cell.CellType.Formula)
            {
                return defaultType;
            }
            if (value == null)
            {
                return Cell.CellType.Empty;
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
                    return Cell.CellType.Number;
                case DateTime _:
                    return Cell.CellType.Date;
                case TimeSpan _:
                    return Cell.CellType.Time;
                case bool _:
                    return Cell.CellType.Bool;
                default:
                    return Cell.CellType.String;
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
                object tempDouble = ConvertToDouble(data, readerOptions);
                if (tempDouble != null)
                {
                    return tempDouble;
                }
            }
            else if (readerOptions.GlobalEnforcingType == ReaderOptions.GlobalType.AllNumbersToDecimal)
            {
                object tempDecimal = ConvertToDecimal(data, readerOptions);
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
                return ConvertToString(data, readerOptions);
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
            if (!readerOptions.EnforcedColumnTypes.TryGetValue(address.Column, out var columnType))
            {
                return data;
            }
            if (importedTyp == Cell.CellType.Formula)
            {
                return data;
            }
            switch (columnType)
            {
                case ReaderOptions.ColumnType.Numeric:
                    return GetNumericValue(data, importedTyp, readerOptions);
                case ReaderOptions.ColumnType.Decimal:
                    return ConvertToDecimal(data, readerOptions);
                case ReaderOptions.ColumnType.Double:
                    return ConvertToDouble(data, readerOptions);
                case ReaderOptions.ColumnType.Date:
                    return ConvertToDate(data, readerOptions);
                case ReaderOptions.ColumnType.Time:
                    return ConvertToTime(data, readerOptions);
                case ReaderOptions.ColumnType.Bool:
                    return ConvertToBool(data, readerOptions);
                default:
                    return ConvertToString(data, readerOptions);
            }
        }

        /// <summary>
        /// Tries to convert a value to a bool
        /// </summary>
        /// <param name="data">Raw data</param>
        /// <param name="readerOptions">Reader options</param>
        /// <returns>Bool value or original value if not possible to convert</returns>
        private object ConvertToBool(object data, ReaderOptions readerOptions)
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
                    object tempObject = ConvertToDouble(data, readerOptions);
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
        private static bool? TryParseBool(string raw)
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
        /// <param name="readerOptions">Reader options</param>
        /// <returns>Double value or original value if not possible to convert</returns>
        private object ConvertToDouble(object data, ReaderOptions readerOptions)
        {
            object value = ConvertToDecimal(data, readerOptions);
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
        /// <param name="readerOptions">Reader options</param>
        /// <returns>Decimal value or original value if not possible to convert</returns>
        private object ConvertToDecimal(object data, ReaderOptions readerOptions)
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
                    double tempDouble = converter.ToDouble(DataUtils.InvariantCulture);
                    if (tempDouble > (double)decimal.MaxValue || tempDouble < (double)decimal.MinValue)
                    {
                        return data;
                    }
                    else
                    {
                        return converter.ToDecimal(DataUtils.InvariantCulture);
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
                    DateTime? tempDate = TryParseDate(tempString, readerOptions);
                    if (tempDate != null)
                    {
                        return new decimal(DataUtils.GetOADateTime(tempDate.Value));
                    }
                    TimeSpan? tempTime = TryParseTime(tempString, readerOptions);
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
        private static object ConvertToInt(object data)
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
                    int? tempInt = TryConvertDoubleToInt(data);
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
        /// <param name="readerOptions">Reader options</param>
        /// <returns>DateTime value or original value if not possible to convert</returns>
        private object ConvertToDate(object data, ReaderOptions readerOptions)
        {
            switch (data)
            {
                case DateTime _:
                    return data;
                case TimeSpan _:
                    DateTime root = DataUtils.FirstAllowedExcelDate;
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
                    return ConvertDateFromDouble(data, readerOptions);
                case string _:
                    DateTime? date2 = TryParseDate((string)data, readerOptions);
                    if (date2 != null)
                    {
                        return date2.Value;
                    }
                    return ConvertDateFromDouble(data, readerOptions);
            }
            return data;
        }

        /// <summary>
        /// Tries to parse a DateTime instance from a string
        /// </summary>
        /// <param name="raw">String to parse</param>
        /// <param name="readerOptions">Reader options</param>
        /// <returns>DateTime instance or null if not possible to parse</returns>
        private DateTime? TryParseDate(string raw, ReaderOptions readerOptions)
        {
            DateTime dateTime;
            bool isDateTime;
            if (readerOptions == null || string.IsNullOrEmpty(readerOptions.DateTimeFormat) || readerOptions.TemporalCultureInfo == null)
            {
                isDateTime = DateTime.TryParse(raw, ReaderOptions.DefaultCultureInfo, DateTimeStyles.None, out dateTime);
            }
            else
            {
                isDateTime = DateTime.TryParseExact(raw, readerOptions.DateTimeFormat, readerOptions.TemporalCultureInfo, DateTimeStyles.None, out dateTime);
            }
            if (isDateTime && dateTime >= DataUtils.FirstAllowedExcelDate && dateTime <= DataUtils.LastAllowedExcelDate)
            {
                return dateTime;
            }
            return null;
        }

        /// <summary>
        /// Tries to convert a value to a Time (TimeSpan)
        /// </summary>
        /// <param name="data">Raw data</param>
        /// <param name="readerOptions">Reader options</param>
        /// <returns>TimeSpan value or original value if not possible to convert</returns>
        private object ConvertToTime(object data, ReaderOptions readerOptions)
        {
            switch (data)
            {
                case DateTime _:
                    return ConvertTimeFromDouble(data, readerOptions);
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
                    return ConvertTimeFromDouble(data, readerOptions);
                case string _:
                    TimeSpan? time = TryParseTime((string)data, readerOptions);
                    if (time != null)
                    {
                        return time;
                    }
                    return ConvertTimeFromDouble(data, readerOptions);
            }
            return data;
        }

        /// <summary>
        /// Tries to parse a TimeSpan instance from a string
        /// </summary>
        /// <param name="raw">String to parse</param>
        /// <param name="readerOptions">Reader options</param>
        /// <returns>TimeSpan instance or null if not possible to parse</returns>
        private TimeSpan? TryParseTime(string raw, ReaderOptions readerOptions)
        {
            TimeSpan timeSpan;
            bool isTimeSpan;
            if (readerOptions == null || string.IsNullOrEmpty(readerOptions.TimeSpanFormat) || readerOptions.TemporalCultureInfo == null)
            {
                isTimeSpan = TimeSpan.TryParse(raw, ReaderOptions.DefaultCultureInfo, out timeSpan);
            }
            else
            {
                isTimeSpan = TimeSpan.TryParseExact(raw, readerOptions.TimeSpanFormat, readerOptions.TemporalCultureInfo, out timeSpan);
            }
            if (isTimeSpan && timeSpan.Days >= 0 && timeSpan.Days < DataUtils.MaxOADateValue)
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
        private static object GetDateTimeValue(string raw, Cell.CellType valueType, out Cell.CellType resolvedType)
        {
            double dValue;
            if (!ParserUtils.TryParseDouble(raw, out dValue))
            {
                resolvedType = Cell.CellType.String;
                return raw;
            }
            if ((valueType == Cell.CellType.Date && (dValue < DataUtils.MinOADateValue || dValue > DataUtils.MaxOADateValue)) || (valueType == Cell.CellType.Time && (dValue < 0.0 || dValue > DataUtils.MaxOADateValue)))
            {
                // fallback to number (cannot be anything else)
                resolvedType = Cell.CellType.Number;
                return GetNumericValue(raw);
            }
            DateTime tempDate = DataUtils.GetDateFromOA(dValue);
            if (dValue < 1.0)
            {
                tempDate = tempDate.AddDays(1); // Modify wrong 1st date when < 1
            }
            if (valueType == Cell.CellType.Date)
            {
                resolvedType = Cell.CellType.Date;
                return tempDate;
            }
            else
            {
                resolvedType = Cell.CellType.Time;
                return new TimeSpan((int)dValue, tempDate.Hour, tempDate.Minute, tempDate.Second);
            }
        }

        /// <summary>
        /// Tries to convert a date (DateTime) from a double
        /// </summary>
        /// <param name="data">Raw data (may not be a double)</param>
        /// <param name="readerOptions">Reader options</param>
        /// <returns>DateTime value or original value if not possible to convert</returns>
        private object ConvertDateFromDouble(object data, ReaderOptions readerOptions)
        {
            object oaDate = ConvertToDouble(data, readerOptions);
            if (oaDate is double && (double)oaDate < DataUtils.MaxOADateValue)
            {
                DateTime date = DataUtils.GetDateFromOA((double)oaDate);
                if (date >= DataUtils.FirstAllowedExcelDate && date <= DataUtils.LastAllowedExcelDate)
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
        /// <param name="readerOptions">Reader options</param>
        /// <returns>TimeSpan value or original value if not possible to convert</returns>
        private object ConvertTimeFromDouble(object data, ReaderOptions readerOptions)
        {
            object oaDate = ConvertToDouble(data, readerOptions);
            if (oaDate is double)
            {
                double d = (double)oaDate;
                if (d >= DataUtils.MinOADateValue && d <= DataUtils.MaxOADateValue)
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
        private static int? TryConvertDoubleToInt(object data)
        {
            IConvertible converter = data as IConvertible;
            double dValue = converter.ToDouble(ReaderOptions.DefaultCultureInfo);
            if (dValue > int.MinValue && dValue < int.MaxValue)
            {
                return converter.ToInt32(ReaderOptions.DefaultCultureInfo);
            }
            return null;
        }

        /// <summary>
        /// Converts a double to an integer without checks
        /// </summary>
        /// <param name="data">Numeric value</param>
        /// <returns>Converted Value</returns>
        private static int ConvertDoubleToInt(object data)
        {
            IConvertible converter = data as IConvertible;
            return converter.ToInt32(ReaderOptions.DefaultCultureInfo);
        }

        /// <summary>
        /// Converts an arbitrary value to string 
        /// </summary>
        /// <param name="data">Raw data</param>
        /// <param name="readerOptions">Reader options</param>
        /// <returns>Converted string or null in case of null as input</returns>
        private string ConvertToString(object data, ReaderOptions readerOptions)
        {
            switch (data)
            {
                case int _:
                    return ((int)data).ToString(ReaderOptions.DefaultCultureInfo);
                case uint _:
                    return ((uint)data).ToString(ReaderOptions.DefaultCultureInfo);
                case long _:
                    return ((long)data).ToString(ReaderOptions.DefaultCultureInfo);
                case ulong _:
                    return ((ulong)data).ToString(ReaderOptions.DefaultCultureInfo);
                case float _:
                    return ((float)data).ToString(ReaderOptions.DefaultCultureInfo);
                case double _:
                    return ((double)data).ToString(ReaderOptions.DefaultCultureInfo);
                case bool _:
                    return ((bool)data).ToString(ReaderOptions.DefaultCultureInfo);
                case DateTime _:
                    return ((DateTime)data).ToString(readerOptions.DateTimeFormat, ParserUtils.InvariantCulture);
                case TimeSpan _:
                    return ((TimeSpan)data).ToString(readerOptions.TimeSpanFormat, ParserUtils.InvariantCulture);
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
        /// <param name="readerOptions">Reader options</param>
        /// <returns>Converted value or the raw value if not possible to convert</returns>
        private object GetNumericValue(object raw, Cell.CellType importedType, ReaderOptions readerOptions)
        {
            if (raw == null)
            {
                return null;
            }
            object tempObject;
            switch (importedType)
            {
                case Cell.CellType.String:
                    string tempString = raw.ToString();
                    tempObject = GetNumericValue(tempString);
                    if (tempObject != null)
                    {
                        return tempObject;
                    }
                    DateTime? tempDate = TryParseDate(tempString, readerOptions);
                    if (tempDate != null)
                    {
                        return DataUtils.GetOADateTime(tempDate.Value);
                    }
                    TimeSpan? tempTime = TryParseTime(tempString, readerOptions);
                    if (tempTime != null)
                    {
                        return DataUtils.GetOATime(tempTime.Value);
                    }
                    tempObject = ConvertToBool(raw, readerOptions);
                    if (tempObject is bool)
                    {
                        return (bool)tempObject ? 1 : 0;
                    }
                    break;
                case Cell.CellType.Number:
                    return raw;
                case Cell.CellType.Date:
                    return DataUtils.GetOADateTime((DateTime)raw);
                case Cell.CellType.Time:
                    return DataUtils.GetOATime((TimeSpan)raw);
                case Cell.CellType.Bool:
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
        private static object GetNumericValue(string raw)
        {
            bool hasDecimalPoint = raw.Contains(".");

            // Only try integer parsing if there's no decimal point
            if (!hasDecimalPoint)
            {
                // integer section (unchanged)
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
            }

            decimal dcValue;
            double dValue;
            float fValue;

            // Decimal/float section
            if (ParserUtils.TryParseDecimal(raw, out dcValue))
            {
                // Check if the value can be accurately represented as float
                float testFloat = decimal.ToSingle(dcValue);
                decimal backToDecimal = (decimal)testFloat;

                // If converting to float and back preserves the value, use float
                if (dcValue == backToDecimal)
                {
                    return testFloat;
                }
                else
                {
                    // Otherwise use double for better precision
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
        /// Gets the column width according to <see cref="ReaderOptions.EnforceStrictValidation"/>
        /// </summary>
        /// <param name="rawValue">Raw column width</param>
        /// <param name="readerOptions">Reader options</param>
        /// <returns>Modified column width in case <see cref="ReaderOptions.EnforceStrictValidation"/> is set to false, and the raw value was invalid</returns>
        /// <exception cref="WorksheetException">Throws a WorksheetException if the raw value was invalid and <see cref="ReaderOptions.EnforceStrictValidation"/> is set to true</exception>
        private float GetValidatedWidth(float rawValue, ReaderOptions readerOptions)
        {
            if (rawValue < Worksheet.MinColumnWidth)
            {
                if (readerOptions.EnforceStrictValidation)
                {
                    throw new WorksheetException($"The worksheet contains an invalid column width (too small: {rawValue}) value. This error is ignored when disabling the reader option 'EnforceStrictValidation'");
                }
                else
                {
                    return Worksheet.MinColumnWidth;
                }
            }
            else if (rawValue > Worksheet.MaxColumnWidth)
            {
                if (readerOptions.EnforceStrictValidation)
                {
                    throw new WorksheetException($"The worksheet contains an invalid column width (too large: {rawValue}) value.  This error is ignored when disabling the reader option 'EnforceStrictValidation'");
                }
                else
                {
                    return Worksheet.MaxColumnWidth;
                }
            }
            else
            {
                return rawValue;
            }
        }

        /// <summary>
        /// Gets the row height according to <see cref="ReaderOptions.EnforceStrictValidation"/>
        /// </summary>
        /// <param name="rawValue">Raw row height</param>
        /// <param name="readerOptions">Reader options</param>
        /// <returns>Modified row height in case <see cref="ReaderOptions.EnforceStrictValidation"/> is set to false, and the raw value was invalid</returns>
        /// <exception cref="WorksheetException">Throws a WorksheetException if the raw value was invalid and <see cref="ReaderOptions.EnforceStrictValidation"/> is set to true</exception>
        private float GetValidatedHeight(float rawValue, ReaderOptions readerOptions)
        {
            if (rawValue < Worksheet.MinRowHeight)
            {
                if (readerOptions.EnforceStrictValidation)
                {
                    throw new WorksheetException($"The worksheet contains an invalid row height (too small: {rawValue}) value. Consider using the ImportOption 'EnforceValidRowDimensions' to ignore this error.");
                }
                else
                {
                    return Worksheet.MinRowHeight;
                }
            }
            else if (rawValue > Worksheet.MaxRowHeight)
            {
                if (readerOptions.EnforceStrictValidation)
                {
                    throw new WorksheetException($"The worksheet contains an invalid row height (too large: {rawValue}) value. Consider using the ImportOption 'EnforceValidRowDimensions' to ignore this error.");
                }
                else
                {
                    return Worksheet.MaxRowHeight;
                }
            }
            else
            {
                return rawValue;
            }
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
            if (styleNumber != null && resolvedStyles.TryGetValue(styleNumber, out var styleValue))
            {
                cell.SetStyle(styleValue);
            }
            return cell;
        }
        #endregion
    }
}
