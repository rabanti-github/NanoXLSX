﻿/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using NanoXLSX.Exceptions;
using NanoXLSX.Styles;
using FormatException = NanoXLSX.Exceptions.FormatException;
using IOException = NanoXLSX.Exceptions.IOException;

namespace NanoXLSX.LowLevel
{
    /// <summary>
    /// Class for low level handling (XML, formatting, packing)
    /// </summary>
    /// <remarks>This class is only for internal use. Use the high level API (e.g. class Workbook) to manipulate data and create Excel files</remarks>
    class XlsxWriter
    {
        #region constants
        /// <summary>
        /// Threshold, using when floats are compared
        /// </summary>
        private const float FLOAT_THRESHOLD = 0.0001f;
        #endregion


        #region staticFields
        private static DocumentPath WORKBOOK = new DocumentPath("workbook.xml", "xl/");
        private static DocumentPath STYLES = new DocumentPath("styles.xml", "xl/");
        private static DocumentPath APP_PROPERTIES = new DocumentPath("app.xml", "docProps/");
        private static DocumentPath CORE_PROPERTIES = new DocumentPath("core.xml", "docProps/");
        private static DocumentPath SHARED_STRINGS = new DocumentPath("sharedStrings.xml", "xl/");
        #endregion

        #region privateFields
        private CultureInfo culture;
        private Workbook workbook;
        private StyleManager styles;
        private SortedMap sharedStrings;
        private int sharedStringsTotalCount;
        #endregion

        #region constructors
        /// <summary>
        /// Constructor with defined workbook object
        /// </summary>
        /// <param name="workbook">Workbook to process</param>
        public XlsxWriter(Workbook workbook)
        {
            culture = CultureInfo.InvariantCulture;
            this.workbook = workbook;
            sharedStrings = new SortedMap();
            sharedStringsTotalCount = 0;
        }
        #endregion

        #region documentCreation_methods

        /// <summary>
        /// Method to create the app-properties (part of meta data) as raw XML string
        /// </summary>
        /// <returns>Raw XML string</returns>
        private string CreateAppPropertiesDocument()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">");
            sb.Append(CreateAppString());
            sb.Append("</Properties>");
            return sb.ToString();
        }

        /// <summary>
        /// Method to create the core-properties (part of meta data) as raw XML string
        /// </summary>
        /// <returns>Raw XML string</returns>
        private string CreateCorePropertiesDocument()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">");
            sb.Append(CreateCorePropertiesString());
            sb.Append("</cp:coreProperties>");
            return sb.ToString();
        }

        /// <summary>
        /// Method to create shared strings as raw XML string
        /// </summary>
        /// <returns>Raw XML string</returns>
        private string CreateSharedStringsDocument()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"");
            sb.Append(sharedStringsTotalCount.ToString("G", culture));
            sb.Append("\" uniqueCount=\"");
            sb.Append(sharedStrings.Count.ToString("G", culture));
            sb.Append("\">");
            foreach (string str in sharedStrings.Keys)
            {
                AppendSharedString(sb, EscapeXmlChars(str));
            }
            sb.Append("</sst>");
            return sb.ToString();
        }

        /// <summary>
        /// Method to append shared string values and to handle leading or trailing white spaces
        /// </summary>
        /// <param name="sb">StringBuilder instance</param>
        /// <param name="value">Escaped string value (not null)</param>
        private void AppendSharedString(StringBuilder sb, string value)
        {
            int len = value.Length;
            sb.Append("<si>");
            if (len == 0)
            {
                sb.Append("<t></t>");
            }
            else
            {
                if (Char.IsWhiteSpace(value,0) || Char.IsWhiteSpace(value, len-1))
                {
                    sb.Append("<t xml:space=\"preserve\">");
                }
                else
                {
                    sb.Append("<t>");
                }
                sb.Append(NormalizeNewLines(value)).Append("</t>");
            }
            sb.Append("</si>");
        }

        /// <summary>
        /// Method to normalize all newlines to CR+LF
        /// </summary>
        /// <param name="value">Input value</param>
        /// <returns>Normalized value</returns>
        private string NormalizeNewLines(string value)
        {
            if (value == null ||  (!value.Contains('\n') && !value.Contains('\r')))
            {
                return value;
            }
            string normalized = value.Replace("\r\n", "\n").Replace("\r", "\n");
            return normalized.Replace("\n", "\r\n");
        }

        /// <summary>
        /// Method to create a style sheet as raw XML string
        /// </summary>
        /// <returns>Raw XML string</returns>
        /// <exception cref="StyleException">Throws a StyleException if one of the styles cannot be referenced or is null</exception>
        /// <remarks>The UndefinedStyleException should never happen in this state if the internally managed style collection was not tampered. </remarks>
        private string CreateStyleSheetDocument()
        {
            string bordersString = CreateStyleBorderString();
            string fillsString = CreateStyleFillString();
            string fontsString = CreateStyleFontString();
            string numberFormatsString = CreateStyleNumberFormatString();
            string xfsStings = CreateStyleXfsString();
            string mruColorString = CreateMruColorsString();
            int fontCount = styles.GetFontStyleNumber();
            int fillCount = styles.GetFillStyleNumber();
            int styleCount = styles.GetStyleNumber();
            int borderCount = styles.GetBorderStyleNumber();
            StringBuilder sb = new StringBuilder();
            sb.Append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">");
            int numFormatCount = styles.GetNumberFormatStyleNumber();
            if (numFormatCount > 0)
            {
                sb.Append("<numFmts count=\"").Append(numFormatCount.ToString("G", culture)).Append("\">");
                sb.Append(numberFormatsString + "</numFmts>");
            }
            sb.Append("<fonts x14ac:knownFonts=\"1\" count=\"").Append(fontCount.ToString("G", culture)).Append("\">");
            sb.Append(fontsString).Append("</fonts>");
            sb.Append("<fills count=\"").Append(fillCount.ToString("G", culture)).Append("\">");
            sb.Append(fillsString).Append("</fills>");
            sb.Append("<borders count=\"").Append(borderCount.ToString("G", culture)).Append("\">");
            sb.Append(bordersString).Append("</borders>");
            sb.Append("<cellXfs count=\"").Append(styleCount.ToString("G", culture)).Append("\">");
            sb.Append(xfsStings).Append("</cellXfs>");
            if (workbook.WorkbookMetadata != null)
            {
                if (!string.IsNullOrEmpty(mruColorString))
                {
                    sb.Append("<colors>");
                    sb.Append(mruColorString);
                    sb.Append("</colors>");
                }
            }
            sb.Append("</styleSheet>");
            return sb.ToString();
        }

        /// <summary>
        /// Method to create a workbook as raw XML string
        /// </summary>
        /// <returns>Raw XML string</returns>
        /// <exception cref="RangeException">Throws a RangeException if an address was out of range</exception>
        private string CreateWorkbookDocument()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
            if (workbook.SelectedWorksheet > 0 || workbook.Hidden)
            {
                sb.Append("<bookViews><workbookView ");
                if (workbook.Hidden)
                {
                    sb.Append("visibility=\"hidden\"");
                }
                else
                {
                    sb.Append("activeTab=\"").Append(workbook.SelectedWorksheet.ToString("G", culture)).Append("\"");
                }
                sb.Append("/></bookViews>");
            }
            CreateWorkbookProtectionString(sb);
            sb.Append("<sheets>");
            if (workbook.Worksheets.Count > 0)
            {
                foreach (Worksheet item in workbook.Worksheets)
                {
                    sb.Append("<sheet r:id=\"rId").Append(item.SheetID.ToString()).Append("\" sheetId=\"").Append(item.SheetID.ToString()).Append("\" name=\"").Append(EscapeXmlAttributeChars(item.SheetName)).Append("\"");
                    if (item.Hidden)
                    {
                        sb.Append(" state=\"hidden\"");
                    }
                    sb.Append("/>");
                }
            }
            else
            {
                // Fallback on empty workbook
                sb.Append("<sheet r:id=\"rId1\" sheetId=\"1\" name=\"sheet1\"/>");
            }
            sb.Append("</sheets>");
            sb.Append("</workbook>");
            return sb.ToString();
        }

        /// <summary>
        /// Method to create the (sub) part of the workbook protection within the workbook XML document
        /// </summary>
        /// <param name="sb">reference to the StringBuilder</param>
        private void CreateWorkbookProtectionString(StringBuilder sb)
        {
            if (workbook.UseWorkbookProtection)
            {
                sb.Append("<workbookProtection");
                if (workbook.LockWindowsIfProtected)
                {
                    sb.Append(" lockWindows=\"1\"");
                }
                if (workbook.LockStructureIfProtected)
                {
                    sb.Append(" lockStructure=\"1\"");
                }
                if (!string.IsNullOrEmpty(workbook.WorkbookProtectionPassword))
                {
                    sb.Append(" workbookPassword=\"");
                    sb.Append(workbook.WorkbookProtectionPasswordHash);
                    sb.Append("\"");
                }
                sb.Append("/>");
            }
        }

        /// <summary>
        /// Method to create a worksheet part as a raw XML string
        /// </summary>
        /// <param name="worksheet">worksheet object to process</param>
        /// <returns>Raw XML string</returns>
        /// <exception cref="Exceptions.FormatException">Throws a FormatException if a handled date cannot be translated to (Excel internal) OADate</exception>
        private string CreateWorksheetPart(Worksheet worksheet)
        {
            worksheet.RecalculateAutoFilter();
            worksheet.RecalculateColumns();
            StringBuilder sb = new StringBuilder();
            sb.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">");
            if (worksheet.GetLastCellAddress().HasValue && worksheet.GetFirstCellAddress().HasValue)
            {
                sb.Append("<dimension ref=\"").Append(new Range(worksheet.GetFirstCellAddress().Value, worksheet.GetLastCellAddress().Value)).Append("\"/>");
            }
            if (worksheet.SelectedCellRanges.Count > 0 || worksheet.PaneSplitTopHeight != null || worksheet.PaneSplitLeftWidth != null || worksheet.PaneSplitAddress != null ||
                worksheet.Hidden || worksheet.ZoomFactor != 100 || worksheet.ZoomFactors.Count > 1 || !worksheet.ShowGridLines || !worksheet.ShowRuler || !worksheet.ShowRowColumnHeaders || worksheet.ViewType != Worksheet.SheetViewType.normal)
            {
                CreateSheetViewString(worksheet, sb);
            }
            sb.Append("<sheetFormatPr");
            if (!HasPaneSplitting(worksheet))
            {
                // TODO: Find the right calculation to compensate baseColWidth when using pane splitting
                sb.Append(" defaultColWidth=\"")
             .Append(worksheet.DefaultColumnWidth.ToString("G", culture))
                .Append("\"");
            }
            sb.Append(" defaultRowHeight=\"")
             .Append(worksheet.DefaultRowHeight.ToString("G", culture))
             .Append("\" baseColWidth=\"")
             .Append(worksheet.DefaultColumnWidth.ToString("G", culture))
             .Append("\" x14ac:dyDescent=\"0.25\"/>");

            string colDefinitions = CreateColsString(worksheet);
            if (!string.IsNullOrEmpty(colDefinitions))
            {
                sb.Append("<cols>");
                sb.Append(colDefinitions);
                sb.Append("</cols>");
            }
            sb.Append("<sheetData>");
            CreateRowsString(worksheet, sb);
            sb.Append("</sheetData>");
            sb.Append(CreateMergedCellsString(worksheet));
            sb.Append(CreateSheetProtectionString(worksheet));

            if (worksheet.AutoFilterRange != null)
            {
                sb.Append("<autoFilter ref=\"").Append(worksheet.AutoFilterRange.Value.ToString()).Append("\"/>");
            }

            sb.Append("</worksheet>");
            return sb.ToString();
        }

        /// <summary>
        /// Checks whether pane splitting is applied in the given worksheet
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns>True if applied, otherwise false</returns>
        private bool HasPaneSplitting(Worksheet worksheet)
        {
            if (worksheet.PaneSplitLeftWidth == null && worksheet.PaneSplitTopHeight == null && worksheet.PaneSplitAddress == null)
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// Method to create the enclosing part of the rows
        /// </summary>
        /// <param name="worksheet">Worksheet to process</param>
        /// <param name="sb">reference to the StringBuilder</param>
        private void CreateRowsString(Worksheet worksheet, StringBuilder sb)
        {
            List<DynamicRow> cellData = GetSortedSheetData(worksheet);
            string line;
            foreach(DynamicRow row in cellData)
            {
                line = CreateRowString(row, worksheet);
                sb.Append(line);
            }
        }

        /// <summary>
        /// Method to create the (sub) part of the sheet view (selected cells and panes) within the worksheet XML document
        /// </summary>
        /// <param name="worksheet">worksheet object to process</param>
        /// <param name="sb">reference to the StringBuilder</param>
        private void CreateSheetViewString(Worksheet worksheet, StringBuilder sb)
        {
            sb.Append("<sheetViews><sheetView workbookViewId=\"0\"");
            if (workbook.SelectedWorksheet == worksheet.SheetID - 1 && !worksheet.Hidden)
            {
                sb.Append(" tabSelected=\"1\"");
            }
            if (worksheet.ViewType != Worksheet.SheetViewType.normal)
            {
                if (worksheet.ViewType == Worksheet.SheetViewType.pageLayout)
                {
                    if (worksheet.ShowRuler)
                    {
                        sb.Append(" showRuler=\"1\"");
                    }
                    else
                    {
                        sb.Append(" showRuler=\"0\"");
                    }
                    sb.Append(" view=\"pageLayout\"");

                }
                else if (worksheet.ViewType == Worksheet.SheetViewType.pageBreakPreview)
                {
                    sb.Append(" view=\"pageBreakPreview\"");
                }
            }
            if (!worksheet.ShowGridLines)
            {
                sb.Append(" showGridLines=\"0\"");
            }
            if (!worksheet.ShowRowColumnHeaders)
            {
                sb.Append("  showRowColHeaders=\"0\"");
            }
            sb.Append(" zoomScale=\"").Append(worksheet.ZoomFactor.ToString("G", culture)).Append("\"");
            foreach (KeyValuePair<Worksheet.SheetViewType, int> scaleFactor in worksheet.ZoomFactors)
            {
                if (scaleFactor.Key == worksheet.ViewType)
                {
                    continue;
                }
                if (scaleFactor.Key == Worksheet.SheetViewType.normal)
                {
                    sb.Append(" zoomScaleNormal=\"").Append(scaleFactor.Value.ToString("G", culture)).Append("\"");
                }
                else if (scaleFactor.Key == Worksheet.SheetViewType.pageBreakPreview)
                {
                    sb.Append(" zoomScaleSheetLayoutView=\"").Append(scaleFactor.Value.ToString("G", culture)).Append("\"");
                }
                else if (scaleFactor.Key == Worksheet.SheetViewType.pageLayout)
                {
                    sb.Append(" zoomScalePageLayoutView=\"").Append(scaleFactor.Value.ToString("G", culture)).Append("\"");
                }
            }
            sb.Append(">");
            CreatePaneString(worksheet, sb);
            if (worksheet.SelectedCellRanges.Count > 0)
            {
                sb.Append("<selection sqref=\"");
                for (int i = 0; i < worksheet.SelectedCellRanges.Count; i++)
                {
                    sb.Append(worksheet.SelectedCellRanges[i].ToString());
                    if (i < worksheet.SelectedCellRanges.Count - 1)
                    {
                        sb.Append(" ");
                    }
                }
                sb.Append("\" activeCell=\"");
                sb.Append(worksheet.SelectedCellRanges[0].StartAddress.ToString());
                sb.Append("\"/>");
            }
            sb.Append("</sheetView></sheetViews>");
        }

        /// <summary>
        /// Method to create the (sub) part of the pane (splitting and freezing) within the worksheet XML document
        /// </summary>
        /// <param name="worksheet">worksheet object to process</param>
        /// <param name="sb">reference to the StringBuilder</param>
        private void CreatePaneString(Worksheet worksheet, StringBuilder sb)
        {
            if (!HasPaneSplitting(worksheet))
            {
                return;
            }
            sb.Append("<pane");
            bool applyXSplit = false;
            bool applyYSplit = false;
            if (worksheet.PaneSplitAddress != null)
            {
                bool freeze = worksheet.FreezeSplitPanes != null && worksheet.FreezeSplitPanes.Value;
                int xSplit = worksheet.PaneSplitAddress.Value.Column;
                int ySplit = worksheet.PaneSplitAddress.Value.Row;
                if (xSplit > 0)
                {
                    if (freeze)
                    {
                        sb.Append(" xSplit=\"").Append(Utils.ToString(xSplit)).Append("\"");
                    }
                    else
                    {
                        sb.Append(" xSplit=\"").Append(CalculatePaneWidth(worksheet, xSplit).ToString("G", culture)).Append("\"");
                    }
                    applyXSplit = true;
                }
                if (ySplit > 0)
                {
                    if (freeze)
                    {
                        sb.Append(" ySplit=\"").Append(Utils.ToString(ySplit)).Append("\"");
                    }
                    else
                    {
                        sb.Append(" ySplit=\"").Append(CalculatePaneHeight(worksheet, ySplit).ToString("G", culture)).Append("\"");
                    }
                    applyYSplit = true;
                }
                if (freeze && applyXSplit && applyYSplit)
                {
                    sb.Append(" state=\"frozenSplit\"");
                }
                else if (freeze)
                {
                    sb.Append(" state=\"frozen\"");
                }
            }
            else
            {
                if (worksheet.PaneSplitLeftWidth != null)
                {
                    sb.Append(" xSplit=\"").Append(Utils.GetInternalPaneSplitWidth(worksheet.PaneSplitLeftWidth.Value).ToString("G", culture)).Append("\"");
                    applyXSplit = true;
                }
                if (worksheet.PaneSplitTopHeight != null)
                {
                    sb.Append(" ySplit=\"").Append(Utils.GetInternalPaneSplitHeight(worksheet.PaneSplitTopHeight.Value).ToString("G", culture)).Append("\"");
                    applyYSplit = true;
                }
            }
            if ((applyXSplit || applyYSplit) && worksheet.ActivePane != null)
            {
                switch (worksheet.ActivePane.Value)
                {
                    case Worksheet.WorksheetPane.bottomLeft:
                        sb.Append(" activePane=\"bottomLeft\"");
                        break;
                    case Worksheet.WorksheetPane.bottomRight:
                        sb.Append(" activePane=\"bottomRight\"");
                        break;
                    case Worksheet.WorksheetPane.topLeft:
                        sb.Append(" activePane=\"topLeft\"");
                        break;
                    case Worksheet.WorksheetPane.topRight:
                        sb.Append(" activePane=\"topRight\"");
                        break;
                }
            }
            string topLeftCell = worksheet.PaneSplitTopLeftCell.Value.GetAddress();
            sb.Append(" topLeftCell=\"").Append(topLeftCell).Append("\" ");
            sb.Append("/>");
            if (applyXSplit && !applyYSplit)
            {
                sb.Append("<selection pane=\"topRight\" activeCell=\"" + topLeftCell + "\"  sqref=\"" + topLeftCell + "\" />");
            }
            else if (applyYSplit && !applyXSplit)
            {
                sb.Append("<selection pane=\"bottomLeft\" activeCell=\"" + topLeftCell + "\"  sqref=\"" + topLeftCell + "\" />");
            }
            else if (applyYSplit && applyXSplit)
            {
                sb.Append("<selection activeCell=\"" + topLeftCell + "\"  sqref=\"" + topLeftCell + "\" />");
            }
        }

        /// <summary>
        /// Method to calculate the pane height, based on the number of rows
        /// </summary>
        /// <param name="worksheet">worksheet object to get the row definitions from</param>
        /// <param name="numberOfRows">Number of rows from the top to the split position</param>
        /// <returns>Internal height from the top of the worksheet to the pane split position</returns>
        private float CalculatePaneHeight(Worksheet worksheet, int numberOfRows)
        {
            float height = 0;
            for (int i = 0; i < numberOfRows; i++)
            {
                if (worksheet.RowHeights.ContainsKey(i))
                {
                    height += Utils.GetInternalRowHeight(worksheet.RowHeights[i]);
                }
                else
                {
                    height += Utils.GetInternalRowHeight(Worksheet.DEFAULT_ROW_HEIGHT);
                }
            }
            return Utils.GetInternalPaneSplitHeight(height);
        }

        /// <summary>
        /// Method to calculate the pane width, based on the number of columns
        /// </summary>
        /// <param name="worksheet">worksheet object to get the column definitions from</param>
        /// <param name="numberOfColumns">Number of columns from the left to the split position</param>
        /// <returns>Internal width from the left of the worksheet to the pane split position</returns>
        private float CalculatePaneWidth(Worksheet worksheet, int numberOfColumns)
        {
            float width = 0;
            for (int i = 0; i < numberOfColumns; i++)
            {
                if (worksheet.Columns.ContainsKey(i))
                {
                    width += Utils.GetInternalColumnWidth(worksheet.Columns[i].Width);
                }
                else
                {
                    width += Utils.GetInternalColumnWidth(Worksheet.DEFAULT_COLUMN_WIDTH);
                }
            }
            // Add padding of 75 per column
            return Utils.GetInternalPaneSplitWidth(width) + ((numberOfColumns - 1) * 0f);
        }

        /// <summary>
        /// Method to save the workbook
        /// </summary>
        /// <exception cref="Exceptions.IOException">Throws IOException in case of an error</exception>
        /// <exception cref="RangeException">Throws a RangeException if the start or end address of a handled cell range was out of range</exception>
        /// <exception cref="Exceptions.FormatException">Throws a FormatException if a handled date cannot be translated to (Excel internal) OADate</exception>
        /// <exception cref="StyleException">Throws a StyleException if one of the styles of the workbook cannot be referenced or is null</exception>
        /// <remarks>The StyleException should never happen in this state if the internally managed style collection was not tampered. </remarks>
        public void Save()
        {
            try
            {
                FileStream fs = new FileStream(workbook.Filename, FileMode.Create);
                SaveAsStream(fs);

            }
            catch (Exception e)
            {
                throw new IOException("An error occurred while saving. See inner exception for details: " + e.Message, e);
            }
        }

        /// <summary>
        /// Method to save the workbook asynchronous.
        /// </summary>
        /// <remarks>Possible Exceptions are <see cref="Exceptions.IOException">IOException</see>, <see cref="RangeException">RangeException</see>, <see cref="Exceptions.FormatException"></see> and <see cref="StyleException">StyleException</see>. These exceptions may not emerge directly if using the async method since async/await adds further abstraction layers.</remarks>
        /// <returns>Async Task</returns>
        public async Task SaveAsync()
        {
            await Task.Run(() => { Save(); });
        }

        /// <summary>
        /// Method to save the workbook as stream
        /// </summary>
        /// <param name="stream">Writable stream as target</param>
        /// <param name="leaveOpen">Optional parameter to keep the stream open after writing (used for MemoryStreams; default is false)</param>
        /// <exception cref="IOException">Throws IOException in case of an error</exception>
        /// <exception cref="RangeException">Throws a RangeException if the start or end address of a handled cell range was out of range</exception>
        /// <exception cref="FormatException">Throws a FormatException if a handled date cannot be translated to (Excel internal) OADate</exception>
        /// <exception cref="StyleException">Throws a StyleException if one of the styles of the workbook cannot be referenced or is null</exception>
        /// <remarks>The StyleException should never happen in this state if the internally managed style collection was not tampered. </remarks>
        public void SaveAsStream(Stream stream, bool leaveOpen = false)
        {
            workbook.ResolveMergedCells();
            this.styles = StyleManager.GetManagedStyles(workbook);
            DocumentPath sheetPath;
            List<Uri> sheetURIs = new List<Uri>();
            try
            {
                using (Package p = Package.Open(stream, FileMode.Create))
                {
                    Uri workbookUri = new Uri(WORKBOOK.GetFullPath(), UriKind.Relative);
                    Uri stylesheetUri = new Uri(STYLES.GetFullPath(), UriKind.Relative);
                    Uri appPropertiesUri = new Uri(APP_PROPERTIES.GetFullPath(), UriKind.Relative);
                    Uri corePropertiesUri = new Uri(CORE_PROPERTIES.GetFullPath(), UriKind.Relative);
                    Uri sharedStringsUri = new Uri(SHARED_STRINGS.GetFullPath(), UriKind.Relative);

                    PackagePart pp = p.CreatePart(workbookUri, @"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", CompressionOption.Normal);
                    p.CreateRelationship(pp.Uri, TargetMode.Internal, @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "rId1");
                    p.CreateRelationship(corePropertiesUri, TargetMode.Internal, @"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "rId2"); //!
                    p.CreateRelationship(appPropertiesUri, TargetMode.Internal, @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", "rId3"); //!

                    AppendXmlToPackagePart(CreateWorkbookDocument(), pp);
                    int idCounter;
                    if (workbook.Worksheets.Count > 0)
                    {
                        idCounter = workbook.Worksheets.Count + 1;
                    }
                    else
                    {
                        //  Fallback on empty workbook
                        idCounter = 2;
                    }
                    pp.CreateRelationship(stylesheetUri, TargetMode.Internal, @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "rId" + idCounter);
                    pp.CreateRelationship(sharedStringsUri, TargetMode.Internal, @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings", "rId" + (idCounter + 1));

                    if (workbook.Worksheets.Count > 0)
                    {
                        foreach (Worksheet item in workbook.Worksheets)
                        {
                            sheetPath = new DocumentPath("sheet" + item.SheetID + ".xml", "xl/worksheets");
                            sheetURIs.Add(new Uri(sheetPath.GetFullPath(), UriKind.Relative));
                            pp.CreateRelationship(sheetURIs[sheetURIs.Count - 1], TargetMode.Internal, @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", "rId" + item.SheetID);
                        }
                    }
                    else
                    {
                        //  Fallback on empty workbook
                        sheetPath = new DocumentPath("sheet1.xml", "xl/worksheets");
                        sheetURIs.Add(new Uri(sheetPath.GetFullPath(), UriKind.Relative));
                        pp.CreateRelationship(sheetURIs[sheetURIs.Count - 1], TargetMode.Internal, @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", "rId1");
                    }

                    pp = p.CreatePart(stylesheetUri, @"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml", CompressionOption.Normal);
                    AppendXmlToPackagePart(CreateStyleSheetDocument(), pp);

                    int i = 0;
                    if (workbook.Worksheets.Count > 0)
                    {
                        foreach (Worksheet item in workbook.Worksheets)
                        {
                            pp = p.CreatePart(sheetURIs[i], @"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", CompressionOption.Normal);
                            i++;
                            AppendXmlToPackagePart(CreateWorksheetPart(item), pp);
                        }
                    }
                    else
                    {
                        pp = p.CreatePart(sheetURIs[i], @"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", CompressionOption.Normal);
                        i++;
                        AppendXmlToPackagePart(CreateWorksheetPart(new Worksheet("sheet1")), pp);
                    }
                    pp = p.CreatePart(sharedStringsUri, @"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml", CompressionOption.Normal);
                    AppendXmlToPackagePart(CreateSharedStringsDocument(), pp);

                    if (workbook.WorkbookMetadata != null)
                    {
                        pp = p.CreatePart(appPropertiesUri, @"application/vnd.openxmlformats-officedocument.extended-properties+xml", CompressionOption.Normal);
                        AppendXmlToPackagePart(CreateAppPropertiesDocument(), pp);
                        pp = p.CreatePart(corePropertiesUri, @"application/vnd.openxmlformats-package.core-properties+xml", CompressionOption.Normal);
                        AppendXmlToPackagePart(CreateCorePropertiesDocument(), pp);
                    }
                    p.Flush();
                    p.Close();
                    if (!leaveOpen)
                    {
                        stream.Close();
                    }
                }
            }
            catch (Exception e)
            {
                throw new IOException("An error occurred while saving. See inner exception for details: " + e.Message, e);
            }
        }

        /// <summary>
        /// Method to save the workbook as stream asynchronous.
        /// </summary>
        /// <param name="stream">Writable stream as target</param>
        /// <param name="leaveOpen">Optional parameter to keep the stream open after writing (used for MemoryStreams; default is false)</param>
        /// <remarks>Possible Exceptions are <see cref="IOException">IOException</see>, <see cref="RangeException">RangeException</see>, <see cref="FormatException"></see> and <see cref="StyleException">StyleException</see>. These exceptions may not emerge directly if using the async method since async/await adds further abstraction layers.</remarks>
        /// <returns>Async Task</returns>
        public async Task SaveAsStreamAsync(Stream stream, bool leaveOpen = false)
        {
            await Task.Run(() => { SaveAsStream(stream, leaveOpen); });
        }

        #endregion

        #region documentUtil_methods

        /// <summary>
        /// Method to append a simple XML tag with an enclosed value to the passed StringBuilder
        /// </summary>
        /// <param name="sb">StringBuilder to append</param>
        /// <param name="value">Value of the XML element</param>
        /// <param name="tagName">Tag name of the XML element</param>
        /// <param name="nameSpace">Optional XML name space. Can be empty or null</param>
        private void AppendXmlTag(StringBuilder sb, string value, string tagName, string nameSpace)
        {
            if (string.IsNullOrEmpty(value)) { return; }
            bool hasNoNs = string.IsNullOrEmpty(nameSpace);
            sb.Append('<');
            if (!hasNoNs)
            {
                sb.Append(nameSpace);
                sb.Append(':');
            }
            sb.Append(tagName).Append(">");
            sb.Append(EscapeXmlChars(value));
            sb.Append("</");
            if (!hasNoNs)
            {
                sb.Append(nameSpace);
                sb.Append(':');
            }
            sb.Append(tagName);
            sb.Append('>');
        }

        /// <summary>
        /// Writes raw XML strings into the passed Package Part
        /// </summary>
        /// <param name="doc">document as raw XML string</param>
        /// <param name="pp">Package part to append the XML data</param>
        private void AppendXmlToPackagePart(string doc, PackagePart pp)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                using (XmlWriter writer = XmlWriter.Create(ms))
                {
                    writer.WriteProcessingInstruction("xml", "version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"");
                    writer.WriteRaw(doc);
                    writer.Flush();
                    ms.Position = 0;
                    ms.CopyTo(pp.GetStream());
                    ms.Flush();
                }
            }
        }

        /// <summary>
        /// Method to create the XML string for the app-properties document
        /// </summary>
        /// <returns>String with formatted XML data</returns>
        private string CreateAppString()
        {
            Metadata md = workbook.WorkbookMetadata;
            StringBuilder sb = new StringBuilder();
            AppendXmlTag(sb, "0", "TotalTime", null);
            AppendXmlTag(sb, md.Application, "Application", null);
            AppendXmlTag(sb, "0", "DocSecurity", null);
            AppendXmlTag(sb, "false", "ScaleCrop", null);
            AppendXmlTag(sb, md.Manager, "Manager", null);
            AppendXmlTag(sb, md.Company, "Company", null);
            AppendXmlTag(sb, "false", "LinksUpToDate", null);
            AppendXmlTag(sb, "false", "SharedDoc", null);
            AppendXmlTag(sb, md.HyperlinkBase, "HyperlinkBase", null);
            AppendXmlTag(sb, "false", "HyperlinksChanged", null);
            AppendXmlTag(sb, md.ApplicationVersion, "AppVersion", null);
            return sb.ToString();
        }

        /// <summary>
        /// Method to create the columns as XML string. This is used to define the width of columns
        /// </summary>
        /// <param name="worksheet">Worksheet to process</param>
        /// <returns>String with formatted XML data</returns>
        private string CreateColsString(Worksheet worksheet)
        {
            if (worksheet.Columns.Count > 0)
            {
                string col;
                string hidden = "";
                StringBuilder sb = new StringBuilder();
                foreach (KeyValuePair<int, Column> column in worksheet.Columns)
                {
                    if (Math.Abs(column.Value.Width - worksheet.DefaultColumnWidth) < FLOAT_THRESHOLD && !column.Value.IsHidden && column.Value.DefaultColumnStyle == null) { continue; }
                    if (worksheet.Columns.ContainsKey(column.Key))
                    {
                        if (worksheet.Columns[column.Key].IsHidden)
                        {
                            hidden = " hidden=\"1\"";
                        }
                    }
                    col = (column.Key + 1).ToString("G", culture); // Add 1 for Address
                    float width = Utils.GetInternalColumnWidth(column.Value.Width);
                    sb.Append("<col customWidth=\"1\" width=\"").Append(width.ToString("G", culture)).Append("\" max=\"").Append(col).Append("\" min=\"").Append(col).Append("\"");
                    if (column.Value.DefaultColumnStyle != null)
					{
                        sb.Append(" style=\"").Append(column.Value.DefaultColumnStyle.InternalID.Value.ToString("G", culture)).Append("\"");
					}
                    sb.Append(hidden).Append("/>");
                }
                string value = sb.ToString();
                if (value.Length > 0)
                {
                    return value;
                }
                return string.Empty;
            }
            return string.Empty;
        }

        /// <summary>
        /// Method to create the XML string for the core-properties document
        /// </summary>
        /// <returns>String with formatted XML data</returns>
        private string CreateCorePropertiesString()
        {
            Metadata md = workbook.WorkbookMetadata;
            StringBuilder sb = new StringBuilder();
            AppendXmlTag(sb, md.Title, "title", "dc");
            AppendXmlTag(sb, md.Subject, "subject", "dc");
            AppendXmlTag(sb, md.Creator, "creator", "dc");
            AppendXmlTag(sb, md.Creator, "lastModifiedBy", "cp");
            AppendXmlTag(sb, md.Keywords, "keywords", "cp");
            AppendXmlTag(sb, md.Description, "description", "dc");
            string time = DateTime.Now.ToString("yyyy-MM-ddThh:mm:ssZ", culture);
            sb.Append("<dcterms:created xsi:type=\"dcterms:W3CDTF\">").Append(time).Append("</dcterms:created>");
            sb.Append("<dcterms:modified xsi:type=\"dcterms:W3CDTF\">").Append(time).Append("</dcterms:modified>");

            AppendXmlTag(sb, md.Category, "category", "cp");
            AppendXmlTag(sb, md.ContentStatus, "contentStatus", "cp");

            return sb.ToString();
        }

        /// <summary>
        /// Method to create the merged cells string of the passed worksheet
        /// </summary>
        /// <param name="sheet">Worksheet to process</param>
        /// <returns>Formatted string with merged cell ranges</returns>
        private string CreateMergedCellsString(Worksheet sheet)
        {
            if (sheet.MergedCells.Count < 1)
            {
                return string.Empty;
            }
            StringBuilder sb = new StringBuilder();
            sb.Append("<mergeCells count=\"").Append(sheet.MergedCells.Count.ToString("G", culture)).Append("\">");
            foreach (KeyValuePair<string, Range> item in sheet.MergedCells)
            {
                sb.Append("<mergeCell ref=\"").Append(item.Value.ToString()).Append("\"/>");
            }
            sb.Append("</mergeCells>");
            return sb.ToString();
        }

        /// <summary>
        /// Method to create a row string
        /// </summary>
        /// <param name="dynamicRow">Dynamic row with List of cells, heights and hidden states</param>
        /// <param name="worksheet">Worksheet to process</param>
        /// <returns>Formatted row string</returns>
        /// <exception cref="Exceptions.FormatException">Throws a FormatException if a handled date cannot be translated to (Excel internal) OADate</exception>
        private string CreateRowString(DynamicRow dynamicRow, Worksheet worksheet)
        {
            int rowNumber = dynamicRow.RowNumber;
            string height = "";
            string hidden = "";
            if (worksheet.RowHeights.ContainsKey(rowNumber) && Math.Abs(worksheet.RowHeights[rowNumber] - worksheet.DefaultRowHeight) > FLOAT_THRESHOLD)
            {
               height = " x14ac:dyDescent=\"0.25\" customHeight=\"1\" ht=\"" + Utils.GetInternalRowHeight(worksheet.RowHeights[rowNumber]).ToString("G", culture) + "\"";
            }
            if (worksheet.HiddenRows.ContainsKey(rowNumber) && worksheet.HiddenRows[rowNumber])
            {
               hidden = " hidden=\"1\"";
            }
            StringBuilder sb = new StringBuilder();
            sb.Append("<row r=\"").Append((rowNumber + 1).ToString()).Append("\"").Append(height).Append(hidden).Append(">");
            string typeAttribute;
            string styleDef;
            string typeDef;
            string valueDef = "";
            bool boolValue;

            int col = 0;
            foreach (Cell item in dynamicRow.CellDefinitions)
            {
                // Data type must be resolved
                typeDef = " ";
                if (item.CellStyle != null)
                {
                    styleDef = " s=\"" + item.CellStyle.InternalID.Value.ToString("G", culture) + "\" ";
                }
                else
                {
                    styleDef = "";
                }
                if (item.DataType == Cell.CellType.BOOL)
                {
                    typeAttribute = "b";
                    typeDef = " t=\"" + typeAttribute + "\" ";
                    boolValue = (bool)item.Value;
                    if (boolValue) { valueDef = "1"; }
                    else { valueDef = "0"; }

                }
                // Number casting
                else if (item.DataType == Cell.CellType.NUMBER)
                {
                    typeAttribute = "n";
                    typeDef = " t=\"" + typeAttribute + "\" ";
                    Type t = item.Value.GetType();

                    if (t == typeof(byte)) { valueDef = ((byte)item.Value).ToString("G", culture); }
                    else if (t == typeof(sbyte)) { valueDef = ((sbyte)item.Value).ToString("G", culture); }
                    else if (t == typeof(decimal)) { valueDef = ((decimal)item.Value).ToString("G", culture); }
                    else if (t == typeof(double)) { valueDef = ((double)item.Value).ToString("G", culture); }
                    else if (t == typeof(float)) { valueDef = ((float)item.Value).ToString("G", culture); }
                    else if (t == typeof(int)) { valueDef = ((int)item.Value).ToString("G", culture); }
                    else if (t == typeof(uint)) { valueDef = ((uint)item.Value).ToString("G", culture); }
                    else if (t == typeof(long)) { valueDef = ((long)item.Value).ToString("G", culture); }
                    else if (t == typeof(ulong)) { valueDef = ((ulong)item.Value).ToString("G", culture); }
                    else if (t == typeof(short)) { valueDef = ((short)item.Value).ToString("G", culture); }
                    else if (t == typeof(ushort)) { valueDef = ((ushort)item.Value).ToString("G", culture); }
                }
                // Date parsing
                else if (item.DataType == Cell.CellType.DATE)
                {
                    DateTime date = (DateTime)item.Value;
                    valueDef = Utils.GetOADateTimeString(date);
                }
                // Time parsing
                else if (item.DataType == Cell.CellType.TIME)
                {
                    TimeSpan time = (TimeSpan)item.Value;
                    valueDef = Utils.GetOATimeString(time);
                }
                else
                {
                    if (item.Value == null)
                    {
                       typeAttribute = null;
                       valueDef = null;
                    }
                    else // Handle sharedStrings
                    {
                        if (item.DataType == Cell.CellType.FORMULA)
                        {
                            typeAttribute = "str";
                            valueDef = item.Value.ToString();
                        }
                        else
                        {
                            typeAttribute = "s";
                            valueDef = sharedStrings.Add(item.Value.ToString(), sharedStrings.Count.ToString("G", culture));
                            sharedStringsTotalCount++;
                        }
                    }
                    typeDef = " t=\"" + typeAttribute + "\" ";
                }
                if (item.DataType != Cell.CellType.EMPTY)
                {
                    sb.Append("<c r=\"").Append(item.CellAddress).Append("\"").Append(typeDef).Append(styleDef).Append(">");
                    if (item.DataType == Cell.CellType.FORMULA)
                    {
                        sb.Append("<f>").Append(EscapeXmlChars(item.Value.ToString())).Append("</f>");
                    }
                    else
                    {
                        sb.Append("<v>").Append(EscapeXmlChars(valueDef)).Append("</v>");
                    }
                    sb.Append("</c>");
                }
                else if (valueDef == null || item.DataType == Cell.CellType.EMPTY) // Empty cell
                {
                    sb.Append("<c r=\"").Append(item.CellAddress).Append("\"").Append(styleDef).Append("/>");
                }
                col++;
            }
            sb.Append("</row>");
            return sb.ToString();
        }

        /// <summary>
        /// Method to create the protection string of the passed worksheet
        /// </summary>
        /// <param name="sheet">Worksheet to process</param>
        /// <returns>Formatted string with protection statement of the worksheet</returns>
        private string CreateSheetProtectionString(Worksheet sheet)
        {
            if (!sheet.UseSheetProtection)
            {
                return string.Empty;
            }
            Dictionary<Worksheet.SheetProtectionValue, int> actualLockingValues = new Dictionary<Worksheet.SheetProtectionValue, int>();
            if (sheet.SheetProtectionValues.Count == 0)
            {
                actualLockingValues.Add(Worksheet.SheetProtectionValue.selectLockedCells, 1);
                actualLockingValues.Add(Worksheet.SheetProtectionValue.selectUnlockedCells, 1);
            }
            if (!sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.objects))
            {
                actualLockingValues.Add(Worksheet.SheetProtectionValue.objects, 1);
            }
            if (!sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.scenarios))
            {
                actualLockingValues.Add(Worksheet.SheetProtectionValue.scenarios, 1);
            }
            if (!sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.selectLockedCells))
            {
                if (!actualLockingValues.ContainsKey(Worksheet.SheetProtectionValue.selectLockedCells))
                {
                    actualLockingValues.Add(Worksheet.SheetProtectionValue.selectLockedCells, 1);
                }
            }
            if (!sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.selectUnlockedCells) || !sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.selectLockedCells))
            {
                if (!actualLockingValues.ContainsKey(Worksheet.SheetProtectionValue.selectUnlockedCells))
                {
                    actualLockingValues.Add(Worksheet.SheetProtectionValue.selectUnlockedCells, 1);
                }
            }
            if (sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.formatCells)) { actualLockingValues.Add(Worksheet.SheetProtectionValue.formatCells, 0); }
            if (sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.formatColumns)) { actualLockingValues.Add(Worksheet.SheetProtectionValue.formatColumns, 0); }
            if (sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.formatRows)) { actualLockingValues.Add(Worksheet.SheetProtectionValue.formatRows, 0); }
            if (sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.insertColumns)) { actualLockingValues.Add(Worksheet.SheetProtectionValue.insertColumns, 0); }
            if (sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.insertRows)) { actualLockingValues.Add(Worksheet.SheetProtectionValue.insertRows, 0); }
            if (sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.insertHyperlinks)) { actualLockingValues.Add(Worksheet.SheetProtectionValue.insertHyperlinks, 0); }
            if (sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.deleteColumns)) { actualLockingValues.Add(Worksheet.SheetProtectionValue.deleteColumns, 0); }
            if (sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.deleteRows)) { actualLockingValues.Add(Worksheet.SheetProtectionValue.deleteRows, 0); }
            if (sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.sort)) { actualLockingValues.Add(Worksheet.SheetProtectionValue.sort, 0); }
            if (sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.autoFilter)) { actualLockingValues.Add(Worksheet.SheetProtectionValue.autoFilter, 0); }
            if (sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.pivotTables)) { actualLockingValues.Add(Worksheet.SheetProtectionValue.pivotTables, 0); }
            StringBuilder sb = new StringBuilder();
            sb.Append("<sheetProtection");
            string temp;
            foreach (KeyValuePair<Worksheet.SheetProtectionValue, int> item in actualLockingValues)
            {
                    temp = Enum.GetName(typeof(Worksheet.SheetProtectionValue), item.Key); // Note! If the enum names differs from the OOXML definitions, this method will cause invalid OOXML entries
                    sb.Append(" ").Append(temp).Append("=\"").Append(item.Value.ToString("G", culture)).Append("\"");
            }
            if (!string.IsNullOrEmpty(sheet.SheetProtectionPasswordHash))
            {
                sb.Append(" password=\"").Append(sheet.SheetProtectionPasswordHash).Append("\"");
            }
            sb.Append(" sheet=\"1\"/>");
            return sb.ToString();
        }

        /// <summary>
        /// Method to create the XML string for the border part of the style sheet document
        /// </summary>
        /// <returns>String with formatted XML data</returns>
        private string CreateStyleBorderString()
        {
            Border[] borderStyles = styles.GetBorders();
            StringBuilder sb = new StringBuilder();
            foreach (Border item in borderStyles)
            {
                if (item.DiagonalDown && !item.DiagonalUp) { sb.Append("<border diagonalDown=\"1\">"); }
                else if (!item.DiagonalDown && item.DiagonalUp) { sb.Append("<border diagonalUp=\"1\">"); }
                else if (item.DiagonalDown && item.DiagonalUp) { sb.Append("<border diagonalDown=\"1\" diagonalUp=\"1\">"); }
                else { sb.Append("<border>"); }

                if (item.LeftStyle != Border.StyleValue.none)
                {
                    sb.Append("<left style=\"" + Border.GetStyleName(item.LeftStyle) + "\">");
                    if (!string.IsNullOrEmpty(item.LeftColor)) { sb.Append("<color rgb=\"").Append(item.LeftColor).Append("\"/>"); }
                    else { sb.Append("<color auto=\"1\"/>"); }
                    sb.Append("</left>");
                }
                else
                {
                    sb.Append("<left/>");
                }
                if (item.RightStyle != Border.StyleValue.none)
                {
                    sb.Append("<right style=\"").Append(Border.GetStyleName(item.RightStyle)).Append("\">");
                    if (!string.IsNullOrEmpty(item.RightColor)) { sb.Append("<color rgb=\"").Append(item.RightColor).Append("\"/>"); }
                    else { sb.Append("<color auto=\"1\"/>"); }
                    sb.Append("</right>");
                }
                else
                {
                    sb.Append("<right/>");
                }
                if (item.TopStyle != Border.StyleValue.none)
                {
                    sb.Append("<top style=\"").Append(Border.GetStyleName(item.TopStyle)).Append("\">");
                    if (!string.IsNullOrEmpty(item.TopColor)) { sb.Append("<color rgb=\"").Append(item.TopColor).Append("\"/>"); }
                    else { sb.Append("<color auto=\"1\"/>"); }
                    sb.Append("</top>");
                }
                else
                {
                    sb.Append("<top/>");
                }
                if (item.BottomStyle != Border.StyleValue.none)
                {
                    sb.Append("<bottom style=\"").Append(Border.GetStyleName(item.BottomStyle)).Append("\">");
                    if (!string.IsNullOrEmpty(item.BottomColor)) { sb.Append("<color rgb=\"").Append(item.BottomColor).Append("\"/>"); }
                    else { sb.Append("<color auto=\"1\"/>"); }
                    sb.Append("</bottom>");
                }
                else
                {
                    sb.Append("<bottom/>");
                }
                if (item.DiagonalStyle != Border.StyleValue.none)
                {
                    sb.Append("<diagonal style=\"").Append(Border.GetStyleName(item.DiagonalStyle)).Append("\">");
                    if (!string.IsNullOrEmpty(item.DiagonalColor)) { sb.Append("<color rgb=\"").Append(item.DiagonalColor).Append("\"/>"); }
                    else { sb.Append("<color auto=\"1\"/>"); }
                    sb.Append("</diagonal>");
                }
                else
                {
                    sb.Append("<diagonal/>");
                }

                sb.Append("</border>");
            }
            return sb.ToString();
        }

        /// <summary>
        /// Method to create the XML string for the font part of the style sheet document
        /// </summary>
        /// <returns>String with formatted XML data</returns>
        private string CreateStyleFontString()
        {
            Font[] fontStyles = styles.GetFonts();
            StringBuilder sb = new StringBuilder();
            foreach (Font item in fontStyles)
            {
                sb.Append("<font>");
                if (item.Bold) { sb.Append("<b/>"); }
                if (item.Italic) { sb.Append("<i/>"); }
                if (item.Strike) { sb.Append("<strike/>"); }
                if (item.Underline != Font.UnderlineValue.none)
                {
                    if (item.Underline == Font.UnderlineValue.u_double) { sb.Append("<u val=\"double\"/>"); }
                    else if (item.Underline == Font.UnderlineValue.singleAccounting) { sb.Append("<u val=\"singleAccounting\"/>"); }
                    else if (item.Underline == Font.UnderlineValue.doubleAccounting) { sb.Append("<u val=\"doubleAccounting\"/>"); }
                    else { sb.Append("<u/>"); }
                }
                if (item.VerticalAlign == Font.VerticalAlignValue.subscript) { sb.Append("<vertAlign val=\"subscript\"/>"); }
                else if (item.VerticalAlign == Font.VerticalAlignValue.superscript) { sb.Append("<vertAlign val=\"superscript\"/>"); }
                sb.Append("<sz val=\"").Append(item.Size.ToString("G", culture)).Append("\"/>");
                if (string.IsNullOrEmpty(item.ColorValue))
                {
                    sb.Append("<color theme=\"").Append(item.ColorTheme.ToString("G", culture)).Append("\"/>");
                }
                else
                {
                    sb.Append("<color rgb=\"").Append(item.ColorValue).Append("\"/>");
                }
                sb.Append("<name val=\"").Append(item.Name).Append("\"/>");
                sb.Append("<family val=\"").Append(item.Family).Append("\"/>");
                if (item.Scheme != Font.SchemeValue.none)
                {
                    if (item.Scheme == Font.SchemeValue.major)
                    { sb.Append("<scheme val=\"major\"/>"); }
                    else if (item.Scheme == Font.SchemeValue.minor)
                    { sb.Append("<scheme val=\"minor\"/>"); }
                }
                if (!string.IsNullOrEmpty(item.Charset))
                {
                    sb.Append("<charset val=\"").Append(item.Charset).Append("\"/>");
                }
                sb.Append("</font>");
            }
            return sb.ToString();
        }

        /// <summary>
        /// Method to create the XML string for the fill part of the style sheet document
        /// </summary>
        /// <returns>String with formatted XML data</returns>
        private string CreateStyleFillString()
        {
            Fill[] fillStyles = styles.GetFills();
            StringBuilder sb = new StringBuilder();
            foreach (Fill item in fillStyles)
            {
                sb.Append("<fill>");
                sb.Append("<patternFill patternType=\"").Append(Fill.GetPatternName(item.PatternFill)).Append("\"");
                if (item.PatternFill == Fill.PatternValue.solid)
                {
                    sb.Append(">");
                    sb.Append("<fgColor rgb=\"").Append(item.ForegroundColor).Append("\"/>");
                    sb.Append("<bgColor indexed=\"").Append(item.IndexedColor.ToString("G", culture)).Append("\"/>");
                    sb.Append("</patternFill>");
                }
                else if (item.PatternFill == Fill.PatternValue.mediumGray || item.PatternFill == Fill.PatternValue.lightGray || item.PatternFill == Fill.PatternValue.gray0625 || item.PatternFill == Fill.PatternValue.darkGray)
                {
                    sb.Append(">");
                    sb.Append("<fgColor rgb=\"").Append(item.ForegroundColor).Append("\"/>");
                    if (!string.IsNullOrEmpty(item.BackgroundColor))
                    {
                        sb.Append("<bgColor rgb=\"").Append(item.BackgroundColor).Append("\"/>");
                    }
                    sb.Append("</patternFill>");
                }
                else
                {
                    sb.Append("/>");
                }
                sb.Append("</fill>");
            }
            return sb.ToString();
        }

        /// <summary>
        /// Method to create the XML string for the number format part of the style sheet document 
        /// </summary>
        /// <returns>String with formatted XML data</returns>
        private string CreateStyleNumberFormatString()
        {
            NumberFormat[] numberFormatStyles = styles.GetNumberFormats();
            StringBuilder sb = new StringBuilder();
            foreach (NumberFormat item in numberFormatStyles)
            {
                if (item.IsCustomFormat)
                {
                    if (string.IsNullOrEmpty(item.CustomFormatCode))
                    {
                        throw new FormatException("The number format style component with the ID " + Utils.ToString(item.CustomFormatID) + " cannot be null or empty");
                    }
                    // OOXML: Escaping according to Chp.18.8.31
                    // TODO: v3> Add a custom format builder
                    sb.Append("<numFmt formatCode=\"").Append(EscapeXmlAttributeChars(item.CustomFormatCode)).Append("\" numFmtId=\"").Append(item.CustomFormatID.ToString("G", culture)).Append("\"/>");
                }
            }
            return sb.ToString();
        }

        /// <summary>
        /// Method to create the XML string for the XF part of the style sheet document
        /// </summary>
        /// <returns>String with formatted XML data</returns>
        private string CreateStyleXfsString()
        {
            Style[] styleItems = this.styles.GetStyles();
            StringBuilder sb = new StringBuilder();
            StringBuilder sb2 = new StringBuilder();
            string alignmentString, protectionString;
            int formatNumber, textRotation;
            foreach (Style style in styleItems)
            {
                textRotation = style.CurrentCellXf.CalculateInternalRotation();
                alignmentString = string.Empty;
                protectionString = string.Empty;
                if (style.CurrentCellXf.HorizontalAlign != CellXf.HorizontalAlignValue.none || style.CurrentCellXf.VerticalAlign != CellXf.VerticalAlignValue.none || style.CurrentCellXf.Alignment != CellXf.TextBreakValue.none || textRotation != 0)
                {
                    sb2.Clear();
                    sb2.Append("<alignment");
                    if (style.CurrentCellXf.HorizontalAlign != CellXf.HorizontalAlignValue.none)
                    {
                        sb2.Append(" horizontal=\"");
                        if (style.CurrentCellXf.HorizontalAlign == CellXf.HorizontalAlignValue.center) { sb2.Append("center"); }
                        else if (style.CurrentCellXf.HorizontalAlign == CellXf.HorizontalAlignValue.right) { sb2.Append("right"); }
                        else if (style.CurrentCellXf.HorizontalAlign == CellXf.HorizontalAlignValue.centerContinuous) { sb2.Append("centerContinuous"); }
                        else if (style.CurrentCellXf.HorizontalAlign == CellXf.HorizontalAlignValue.distributed) { sb2.Append("distributed"); }
                        else if (style.CurrentCellXf.HorizontalAlign == CellXf.HorizontalAlignValue.fill) { sb2.Append("fill"); }
                        else if (style.CurrentCellXf.HorizontalAlign == CellXf.HorizontalAlignValue.general) { sb2.Append("general"); }
                        else if (style.CurrentCellXf.HorizontalAlign == CellXf.HorizontalAlignValue.justify) { sb2.Append("justify"); }
                        else { sb2.Append("left"); }
                        sb2.Append("\"");
                    }
                    if (style.CurrentCellXf.VerticalAlign != CellXf.VerticalAlignValue.none)
                    {
                        sb2.Append(" vertical=\"");
                        if (style.CurrentCellXf.VerticalAlign == CellXf.VerticalAlignValue.center) { sb2.Append("center"); }
                        else if (style.CurrentCellXf.VerticalAlign == CellXf.VerticalAlignValue.distributed) { sb2.Append("distributed"); }
                        else if (style.CurrentCellXf.VerticalAlign == CellXf.VerticalAlignValue.justify) { sb2.Append("justify"); }
                        else if (style.CurrentCellXf.VerticalAlign == CellXf.VerticalAlignValue.top) { sb2.Append("top"); }
                        else { sb2.Append("bottom"); }
                        sb2.Append("\"");
                    }
                    if (style.CurrentCellXf.Indent > 0 &&
                        (style.CurrentCellXf.HorizontalAlign == CellXf.HorizontalAlignValue.left
                        || style.CurrentCellXf.HorizontalAlign == CellXf.HorizontalAlignValue.right
                        || style.CurrentCellXf.HorizontalAlign == CellXf.HorizontalAlignValue.distributed))
                    {
                        sb2.Append(" indent=\"");
                        sb2.Append(style.CurrentCellXf.Indent.ToString("G", culture));
                        sb2.Append("\"");
                    }
                    if (style.CurrentCellXf.Alignment != CellXf.TextBreakValue.none)
                    {
                        if (style.CurrentCellXf.Alignment == CellXf.TextBreakValue.shrinkToFit) { sb2.Append(" shrinkToFit=\"1"); }
                        else { sb2.Append(" wrapText=\"1"); }
                        sb2.Append("\"");
                    }
                    if (textRotation != 0)
                    {
                        sb2.Append(" textRotation=\"");
                        sb2.Append(textRotation.ToString("G", culture));
                        sb2.Append("\"");
                    }
                    sb2.Append("/>"); // </xf>
                    alignmentString = sb2.ToString();
                }

                if (style.CurrentCellXf.Hidden || style.CurrentCellXf.Locked)
                {
                    if (style.CurrentCellXf.Hidden && style.CurrentCellXf.Locked)
                    {
                        protectionString = "<protection locked=\"1\" hidden=\"1\"/>";
                    }
                    else if (style.CurrentCellXf.Hidden && !style.CurrentCellXf.Locked)
                    {
                        protectionString = "<protection hidden=\"1\" locked=\"0\"/>";
                    }
                    else
                    {
                        protectionString = "<protection hidden=\"0\" locked=\"1\"/>";
                    }
                }

                sb.Append("<xf numFmtId=\"");
                if (style.CurrentNumberFormat.IsCustomFormat)
                {
                    sb.Append(style.CurrentNumberFormat.CustomFormatID.ToString("G", culture));
                }
                else
                {
                    formatNumber = (int)style.CurrentNumberFormat.Number;
                    sb.Append(formatNumber.ToString("G", culture));
                }

                sb.Append("\" borderId=\"").Append(style.CurrentBorder.InternalID.Value.ToString("G", culture));
                sb.Append("\" fillId=\"").Append(style.CurrentFill.InternalID.Value.ToString("G", culture));
                sb.Append("\" fontId=\"").Append(style.CurrentFont.InternalID.Value.ToString("G", culture));
                if (!style.CurrentFont.IsDefaultFont)
                {
                    sb.Append("\" applyFont=\"1");
                }
                if (style.CurrentFill.PatternFill != Fill.PatternValue.none)
                {
                    sb.Append("\" applyFill=\"1");
                }
                if (!style.CurrentBorder.IsEmpty())
                {
                    sb.Append("\" applyBorder=\"1");
                }
                if (alignmentString != string.Empty || style.CurrentCellXf.ForceApplyAlignment)
                {
                    sb.Append("\" applyAlignment=\"1");
                }
                if (protectionString != string.Empty)
                {
                    sb.Append("\" applyProtection=\"1");
                }
                if (style.CurrentNumberFormat.Number != NumberFormat.FormatNumber.none)
                {
                    sb.Append("\" applyNumberFormat=\"1\"");
                }
                else
                {
                    sb.Append("\"");
                }
                if (alignmentString != string.Empty || protectionString != string.Empty)
                {
                    sb.Append(">");
                    sb.Append(alignmentString);
                    sb.Append(protectionString);
                    sb.Append("</xf>");
                }
                else
                {
                    sb.Append("/>");
                }
            }
            return sb.ToString();
        }

        /// <summary>
        /// Method to create the XML string for the color-MRU part of the style sheet document (recent colors)
        /// </summary>
        /// <returns>String with formatted XML data</returns>
        private string CreateMruColorsString()
        {
            StringBuilder sb = new StringBuilder();
            List<string> tempColors = new List<string>();
            foreach (string item in this.workbook.GetMruColors())
            {
                if (item == Fill.DEFAULT_COLOR)
                { 
                    continue; 
                }
                if (!tempColors.Contains(item)) { tempColors.Add(item); }
            }
            if (tempColors.Count > 0)
            {
                sb.Append("<mruColors>");
                foreach (string item in tempColors)
                {
                    sb.Append("<color rgb=\"").Append(item).Append("\"/>");
                }
                sb.Append("</mruColors>");
                return sb.ToString();
            }
            return string.Empty;
        }

        /// <summary>
        /// Method to sort the cells of a worksheet as preparation for the XML document
        /// </summary>
        /// <param name="sheet">Worksheet to process</param>
        /// <returns>Sorted list of dynamic rows that are either defined by cells or row widths / hidden states. The list is sorted by row numbers (zero-based)</returns>
        private List<DynamicRow> GetSortedSheetData(Worksheet sheet)
        {
            List<Cell> temp = new List<Cell>();
            foreach (KeyValuePair<string, Cell> item in sheet.Cells)
            {
                temp.Add(item.Value);
            }
            temp.Sort();
            DynamicRow row = new DynamicRow(); ;
            Dictionary<int, DynamicRow> rows = new Dictionary<int, DynamicRow>();
            int rowNumber;
            if (temp.Count > 0)
            {
                rowNumber = temp[0].RowNumber;
                row.RowNumber = rowNumber;
                foreach (Cell cell in temp)
                {
                    if (cell.RowNumber != rowNumber)
                    {
                        rows.Add(rowNumber, row);
                        row = new DynamicRow();
                        row.RowNumber = cell.RowNumber;
                        rowNumber = cell.RowNumber;
                    }
                    row.CellDefinitions.Add(cell);
                }
                if (row.CellDefinitions.Count > 0)
                {
                    rows.Add(rowNumber, row);
                }
            }
            foreach (KeyValuePair<int, float> rowHeight in sheet.RowHeights)
            {
                if (!rows.ContainsKey(rowHeight.Key))
                {
                    row = new DynamicRow();
                    row.RowNumber = rowHeight.Key;
                    rows.Add(rowHeight.Key, row);
                }
            }
            foreach(KeyValuePair<int, bool> hiddenRow in sheet.HiddenRows)
            {
               if (!rows.ContainsKey(hiddenRow.Key))
               {
                    row = new DynamicRow();
                    row.RowNumber = hiddenRow.Key;
                    rows.Add(hiddenRow.Key, row);
                }
            }
            List<DynamicRow> output = rows.Values.ToList();
            output.Sort((r1, r2) => (r1.RowNumber.CompareTo(r2.RowNumber))); // Lambda sort
            return output;
        }



        #endregion

        #region staticMethods

        /// <summary>
        /// Method to escape XML characters between two XML tags
        /// </summary>
        /// <param name="input">Input string to process</param>
        /// <returns>Escaped string</returns>
        /// <remarks>Note: The XML specs allow characters up to the character value of 0x10FFFF. However, the C# char range is only up to 0xFFFF. NanoXLSX will neglect all values above this level in the sanitizing check. Illegal characters like 0x1 will be replaced with a white space (0x20)</remarks>
        public static string EscapeXmlChars(string input)
        {
            if (input == null) { return ""; }
            int len = input.Length;
            List<int> illegalCharacters = new List<int>(len);
            List<byte> characterTypes = new List<byte>(len);
            int i;
            for (i = 0; i < len; i++)
            {
                if ((input[i] < 0x9) || (input[i] > 0xA && input[i] < 0xD) || (input[i] > 0xD && input[i] < 0x20) || (input[i] > 0xD7FF && input[i] < 0xE000) || (input[i] > 0xFFFD))
                {
                    illegalCharacters.Add(i);
                    characterTypes.Add(0);
                    continue;
                } // Note: XML specs allow characters up to 0x10FFFF. However, the C# char range is only up to 0xFFFF; Higher values are neglected here 
                if (input[i] == 0x3C) // <
                {
                    illegalCharacters.Add(i);
                    characterTypes.Add(1);
                }
                else if (input[i] == 0x3E) // >
                {
                    illegalCharacters.Add(i);
                    characterTypes.Add(2);
                }
                else if (input[i] == 0x26) // &
                {
                    illegalCharacters.Add(i);
                    characterTypes.Add(3);
                }
            }
            if (illegalCharacters.Count == 0)
            {
                return input;
            }

            StringBuilder sb = new StringBuilder(len);
            int lastIndex = 0;
            len = illegalCharacters.Count;
            for (i = 0; i < len; i++)
            {
                sb.Append(input.Substring(lastIndex, illegalCharacters[i] - lastIndex));
                if (characterTypes[i] == 0)
                {
                    sb.Append(' '); // Whitespace as fall back on illegal character
                }
                else if (characterTypes[i] == 1) // replace <
                {
                    sb.Append("&lt;");
                }
                else if (characterTypes[i] == 2) // replace >
                {
                    sb.Append("&gt;");
                }
                else if (characterTypes[i] == 3) // replace &
                {
                    sb.Append("&amp;");
                }
                lastIndex = illegalCharacters[i] + 1;
            }
            sb.Append(input.Substring(lastIndex));
            return sb.ToString();
        }

        /// <summary>
        /// Method to escape XML characters in an XML attribute
        /// </summary>
        /// <param name="input">Input string to process</param>
        /// <returns>Escaped string</returns>
        public static string EscapeXmlAttributeChars(string input)
        {
            input = EscapeXmlChars(input); // Sanitize string from illegal characters beside quotes
            input = input.Replace("\"", "&quot;");
            return input;
        }
        #endregion

        #region helperClasses
        /// <summary>
        /// Class representing a row that is either empty or containing cells. Empty rows can also carry information about height or visibility
        /// </summary>
        private sealed class DynamicRow
        {
            private readonly List<Cell> cellDefinitions;
            public int RowNumber { get; set; }

            /// <summary>
            /// Gets the List of cells if not empty
            /// </summary>
            public List<Cell> CellDefinitions
            {
                get { return cellDefinitions; }
            }

            /// <summary>
            /// Default constructor. Defines an empty row if no additional operations are made on the object
            /// </summary>
            public DynamicRow()
            {
                this.cellDefinitions = new List<Cell>();
            }
        }
        #endregion

    }
    #region doc
    /// <summary>
    /// Sub-namespace with all low-level classes and functions. This namespace is necessary to read and generate files but should not be used as pat of the API. Use the classes and functions in the namespace NanoXLSX instead
    /// </summary>
    [CompilerGenerated]
    class NamespaceDoc // This class is only for documentation purpose (Sandcastle)
    { }
    #endregion

}
