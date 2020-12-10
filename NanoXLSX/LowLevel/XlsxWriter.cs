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
using System.IO.Packaging;
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
        private SortedMap sharedStrings;
        private int sharedStringsTotalCount;
        private Dictionary<string, XmlDocument> interceptedDocuments;
        private bool interceptDocuments;
        #endregion

        #region properties
        /// <summary>
        /// Gets or set whether XML documents are intercepted during creation
        /// </summary>
        public bool InterceptDocuments
        {
            get { return interceptDocuments; }
            set
            {
                interceptDocuments = value;
                if (interceptDocuments && interceptedDocuments == null)
                {
                    interceptedDocuments = new Dictionary<string, XmlDocument>();
                }
                else if (!interceptDocuments)
                {
                    interceptedDocuments = null;
                }
            }
        }

        /// <summary>
        /// Gets the intercepted documents if interceptDocuments is set to true
        /// </summary>
        public Dictionary<string, XmlDocument> InterceptedDocuments
        {
            get { return interceptedDocuments; }
        }

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
                sb.Append("<si><t>");
                sb.Append(EscapeXmlChars(str));
                sb.Append("</t></si>");
            }
            sb.Append("</sst>");
            return sb.ToString();
        }

        /// <summary>
        /// Method to create a style sheet as raw XML string
        /// </summary>
        /// <returns>Raw XML string</returns>
        /// <exception cref="StyleException">Throws an StyleException if one of the styles cannot be referenced or is null</exception>
        /// <remarks>The UndefinedStyleException should never happen in this state if the internally managed style collection was not tampered. </remarks>
        private string CreateStyleSheetDocument()
        {
            string bordersString = CreateStyleBorderString();
            string fillsString = CreateStyleFillString();
            string fontsString = CreateStyleFontString();
            string numberFormatsString = CreateStyleNumberFormatString();
            string xfsStings = CreateStyleXfsString();
            string mruColorString = CreateMruColorsString();
            int fontCount = workbook.Styles.GetFontStyleNumber();
            int fillCount = workbook.Styles.GetFillStyleNumber();
            int styleCount = workbook.Styles.GetStyleNumber();
            int borderCount = workbook.Styles.GetBorderStyleNumber();
            StringBuilder sb = new StringBuilder();
            sb.Append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">");
            int numFormatCount = workbook.Styles.GetNumberFormatStyleNumber();
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
                if (!string.IsNullOrEmpty(mruColorString) && workbook.WorkbookMetadata.UseColorMRU)
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
        /// <exception cref="RangeException">Throws an OutOfRangeException if an address was out of range</exception>
        private string CreateWorkbookDocument()
        {
            if (workbook.Worksheets.Count == 0)
            {
                throw new RangeException(RangeException.GENERAL, "The workbook can not be created because no worksheet was defined.");
            }
            StringBuilder sb = new StringBuilder();
            sb.Append("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
            if (workbook.SelectedWorksheet > 0)
            {
                sb.Append("<bookViews><workbookView activeTab=\"");
                sb.Append(workbook.SelectedWorksheet.ToString("G", culture));
                sb.Append("\"/></bookViews>");
            }
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
                    sb.Append("workbookPassword=\"");
                    sb.Append(GeneratePasswordHash(workbook.WorkbookProtectionPassword));
                    sb.Append("\"");
                }
                sb.Append("/>");
            }
            sb.Append("<sheets>");
            foreach (Worksheet item in workbook.Worksheets)
            {
                sb.Append("<sheet r:id=\"rId").Append(item.SheetID.ToString()).Append("\" sheetId=\"").Append(item.SheetID.ToString()).Append("\" name=\"").Append(EscapeXmlAttributeChars(item.SheetName)).Append("\"/>");
            }
            sb.Append("</sheets>");
            sb.Append("</workbook>");
            return sb.ToString();
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
            List<List<Cell>> celldata = GetSortedSheetData(worksheet);
            StringBuilder sb = new StringBuilder();
            string line;
            sb.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">");

            if (worksheet.SelectedCells != null)
            {
                sb.Append("<sheetViews><sheetView workbookViewId=\"0\"");
                if (workbook.SelectedWorksheet == worksheet.SheetID - 1)
                {
                    sb.Append(" tabSelected=\"1\"");
                }
                sb.Append("><selection sqref=\"");
                sb.Append(worksheet.SelectedCells.ToString());
                sb.Append("\" activeCell=\"");
                sb.Append(worksheet.SelectedCells.Value.StartAddress.ToString());
                sb.Append("\"/></sheetView></sheetViews>");
            }

            sb.Append("<sheetFormatPr x14ac:dyDescent=\"0.25\" defaultRowHeight=\"").Append(worksheet.DefaultRowHeight.ToString("G", culture)).Append("\" baseColWidth=\"").Append(worksheet.DefaultColumnWidth.ToString("G", culture)).Append("\"/>");

            string colWidths = CreateColsString(worksheet);
            if (!string.IsNullOrEmpty(colWidths))
            {
                sb.Append("<cols>");
                sb.Append(colWidths);
                sb.Append("</cols>");
            }
            sb.Append("<sheetData>");
            foreach (List<Cell> item in celldata)
            {
                line = CreateRowString(item, worksheet);
                sb.Append(line);
            }
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
        /// Method to save the workbook
        /// </summary>
        /// <exception cref="Exceptions.IOException">Throws IOException in case of an error</exception>
        /// <exception cref="RangeException">Throws an OutOfRangeException if the start or end address of a handled cell range was out of range</exception>
        /// <exception cref="Exceptions.FormatException">Throws a FormatException if a handled date cannot be translated to (Excel internal) OADate</exception>
        /// <exception cref="StyleException">Throws an StyleException if one of the styles of the workbook cannot be referenced or is null</exception>
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
                throw new IOException("SaveException", "An error occurred while saving. See inner exception for details: " + e.Message, e);
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
        /// <exception cref="RangeException">Throws an OutOfRangeException if the start or end address of a handled cell range was out of range</exception>
        /// <exception cref="FormatException">Throws a FormatException if a handled date cannot be translated to (Excel internal) OADate</exception>
        /// <exception cref="StyleException">Throws an StyleException if one of the styles of the workbook cannot be referenced or is null</exception>
        /// <remarks>The StyleException should never happen in this state if the internally managed style collection was not tampered. </remarks>
        public void SaveAsStream(Stream stream, bool leaveOpen = false)
        {
            workbook.ResolveMergedCells();
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

                    AppendXmlToPackagePart(CreateWorkbookDocument(), pp, "WORKBOOK");
                    int idCounter = workbook.Worksheets.Count + 1;

                    pp.CreateRelationship(stylesheetUri, TargetMode.Internal, @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "rId" + idCounter);
                    pp.CreateRelationship(sharedStringsUri, TargetMode.Internal, @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings", "rId" + (idCounter + 1));

                    foreach (Worksheet item in workbook.Worksheets)
                    {
                        sheetPath = new DocumentPath("sheet" + item.SheetID + ".xml", "xl/worksheets");
                        sheetURIs.Add(new Uri(sheetPath.GetFullPath(), UriKind.Relative));
                        pp.CreateRelationship(sheetURIs[sheetURIs.Count - 1], TargetMode.Internal, @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", "rId" + item.SheetID);
                    }

                    pp = p.CreatePart(stylesheetUri, @"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml", CompressionOption.Normal);
                    AppendXmlToPackagePart(CreateStyleSheetDocument(), pp, "STYLESHEET");

                    int i = 0;
                    foreach (Worksheet item in workbook.Worksheets)
                    {
                        pp = p.CreatePart(sheetURIs[i], @"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", CompressionOption.Normal);
                        i++;
                        AppendXmlToPackagePart(CreateWorksheetPart(item), pp, "WORKSHEET:" + item.SheetName);
                    }
                    pp = p.CreatePart(sharedStringsUri, @"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml", CompressionOption.Normal);
                    AppendXmlToPackagePart(CreateSharedStringsDocument(), pp, "SHAREDSTRINGS");

                    if (workbook.WorkbookMetadata != null)
                    {
                        pp = p.CreatePart(appPropertiesUri, @"application/vnd.openxmlformats-officedocument.extended-properties+xml", CompressionOption.Normal);
                        AppendXmlToPackagePart(CreateAppPropertiesDocument(), pp, "APPPROPERTIES");
                        pp = p.CreatePart(corePropertiesUri, @"application/vnd.openxmlformats-package.core-properties+xml", CompressionOption.Normal);
                        AppendXmlToPackagePart(CreateCorePropertiesDocument(), pp, "COREPROPERTIES");
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
                throw new IOException("SaveException", "An error occurred while saving. See inner exception for details: " + e.Message, e);
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
            if (sb == null || string.IsNullOrEmpty(tagName)) { return; }
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
        /// <param name="title">Title for interception / debugging purpose</param>
        /// <exception cref="Exceptions.IOException">Throws an IOException if the XML data could not be written into the Package Part</exception>
        private void AppendXmlToPackagePart(string doc, PackagePart pp, string title)
        {
            try
            {
                if (interceptDocuments)
                {
                    XmlDocument xDoc = new XmlDocument();
                    xDoc.LoadXml(doc);
                    interceptedDocuments.Add(title, xDoc);
                }
                using (MemoryStream ms = new MemoryStream()) // Write workbook.xml
                {
                    if (!ms.CanWrite) { return; }
                    using (XmlWriter writer = XmlWriter.Create(ms))
                    {
                        //doc.WriteTo(writer);
                        writer.WriteProcessingInstruction("xml", "version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"");
                        writer.WriteRaw(doc);
                        writer.Flush();
                        ms.Position = 0;
                        ms.CopyTo(pp.GetStream());
                        ms.Flush();
                    }
                }
            }
            catch (Exception e)
            {
                throw new IOException("MemoryStreamException", "The XML document could not be saved into the memory stream", e);
            }
        }

        /// <summary>
        /// Method to create the XML string for the app-properties document
        /// </summary>
        /// <returns>String with formatted XML data</returns>
        private string CreateAppString()
        {
            if (workbook.WorkbookMetadata == null) { return string.Empty; }
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
                    if (column.Value.Width == worksheet.DefaultColumnWidth && !column.Value.IsHidden) { continue; }
                    if (worksheet.Columns.ContainsKey(column.Key))
                    {
                        if (worksheet.Columns[column.Key].IsHidden)
                        {
                            hidden = " hidden=\"1\"";
                        }
                    }
                    col = (column.Key + 1).ToString("G", culture); // Add 1 for Address
                    sb.Append("<col customWidth=\"1\" width=\"").Append(column.Value.Width.ToString("G", culture)).Append("\" max=\"").Append(col).Append("\" min=\"").Append(col).Append("\"").Append(hidden).Append("/>");
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
            if (workbook.WorkbookMetadata == null) { return string.Empty; }
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
        /// <param name="columnFields">List of cells</param>
        /// <param name="worksheet">Worksheet to process</param>
        /// <returns>Formatted row string</returns>
        /// <exception cref="Exceptions.FormatException">Throws a FormatException if a handled date cannot be translated to (Excel internal) OADate</exception>
        private string CreateRowString(List<Cell> columnFields, Worksheet worksheet)
        {
            int rowNumber = columnFields[0].RowNumber;
            string height = "";
            string hidden = "";
            if (worksheet.RowHeights.ContainsKey(rowNumber))
            {
                if (worksheet.RowHeights[rowNumber] != worksheet.DefaultRowHeight)
                {
                    height = " x14ac:dyDescent=\"0.25\" customHeight=\"1\" ht=\"" + worksheet.RowHeights[rowNumber].ToString("G", culture) + "\"";
                }
            }
            if (worksheet.HiddenRows.ContainsKey(rowNumber))
            {
                if (worksheet.HiddenRows[rowNumber])
                {
                    hidden = " hidden=\"1\"";
                }
            }
            StringBuilder sb = new StringBuilder();
            if (columnFields.Count > 0)
            {
                sb.Append("<row r=\"").Append((rowNumber + 1).ToString()).Append("\"").Append(height).Append(hidden).Append(">");
            }
            else
            {
                sb.Append("<row").Append(height).Append(">");
            }
            string typeAttribute;
            string sValue = "";
            string tValue = "";
            string value = "";
            bool bVal;

            int col = 0;
            foreach (Cell item in columnFields)
            {
                tValue = " ";
                if (item.CellStyle != null)
                {
                    sValue = " s=\"" + item.CellStyle.InternalID.Value.ToString("G", culture) + "\" ";
                }
                else
                {
                    sValue = "";
                }
                item.ResolveCellType(); // Recalculate the type (for handling DEFAULT)
                if (item.DataType == Cell.CellType.BOOL)
                {
                    typeAttribute = "b";
                    tValue = " t=\"" + typeAttribute + "\" ";
                    bVal = (bool)item.Value;
                    if (bVal) { value = "1"; }
                    else { value = "0"; }

                }
                // Number casting
                else if (item.DataType == Cell.CellType.NUMBER)
                {
                    typeAttribute = "n";
                    tValue = " t=\"" + typeAttribute + "\" ";
                    Type t = item.Value.GetType();

                    if (t == typeof(byte)) { value = ((byte)item.Value).ToString("G", culture); }
                    else if (t == typeof(sbyte)) { value = ((sbyte)item.Value).ToString("G", culture); }
                    else if (t == typeof(decimal)) { value = ((decimal)item.Value).ToString("G", culture); }
                    else if (t == typeof(double)) { value = ((double)item.Value).ToString("G", culture); }
                    else if (t == typeof(float)) { value = ((float)item.Value).ToString("G", culture); }
                    else if (t == typeof(int)) { value = ((int)item.Value).ToString("G", culture); }
                    else if (t == typeof(uint)) { value = ((uint)item.Value).ToString("G", culture); }
                    else if (t == typeof(long)) { value = ((long)item.Value).ToString("G", culture); }
                    else if (t == typeof(ulong)) { value = ((ulong)item.Value).ToString("G", culture); }
                    else if (t == typeof(short)) { value = ((short)item.Value).ToString("G", culture); }
                    else if (t == typeof(ushort)) { value = ((ushort)item.Value).ToString("G", culture); }
                }
                // Date parsing
                else if (item.DataType == Cell.CellType.DATE)
                {
                    typeAttribute = "d";
                    DateTime date = (DateTime)item.Value;
                    value = Utils.GetOADateTimeString(date, culture);
                }
                // Time parsing
                else if (item.DataType == Cell.CellType.TIME)
                {
                    typeAttribute = "d";
                    // TODO: 'd' is probably an outdated attribute (to be checked for dates and times)
                    TimeSpan time = (TimeSpan)item.Value;
                    value = Utils.GetOATimeString(time, culture);
                }
                else
                {
                    if (item.Value == null)
                    {
                        typeAttribute = "str";
                        value = string.Empty;
                    }
                    else // Handle sharedStrings
                    {
                        if (item.DataType == Cell.CellType.FORMULA)
                        {
                            typeAttribute = "str";
                            value = item.Value.ToString();
                        }
                        else
                        {
                            typeAttribute = "s";
                            value = item.Value.ToString();
                            if (!sharedStrings.ContainsKey(value))
                            {
                                sharedStrings.Add(value, sharedStrings.Count.ToString("G", culture));
                            }
                            value = sharedStrings[value];
                            sharedStringsTotalCount++;
                        }
                    }
                    tValue = " t=\"" + typeAttribute + "\" ";
                }
                if (item.DataType != Cell.CellType.EMPTY)
                {
                    sb.Append("<c").Append(tValue).Append("r=\"").Append(item.CellAddress).Append("\"").Append(sValue).Append(">");
                    if (item.DataType == Cell.CellType.FORMULA)
                    {
                        sb.Append("<f>").Append(EscapeXmlChars(item.Value.ToString())).Append("</f>");
                    }
                    else
                    {
                        sb.Append("<v>").Append(EscapeXmlChars(value)).Append("</v>");
                    }
                    sb.Append("</c>");
                }
                else // Empty cell
                {
                    sb.Append("<c").Append(tValue).Append("r=\"").Append(item.CellAddress).Append("\"").Append(sValue).Append("/>");
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
                try
                {
                    temp = Enum.GetName(typeof(Worksheet.SheetProtectionValue), item.Key); // Note! If the enum names differs from the OOXML definitions, this method will cause invalid OOXML entries
                    sb.Append(" ").Append(temp).Append("=\"").Append(item.Value.ToString("G", culture)).Append("\"");
                }
                catch { }
            }
            if (!string.IsNullOrEmpty(sheet.SheetProtectionPassword))
            {
                string hash = GeneratePasswordHash(sheet.SheetProtectionPassword);
                sb.Append(" password=\"").Append(hash).Append("\"");
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
            Border[] borderStyles = workbook.Styles.GetBorders();
            StringBuilder sb = new StringBuilder();
            foreach (Border item in borderStyles)
            {
                if (!item.DiagonalDown && item.DiagonalUp) { sb.Append("<border diagonalDown=\"1\">"); }
                else if (!item.DiagonalDown && item.DiagonalUp) { sb.Append("<border diagonalUp=\"1\">"); }
                else if (item.DiagonalDown && item.DiagonalUp) { sb.Append("<border diagonalDown=\"1\" diagonalUp=\"1\">"); }
                else { sb.Append("<border>"); }

                if (item.LeftStyle != Border.StyleValue.none)
                {
                    sb.Append("<left style=\"" + Border.GetStyleName(item.LeftStyle) + "\">");
                    if (string.IsNullOrEmpty(item.LeftColor)) { sb.Append("<color rgb=\"").Append(item.LeftColor).Append("\"/>"); }
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
                    if (string.IsNullOrEmpty(item.RightColor)) { sb.Append("<color rgb=\"").Append(item.RightColor).Append("\"/>"); }
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
                    if (string.IsNullOrEmpty(item.TopColor)) { sb.Append("<color rgb=\"").Append(item.TopColor).Append("\"/>"); }
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
                    if (string.IsNullOrEmpty(item.BottomColor)) { sb.Append("<color rgb=\"").Append(item.BottomColor).Append("\"/>"); }
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
                    if (string.IsNullOrEmpty(item.DiagonalColor)) { sb.Append("<color rgb=\"").Append(item.DiagonalColor).Append("\"/>"); }
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
            Font[] fontStyles = workbook.Styles.GetFonts();
            StringBuilder sb = new StringBuilder();
            foreach (Font item in fontStyles)
            {
                sb.Append("<font>");
                if (item.Bold) { sb.Append("<b/>"); }
                if (item.Italic) { sb.Append("<i/>"); }
                if (item.Underline) { sb.Append("<u/>"); }
                if (item.DoubleUnderline) { sb.Append("<u val=\"double\"/>"); }
                if (item.Strike) { sb.Append("<strike/>"); }
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
            Fill[] fillStyles = workbook.Styles.GetFills();
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
            NumberFormat[] numberFormatStyles = workbook.Styles.GetNumberFormats();
            StringBuilder sb = new StringBuilder();
            foreach (NumberFormat item in numberFormatStyles)
            {
                if (item.IsCustomFormat)
                {
                    sb.Append("<numFmt formatCode=\"").Append(item.CustomFormatCode).Append("\" numFmtId=\"").Append(item.CustomFormatID.ToString("G", culture)).Append("\"/>");
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
            Style[] styles = workbook.Styles.GetStyles();
            StringBuilder sb = new StringBuilder();
            StringBuilder sb2 = new StringBuilder();
            string alignmentString, protectionString;
            int formatNumber, textRotation;
            foreach (Style item in styles)
            {
                textRotation = item.CurrentCellXf.CalculateInternalRotation();
                alignmentString = string.Empty;
                protectionString = string.Empty;
                if (item.CurrentCellXf.HorizontalAlign != CellXf.HorizontalAlignValue.none || item.CurrentCellXf.VerticalAlign != CellXf.VerticalAlignValue.none || item.CurrentCellXf.Alignment != CellXf.TextBreakValue.none || textRotation != 0)
                {
                    sb2.Clear();
                    sb2.Append("<alignment");
                    if (item.CurrentCellXf.HorizontalAlign != CellXf.HorizontalAlignValue.none)
                    {
                        sb2.Append(" horizontal=\"");
                        if (item.CurrentCellXf.HorizontalAlign == CellXf.HorizontalAlignValue.center) { sb2.Append("center"); }
                        else if (item.CurrentCellXf.HorizontalAlign == CellXf.HorizontalAlignValue.right) { sb2.Append("right"); }
                        else if (item.CurrentCellXf.HorizontalAlign == CellXf.HorizontalAlignValue.centerContinuous) { sb2.Append("centerContinuous"); }
                        else if (item.CurrentCellXf.HorizontalAlign == CellXf.HorizontalAlignValue.distributed) { sb2.Append("distributed"); }
                        else if (item.CurrentCellXf.HorizontalAlign == CellXf.HorizontalAlignValue.fill) { sb2.Append("fill"); }
                        else if (item.CurrentCellXf.HorizontalAlign == CellXf.HorizontalAlignValue.general) { sb2.Append("general"); }
                        else if (item.CurrentCellXf.HorizontalAlign == CellXf.HorizontalAlignValue.justify) { sb2.Append("justify"); }
                        else { sb2.Append("left"); }
                        sb2.Append("\"");
                    }
                    if (item.CurrentCellXf.VerticalAlign != CellXf.VerticalAlignValue.none)
                    {
                        sb2.Append(" vertical=\"");
                        if (item.CurrentCellXf.VerticalAlign == CellXf.VerticalAlignValue.center) { sb2.Append("center"); }
                        else if (item.CurrentCellXf.VerticalAlign == CellXf.VerticalAlignValue.distributed) { sb2.Append("distributed"); }
                        else if (item.CurrentCellXf.VerticalAlign == CellXf.VerticalAlignValue.justify) { sb2.Append("justify"); }
                        else if (item.CurrentCellXf.VerticalAlign == CellXf.VerticalAlignValue.top) { sb2.Append("top"); }
                        else { sb2.Append("bottom"); }
                        sb2.Append("\"");
                    }
                    if (item.CurrentCellXf.Indent > 0 &&
                        (item.CurrentCellXf.HorizontalAlign == CellXf.HorizontalAlignValue.left
                        || item.CurrentCellXf.HorizontalAlign == CellXf.HorizontalAlignValue.right
                        || item.CurrentCellXf.HorizontalAlign == CellXf.HorizontalAlignValue.distributed))
                    {
                        sb2.Append(" indent=\"");
                        sb2.Append(item.CurrentCellXf.Indent.ToString("G", culture));
                        sb2.Append("\"");
                    }
                    if (item.CurrentCellXf.Alignment != CellXf.TextBreakValue.none)
                    {
                        if (item.CurrentCellXf.Alignment == CellXf.TextBreakValue.shrinkToFit) { sb2.Append(" shrinkToFit=\"1"); }
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

                if (item.CurrentCellXf.Hidden || item.CurrentCellXf.Locked)
                {
                    if (item.CurrentCellXf.Hidden && item.CurrentCellXf.Locked)
                    {
                        protectionString = "<protection locked=\"1\" hidden=\"1\"/>";
                    }
                    else if (!item.CurrentCellXf.Hidden && item.CurrentCellXf.Locked)
                    {
                        protectionString = "<protection hidden=\"1\" locked=\"0\"/>";
                    }
                    else
                    {
                        protectionString = "<protection hidden=\"0\" locked=\"1\"/>";
                    }
                }

                sb.Append("<xf numFmtId=\"");
                if (item.CurrentNumberFormat.IsCustomFormat)
                {
                    sb.Append(item.CurrentNumberFormat.CustomFormatID.ToString("G", culture));
                }
                else
                {
                    formatNumber = (int)item.CurrentNumberFormat.Number;
                    sb.Append(formatNumber.ToString("G", culture));
                }

                sb.Append("\" borderId=\"").Append(item.CurrentBorder.InternalID.Value.ToString("G", culture));
                sb.Append("\" fillId=\"").Append(item.CurrentFill.InternalID.Value.ToString("G", culture));
                sb.Append("\" fontId=\"").Append(item.CurrentFont.InternalID.Value.ToString("G", culture));
                if (!item.CurrentFont.IsDefaultFont)
                {
                    sb.Append("\" applyFont=\"1");
                }
                if (item.CurrentFill.PatternFill != Fill.PatternValue.none)
                {
                    sb.Append("\" applyFill=\"1");
                }
                if (!item.CurrentBorder.IsEmpty())
                {
                    sb.Append("\" applyBorder=\"1");
                }
                if (alignmentString != string.Empty || item.CurrentCellXf.ForceApplyAlignment)
                {
                    sb.Append("\" applyAlignment=\"1");
                }
                if (protectionString != string.Empty)
                {
                    sb.Append("\" applyProtection=\"1");
                }
                if (item.CurrentNumberFormat.Number != NumberFormat.FormatNumber.none)
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
            Font[] fonts = workbook.Styles.GetFonts();
            Fill[] fills = workbook.Styles.GetFills();
            StringBuilder sb = new StringBuilder();
            List<string> tempColors = new List<string>();
            foreach (Font item in fonts)
            {
                if (string.IsNullOrEmpty(item.ColorValue)) { continue; }
                if (item.ColorValue == Fill.DEFAULTCOLOR) { continue; }
                if (!tempColors.Contains(item.ColorValue)) { tempColors.Add(item.ColorValue); }
            }
            foreach (Fill item in fills)
            {
                if (!string.IsNullOrEmpty(item.BackgroundColor))
                {
                    if (item.BackgroundColor != Fill.DEFAULTCOLOR)
                    {
                        if (!tempColors.Contains(item.BackgroundColor)) { tempColors.Add(item.BackgroundColor); }
                    }
                }
                if (!string.IsNullOrEmpty(item.ForegroundColor))
                {
                    if (item.ForegroundColor != Fill.DEFAULTCOLOR)
                    {
                        if (!tempColors.Contains(item.ForegroundColor)) { tempColors.Add(item.ForegroundColor); }
                    }
                }
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
        /// <returns>Two dimensional array of Cell objects</returns>
        private List<List<Cell>> GetSortedSheetData(Worksheet sheet)
        {
            List<Cell> temp = new List<Cell>();
            foreach (KeyValuePair<string, Cell> item in sheet.Cells)
            {
                temp.Add(item.Value);
            }
            temp.Sort();
            List<Cell> line = new List<Cell>();
            List<List<Cell>> output = new List<List<Cell>>();
            if (temp.Count > 0)
            {
                int rowNumber = temp[0].RowNumber;
                foreach (Cell item in temp)
                {
                    if (item.RowNumber != rowNumber)
                    {
                        output.Add(line);
                        line = new List<Cell>();
                        rowNumber = item.RowNumber;
                    }
                    line.Add(item);
                }
                if (line.Count > 0)
                {
                    output.Add(line);
                }
            }
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

        /// <summary>
        /// Method to generate an Excel internal password hash to protect workbooks or worksheets<br></br>This method is derived from the c++ implementation by Kohei Yoshida (<a href="http://kohei.us/2008/01/18/excel-sheet-protection-password-hash/">http://kohei.us/2008/01/18/excel-sheet-protection-password-hash/</a>)
        /// </summary>
        /// <remarks>WARNING! Do not use this method to encrypt 'real' passwords or data outside from NanoXLSX. This is only a minor security feature. Use a proper cryptography method instead.</remarks>
        /// <param name="password">Password string in UTF-8 to encrypt</param>
        /// <returns>16 bit hash as hex string</returns>
        public static string GeneratePasswordHash(string password)
        {
            if (string.IsNullOrEmpty(password)) { return string.Empty; }
            int passwordLength = password.Length;
            int passwordHash = 0;
            char character;
            for (int i = passwordLength; i > 0; i--)
            {
                character = password[i - 1];
                passwordHash = ((passwordHash >> 14) & 0x01) | ((passwordHash << 1) & 0x7fff);
                passwordHash ^= character;
            }
            passwordHash = ((passwordHash >> 14) & 0x01) | ((passwordHash << 1) & 0x7fff);
            passwordHash ^= (0x8000 | ('N' << 8) | 'K');
            passwordHash ^= passwordLength;
            return passwordHash.ToString("X");
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
