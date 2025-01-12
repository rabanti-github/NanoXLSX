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
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using NanoXLSX.Interfaces.Writer;
using NanoXLSX.Internal.Structures;
using NanoXLSX.Registry;
using NanoXLSX.Exceptions;
using NanoXLSX.Utils;
using NanoXLSX.Styles;
using IOException = NanoXLSX.Exceptions.IOException;

namespace NanoXLSX.Internal.Writers
{
    /// <summary>
    /// Class for internal handling (XML, formatting, packing)
    /// </summary>
    /// \remark <remarks>This class is only for internal use. Use the high level API (e.g. class Workbook) to manipulate data and create Excel files</remarks>
    internal class XlsxWriter
    {

        #region staticFields
        private static DocumentPath WORKBOOK = new DocumentPath("workbook.xml", "xl/");
        private static DocumentPath STYLES = new DocumentPath("styles.xml", "xl/");
        private static DocumentPath APP_PROPERTIES = new DocumentPath("app.xml", "docProps/");
        private static DocumentPath CORE_PROPERTIES = new DocumentPath("core.xml", "docProps/");
        private static DocumentPath SHARED_STRINGS = new DocumentPath("sharedStrings.xml", "xl/");
        private static DocumentPath THEME = new DocumentPath("theme1.xml", "xl/theme/");
        #endregion

        #region privateFields
        private Workbook workbook;
        private StyleManager styles;
        #endregion


        #region properties
        public Workbook Workbook
        {
            get { return workbook; }
        }

        public StyleManager Styles
        {
            get { return styles; }
        }

        #endregion

        #region constructors
        /// <summary>
        /// Constructor with defined workbook object
        /// </summary>
        /// <param name="workbook">Workbook to process</param>
        public XlsxWriter(Workbook workbook)
        {
            this.workbook = workbook;
        }
        #endregion

        #region documentCreation_methods

        /// <summary>
        /// Method to normalize all newlines to CR+LF
        /// </summary>
        /// <param name="value">Input value</param>
        /// <returns>Normalized value</returns>
        internal static string NormalizeNewLines(string value)
        {
            if (value == null || (!value.Contains('\n') && !value.Contains('\r')))
            {
                return value;
            }
            string normalized = value.Replace("\r\n", "\n").Replace("\r", "\n");
            return normalized.Replace("\n", "\r\n");
        }

        /// <summary>
        /// Method to save the workbook
        /// </summary>
        /// <exception cref="NanoXLSX.Exceptions.IOException">Throws IOException in case of an error</exception>
        /// <exception cref="RangeException">Throws a RangeException if the start or end address of a handled cell range was out of range</exception>
        /// <exception cref="NanoXLSX.Exceptions.FormatException">Throws a FormatException if a handled date cannot be translated to (Excel internal) OADate</exception>
        /// <exception cref="StyleException">Throws a StyleException if one of the styles of the workbook cannot be referenced or is null</exception>
        /// \remark <remarks>The StyleException should never happen in this state if the internally managed style collection was not tampered. </remarks>
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
        /// \remark <remarks>Possible Exceptions are <see cref="NanoXLSX.Exceptions.IOException">IOException</see>, <see cref="RangeException">RangeException</see>, <see cref="NanoXLSX.Exceptions.FormatException"></see> and <see cref="StyleException">StyleException</see>. These exceptions may not emerge directly if using the async method since async/await adds further abstraction layers.</remarks>
        /// <returns>Async Task</returns>
        public async Task SaveAsync()
        {
            await Task.Run(() => { Save(); });
        }

        private Dictionary<string, Dictionary<string, PackagePart>> packageParts = new Dictionary<string, Dictionary<string, PackagePart>>();
        private Dictionary<int, DocumentPath> worksheetPaths = new Dictionary<int, DocumentPath>();
        private Package package = null;

        private void PreparePackage()
        {
            int rootIndex = 1;
            int xlIndex = 1;
            // TODO: add themeIndex if once media is embedded
            PackagePart workbookPart = CreatePackagePart(package,
                WORKBOOK,
                @"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
                ref rootIndex);
            if (this.workbook.WorkbookMetadata != null)
            {
                CreatePackagePart(package,
                    CORE_PROPERTIES,
                    @"application/vnd.openxmlformats-package.core-properties+xml",
                    @"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
                    ref rootIndex);
                CreatePackagePart(package,
                    APP_PROPERTIES,
                    @"application/vnd.openxmlformats-officedocument.extended-properties+xml",
                    @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
                    ref rootIndex);
            }

            if (this.workbook.Worksheets.Count == 0)
            {
                // Fallback to default worksheet (seeht1.xml)
                DocumentPath path = new DocumentPath("sheet1.xml", "xl/worksheets");
                CreatePackagePart(workbookPart,
                    path,
                    @"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
                    @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                    ref xlIndex);
                worksheetPaths.Add(0, path);
            }
            else
            {
                for (int i = 0; i < this.Workbook.Worksheets.Count; i++)
                {
                    string fileName = "sheet" + ParserUtils.ToString(i + 1) + ".xml";
                    DocumentPath path = new DocumentPath(fileName, "xl/worksheets");
                    CreatePackagePart(workbookPart,
                        path,
                        @"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
                        @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                        ref xlIndex);
                    worksheetPaths.Add(i, path);
                }
            }

            if (workbook.WorkbookTheme != null)
            {
                CreatePackagePart(workbookPart,
                    THEME,
                    @"application/vnd.openxmlformats-officedocument.theme+xml",
                    @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
                    ref xlIndex);
            }

            CreatePackagePart(workbookPart,
                STYLES,
                @"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
                @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
                ref xlIndex);

            CreatePackagePart(workbookPart,
                SHARED_STRINGS,
                @"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
                @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
                ref xlIndex);
        }


        private PackagePart CreatePackagePart(object relationshipParent, DocumentPath documentPath, string contentType, string relationshipType, ref int index)
        {
            try
            {


                Uri uri = new Uri(documentPath.GetFullPath(), UriKind.Relative);
                PackagePart part = this.package.CreatePart(uri, contentType, CompressionOption.Normal);
                if (!packageParts.ContainsKey(documentPath.Path))
                {
                    packageParts.Add(documentPath.Path, new Dictionary<string, PackagePart>());
                }
                packageParts[documentPath.Path].Add(documentPath.Filename, part);
                if (relationshipParent == null || relationshipParent is Package)
                {
                    this.package.CreateRelationship(uri, TargetMode.Internal, relationshipType, "rId" + ParserUtils.ToString(index));
                }
                else if (relationshipParent is PackagePart)
                {
                    ((PackagePart)relationshipParent).CreateRelationship(uri, TargetMode.Internal, relationshipType, "rId" + ParserUtils.ToString(index));
                }
                index++;

                return part;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void SaveAsStream(Stream stream, bool leaveOpen = false)
        {
            workbook.ResolveMergedCells();
            this.styles = StyleManager.GetManagedStyles(workbook);
            try
            {
                using (Package package = Package.Open(stream, FileMode.Create))
                {
                    this.package = package;
                    PreparePackage();
                    PackagePart part;

                    // Workbook
                    IPluginWriter workbookWriter = PackageRegistry.GetWriter(new WorkbookWriter(this));
                    part = packageParts[WORKBOOK.Path][WORKBOOK.Filename];
                    AppendXmlToPackagePart(workbookWriter.CreateDocument(), part);

                    // Style
                    IPluginWriter styleWriter = PackageRegistry.GetWriter(new StyleWriter(this));
                    part = packageParts[STYLES.Path][STYLES.Filename];
                    AppendXmlToPackagePart(styleWriter.CreateDocument(), part);

                    // Shared strings - preparation
                    ISharedStringWriter sharedStringWriter = (ISharedStringWriter)PackageRegistry.GetWriter(new SharedStringWriter(this));

                    // Worksheets
                    IWorksheetWriter worksheetWriter = (IWorksheetWriter)PackageRegistry.GetWriter(new WorksheetWriter(this, (sharedStringWriter)));
                    if (workbook.Worksheets.Count > 0)
                    {
                        for (int i = 0; i < workbook.Worksheets.Count; i++)
                        {
                            Worksheet item = workbook.Worksheets[i];
                            part = packageParts[worksheetPaths[i].Path][worksheetPaths[i].Filename];
                            worksheetWriter.CurrentWorksheet = item;
                            AppendXmlToPackagePart(worksheetWriter.CreateDocument(), part);
                        }
                    }
                    else
                    {
                        part = packageParts[worksheetPaths[0].Path][worksheetPaths[0].Filename];
                        worksheetWriter.CurrentWorksheet = new Worksheet("sheet1");
                        AppendXmlToPackagePart(worksheetWriter.CreateDocument(), part);
                    }

                    // Shared strings - write after collection of strings
                    part = packageParts[SHARED_STRINGS.Path][SHARED_STRINGS.Filename];
                    AppendXmlToPackagePart(sharedStringWriter.CreateDocument(), part);

                    // Metadata
                    if (this.workbook.WorkbookMetadata != null)
                    {
                        IPluginWriter metadataAppWriter = PackageRegistry.GetWriter(new MetadataAppWriter(this));
                        part = packageParts[APP_PROPERTIES.Path][APP_PROPERTIES.Filename];
                        AppendXmlToPackagePart(metadataAppWriter.CreateDocument(), part);
                        IPluginWriter metadataCoreWriter = PackageRegistry.GetWriter(new MetadataCoreWriter(this));
                        part = packageParts[CORE_PROPERTIES.Path][CORE_PROPERTIES.Filename];
                        AppendXmlToPackagePart(metadataCoreWriter.CreateDocument(), part);
                    }

                    // Theme
                    if (workbook.WorkbookTheme != null)
                    {
                        IPluginWriter themeWriter = PackageRegistry.GetWriter(new ThemeWriter(this));
                        part = packageParts[THEME.Path][THEME.Filename];
                        AppendXmlToPackagePart(themeWriter.CreateDocument(), part);
                    }
                    package.Flush();
                    package.Close();
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
        /// \remark <remarks>Possible Exceptions are <see cref="IOException">IOException</see>, <see cref="RangeException">RangeException</see>, <see cref="FormatException"></see> and <see cref="StyleException">StyleException</see>. These exceptions may not emerge directly if using the async method since async/await adds further abstraction layers.</remarks>
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
        internal static void AppendXmlTag(StringBuilder sb, string value, string tagName, string nameSpace)
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
            sb.Append(XmlUtils.EscapeXmlChars(value));
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

        #endregion

    }
}
