/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using NanoXLSX.Exceptions;
using NanoXLSX.Interfaces.Writer;
using NanoXLSX.Internal.Structures;
using NanoXLSX.Registry;
using NanoXLSX.Styles;
using NanoXLSX.Utils;
using IOException = NanoXLSX.Exceptions.IOException;
using PackagePartType = NanoXLSX.Internal.Structures.PackagePartDefinition.PackagePartType;
using XmlElement = NanoXLSX.Utils.Xml.XmlElement;

namespace NanoXLSX.Internal.Writers
{
    /// <summary>
    /// Class for internal handling (XML, formatting, packing)
    /// </summary>
    /// \remark <remarks>This class is only for internal use. Use the high level API (e.g. class Workbook) to manipulate data and create Excel files</remarks>
    internal class XlsxWriter : IBaseWriter
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
        private int rootPackageIndex = 1;
        private int xlPackageIndex = 1;

        private readonly List<PackagePartDefinition> packagePartDefinitions = new List<PackagePartDefinition>();

        private readonly Dictionary<string, Dictionary<string, PackagePart>> packageParts = new Dictionary<string, Dictionary<string, PackagePart>>();
        private readonly Dictionary<int, DocumentPath> worksheetPaths = new Dictionary<int, DocumentPath>();
        private Package package = null;

        #endregion

        #region properties
        /// <summary>
        /// Workbook to be saved
        /// </summary>
        public Workbook Workbook
        {
            get { return workbook; }
        }

        /// <summary>
        /// Style manager attached to the workbook to save
        /// </summary>
        public StyleManager Styles
        {
            get { return styles; }
        }

        /// <summary>
        /// Shared string writer attached to the workbook to save
        /// </summary>
        public ISharedStringWriter SharedStringWriter { get; set; }

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
        /// Method to save the workbook
        /// </summary>
        /// <exception cref="IOException">Throws IOException in case of an error</exception>
        /// <exception cref="RangeException">Throws a RangeException if the start or end address of a handled cell range was out of range</exception>
        /// <exception cref="FormatException">Throws a FormatException if a handled date cannot be translated to (Excel internal) OADate</exception>
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

        /// <summary>
        /// Method to register the common / mandatory package parts of a XLSX file to be written
        /// </summary>
        private void RegisterCommonPackageParts()
        {
            // Workbook should always be the lowest index
            RegisterPackagePart(PackagePartType.Root, PackagePartDefinition.WORKBOOK_PACKAGE_PART_INDEX, WORKBOOK, @"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
            if (this.workbook.WorkbookMetadata != null)
            {
                int index = PackagePartDefinition.METADATA_PACKAGE_PART_START_INDEX;
                RegisterPackagePart(PackagePartType.Root, index, CORE_PROPERTIES, @"application/vnd.openxmlformats-package.core-properties+xml", @"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties");
                RegisterPackagePart(PackagePartType.Root, index + 1000, APP_PROPERTIES, @"application/vnd.openxmlformats-officedocument.extended-properties+xml", @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties");
            }
            int worksheetOrderNumber = PackagePartDefinition.WORKSHEET_PACKAGE_PART_START_INDEX;
            if (this.workbook.Worksheets.Count == 0)
            {
                RegisterPackagePart(PackagePartType.Worksheet, worksheetOrderNumber, "sheet1.xml", "xl/worksheets", @"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
            }
            else
            {
                for (int i = 0; i < this.workbook.Worksheets.Count; i++)
                {
                    string fileName = "sheet" + ParserUtils.ToString(i + 1) + ".xml";
                    RegisterPackagePart(PackagePartType.Worksheet, worksheetOrderNumber, fileName, "xl/worksheets", @"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
                    worksheetOrderNumber++;
                }
            }
            int postWorksheetOrderNumber = PackagePartDefinition.POST_WORSHEET_PACKAGE_PART_START_INDEX;
            if (workbook.WorkbookTheme != null)
            {
                RegisterPackagePart(PackagePartType.Other, postWorksheetOrderNumber, THEME, @"application/vnd.openxmlformats-officedocument.theme+xml", @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme");
                postWorksheetOrderNumber += 1000;
            }
            RegisterPackagePart(PackagePartType.Other, postWorksheetOrderNumber, STYLES, @"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml", @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
            postWorksheetOrderNumber += 1000;
            RegisterPackagePart(PackagePartType.Other, postWorksheetOrderNumber, SHARED_STRINGS, @"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml", @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings");
            // TODO: add themeIndex once if media is embedded
        }

        /// <summary>
        /// Method to prepare the package as source of the XLSX file to be written. Package parts are to be registered before calling this method
        /// </summary>
        /// <param name="package">Root package</param>
        private void PreparePackage(Package package)
        {
            List<PackagePartDefinition> definitions = PackagePartDefinition.Sort(this.packagePartDefinitions);
            PackagePartDefinition workbookDefinition = definitions.First(p => p.OrderNumber == PackagePartDefinition.WORKBOOK_PACKAGE_PART_INDEX);
            PackagePart workbookPart = CreateRootPackagePart(workbookDefinition.Path, workbookDefinition.ContentType, workbookDefinition.RelationshipType);
            foreach (PackagePartDefinition definition in definitions)
            {
                if (definition.OrderNumber == PackagePartDefinition.WORKBOOK_PACKAGE_PART_INDEX)
                {
                    continue;
                }
                if (definition.PartType == PackagePartType.Root)
                {
                    CreateRootPackagePart(definition.Path, definition.ContentType, definition.RelationshipType);
                }
                else
                {
                    CreateXlPackagePart(workbookPart, definition.Path, definition.ContentType, definition.RelationshipType);
                    if (definition.PartType == PackagePartType.Worksheet)
                    {
                        worksheetPaths.Add(definition.GetWorksheetIndex(), definition.Path);
                    }
                }
            }
        }

        /// <summary>
        /// Method to create root package parts, like workbook or the metadata parts
        /// </summary>
        /// <param name="documentPath">Document path of the part</param>
        /// <param name="contentType">Content type of the part</param>
        /// <param name="relationshipType">Scheme URL of the part</param>
        /// <returns>Created package part</returns>
        internal PackagePart CreateRootPackagePart(DocumentPath documentPath, string contentType, string relationshipType)
        {
            Uri uri = new Uri(documentPath.GetFullPath(), UriKind.Relative);
            PackagePart part = this.package.CreatePart(uri, contentType, CompressionOption.Normal);
            if (!packageParts.ContainsKey(documentPath.Path))
            {
                packageParts.Add(documentPath.Path, new Dictionary<string, PackagePart>());
            }
            packageParts[documentPath.Path].Add(documentPath.Filename, part);
            this.package.CreateRelationship(uri, TargetMode.Internal, relationshipType, "rId" + ParserUtils.ToString(rootPackageIndex));
            rootPackageIndex++;
            return part;
        }

        /// <summary>
        /// Method to create non-root package part, like worksheet or sharedStrings
        /// </summary>
        /// <param name="parentPart">Package part that is the parent of this part</param>
        /// <param name="documentPath">Document path of the part</param>
        /// <param name="contentType">Content type of the part</param>
        /// <param name="relationshipType">Scheme URL of the part</param>
        internal void CreateXlPackagePart(PackagePart parentPart, DocumentPath documentPath, string contentType, string relationshipType)
        {
            Uri uri = new Uri(documentPath.GetFullPath(), UriKind.Relative);
            PackagePart part = this.package.CreatePart(uri, contentType, CompressionOption.Normal);
            if (!packageParts.ContainsKey(documentPath.Path))
            {
                packageParts.Add(documentPath.Path, new Dictionary<string, PackagePart>());
            }
            packageParts[documentPath.Path].Add(documentPath.Filename, part);
            parentPart.CreateRelationship(uri, TargetMode.Internal, relationshipType, "rId" + ParserUtils.ToString(xlPackageIndex));
            xlPackageIndex++;
        }

        /// <summary>
        /// Method to register a package part with path and file name
        /// </summary>
        /// <param name="type">Type of the package part, used for handling differentiation</param>
        /// <param name="orderNumber">Order number during registration</param>
        /// <param name="fileNameInPackage">Relative file name of the target file of the package part, without path</param>
        /// <param name="pathInPackage">Relative path to the file of the package part</param>
        /// <param name="contentType">Content type of the target file of the part (usually kind of XML)</param>
        /// <param name="relationshipType">Schema URL of the target file of the part (usually kind of XML schema)</param>
        internal void RegisterPackagePart(PackagePartDefinition.PackagePartType type, int orderNumber, string fileNameInPackage, string pathInPackage, string contentType, string relationshipType)
        {
            this.packagePartDefinitions.Add(new PackagePartDefinition(type, orderNumber, fileNameInPackage, pathInPackage, contentType, relationshipType));
        }

        /// <summary>
        /// Method to register a package part with a document path
        /// </summary>
        /// <param name="type">Type of the package part, used for handling differentiation</param>
        /// <param name="orderNumber">Order number during registration</param>
        /// <param name="documentPath">Document path with all relevant file and path information</param>
        /// <param name="contentType">Content type of the target file of the part (usually kind of XML)</param>
        /// <param name="relationshipType">Schema URL of the target file of the part (usually kind of XML schema)</param>
        internal void RegisterPackagePart(PackagePartType type, int orderNumber, DocumentPath documentPath, string contentType, string relationshipType)
        {
            this.packagePartDefinitions.Add(new PackagePartDefinition(type, orderNumber, documentPath, contentType, relationshipType));
        }

        /// <summary>
        /// Method to save the workbook as stream.
        /// </summary>
        /// <param name="stream">Writable stream as target</param>
        /// <param name="leaveOpen">Optional parameter to keep the stream open after writing (used for MemoryStreams; default is false)</param>
        /// \remark <remarks>Possible Exceptions are <see cref="IOException">IOException</see>, <see cref="RangeException">RangeException</see>, <see cref="Exceptions.FormatException">FormatException</see> and <see cref="StyleException">StyleException</see>.</remarks>
        public void SaveAsStream(Stream stream, bool leaveOpen = false)
        {
            workbook.ResolveMergedCells();
            this.styles = StyleManager.GetManagedStyles(workbook);
            try
            {
                HandlePackageRegistryQueuePlugIns();
                HandleQueuePlugIns(PlugInUUID.WriterPrependingQueue);

                RegisterCommonPackageParts();
                using (Package xlsxPackage = Package.Open(stream, FileMode.Create))
                {
                    this.package = xlsxPackage;
                    PreparePackage(this.package);
                    PackagePart part;

                    // Workbook
                    IPlugInWriter workbookWriter = PlugInLoader.GetPlugIn<IPlugInWriter>(PlugInUUID.WorkbookWriter, new WorkbookWriter());
                    workbookWriter.Init(this);
                    workbookWriter.Execute();
                    part = packageParts[WORKBOOK.Path][WORKBOOK.Filename];
                    AppendXmlToPackagePart(workbookWriter.XmlElement, part);

                    // Style
                    IPlugInWriter styleWriter = PlugInLoader.GetPlugIn<IPlugInWriter>(PlugInUUID.StyleWriter, new StyleWriter());
                    styleWriter.Init(this);
                    styleWriter.Execute();
                    part = packageParts[STYLES.Path][STYLES.Filename];
                    AppendXmlToPackagePart(styleWriter.XmlElement, part);

                    // Shared strings - preparation
                    SharedStringWriter = PlugInLoader.GetPlugIn<ISharedStringWriter>(PlugInUUID.SharedStringsWriter, new SharedStringWriter());
                    SharedStringWriter.Init(this);
                    // Worksheets
                    IWorksheetWriter worksheetWriter = PlugInLoader.GetPlugIn<IWorksheetWriter>(PlugInUUID.WorksheetWriter, new WorksheetWriter());
                    worksheetWriter.Init(this);
                    if (workbook.Worksheets.Count > 0)
                    {
                        for (int i = 0; i < workbook.Worksheets.Count; i++)
                        {
                            Worksheet item = workbook.Worksheets[i];
                            part = packageParts[worksheetPaths[i].Path][worksheetPaths[i].Filename];
                            worksheetWriter.CurrentWorksheet = item;
                            worksheetWriter.Execute();
                            AppendXmlToPackagePart(worksheetWriter.XmlElement, part);
                        }
                    }
                    else
                    {
                        part = packageParts[worksheetPaths[0].Path][worksheetPaths[0].Filename];
                        worksheetWriter.CurrentWorksheet = new Worksheet("sheet1");
                        worksheetWriter.Execute();
                        AppendXmlToPackagePart(worksheetWriter.XmlElement, part);
                    }

                    // Shared strings - write after collection of strings
                    part = packageParts[SHARED_STRINGS.Path][SHARED_STRINGS.Filename];
                    SharedStringWriter.Execute();
                    AppendXmlToPackagePart(SharedStringWriter.XmlElement, part);

                    // Metadata
                    if (this.workbook.WorkbookMetadata != null)
                    {
                        IPlugInWriter metadataAppWriter = PlugInLoader.GetPlugIn<IPlugInWriter>(PlugInUUID.MetadataAppWriter, new MetadataAppWriter());
                        metadataAppWriter.Init(this);
                        metadataAppWriter.Execute();
                        part = packageParts[APP_PROPERTIES.Path][APP_PROPERTIES.Filename];
                        AppendXmlToPackagePart(metadataAppWriter.XmlElement, part);
                        IPlugInWriter metadataCoreWriter = PlugInLoader.GetPlugIn<IPlugInWriter>(PlugInUUID.MetadataCoreWriter, new MetadataCoreWriter());
                        metadataCoreWriter.Init(this);
                        metadataCoreWriter.Execute();
                        part = packageParts[CORE_PROPERTIES.Path][CORE_PROPERTIES.Filename];
                        AppendXmlToPackagePart(metadataCoreWriter.XmlElement, part);
                    }

                    // Theme
                    if (workbook.WorkbookTheme != null)
                    {
                        IPlugInWriter themeWriter = PlugInLoader.GetPlugIn<IPlugInWriter>(PlugInUUID.ThemeWriter, new ThemeWriter());
                        themeWriter.Init(this);
                        themeWriter.Execute();
                        part = packageParts[THEME.Path][THEME.Filename];
                        AppendXmlToPackagePart(themeWriter.XmlElement, part);
                    }

                    HandleQueuePlugIns(PlugInUUID.WriterAppendingQueue);

                    this.package.Flush();
                    this.package.Close();
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
        /// \remark <remarks>Possible Exceptions are <see cref="IOException">IOException</see>, <see cref="RangeException">RangeException</see>, <see cref="Exceptions.FormatException">FormatException</see> and <see cref="StyleException">StyleException</see>. These exceptions may not emerge directly if using the async method since async/await adds further abstraction layers.</remarks>
        /// <returns>Async Task</returns>
        public async Task SaveAsStreamAsync(Stream stream, bool leaveOpen = false)
        {
            await Task.Run(() => { SaveAsStream(stream, leaveOpen); });
        }
        #endregion

        /// <summary>
        /// Method to handle queue plug-ins
        /// </summary>
        /// <param name="queueUuid">Queue UUID</param>
        private void HandleQueuePlugIns(string queueUuid)
        {
            IPlugInWriter queueWriter = null;
            string lastUuid = null;
            do
            {
                string currentUuid;
                queueWriter = PlugInLoader.GetNextQueuePlugIn<IPlugInWriter>(queueUuid, lastUuid, out currentUuid);
                if (queueWriter != null)
                {
                    queueWriter.Init(this);
                    queueWriter.Execute();
                    if (queueWriter is IPlugInPackageWriter packageWriter)
                    {
                        if (!string.IsNullOrEmpty(packageWriter.PackagePartPath) && !string.IsNullOrEmpty(packageWriter.PackagePartFileName))
                        {
                            if (packageParts.ContainsKey(packageWriter.PackagePartPath) && packageParts[packageWriter.PackagePartPath].ContainsKey(packageWriter.PackagePartFileName))
                            {
                                PackagePart pp = packageParts[packageWriter.PackagePartPath][packageWriter.PackagePartFileName];
                                AppendXmlToPackagePart(packageWriter.XmlElement, pp);
                            }
                        }
                    }
                    lastUuid = currentUuid;
                }
                else
                {
                    lastUuid = null;
                }

            } while (queueWriter != null);
        }

        /// <summary>
        /// Method to handle queue plug-ins that are registering package parts
        /// </summary>
        private void HandlePackageRegistryQueuePlugIns()
        {
            IPlugInPackageWriter queueWriter = null;
            string lastUuid = null;
            do
            {
                string currentUuid;
                queueWriter = PlugInLoader.GetNextQueuePlugIn<IPlugInPackageWriter>(PlugInUUID.WriterPackageRegistryQueue, lastUuid, out currentUuid);
                if (queueWriter != null)
                {
                    queueWriter.Execute(); // Execute anything that could be defined
                    PackagePartType packagePartType;
                    if (queueWriter.IsRootPackagePart)
                    {
                        packagePartType = PackagePartType.Root;
                    }
                    else
                    {
                        packagePartType = PackagePartType.Other;
                    }
                    RegisterPackagePart(packagePartType, queueWriter.OrderNumber, new DocumentPath(queueWriter.PackagePartFileName, queueWriter.PackagePartPath), queueWriter.ContentType, queueWriter.RelationshipType);
                    lastUuid = currentUuid;
                }
                else
                {
                    lastUuid = null;
                }

            } while (queueWriter != null);
        }

        /// <summary>
        /// Method to append XML files to a root package part in the right hierarchy
        /// </summary>
        /// <param name="rootElement">Root element</param>
        /// <param name="pp">Package part</param>
        private void AppendXmlToPackagePart(XmlElement rootElement, PackagePart pp)
        {
            XmlDocument doc = rootElement.TransformToDocument(); // This creates a System.Xml.XmlDocument from a custom XmlElement instance
            using (MemoryStream ms = new MemoryStream())
            {
                XmlWriterSettings settings = new XmlWriterSettings
                {
                    Encoding = new UTF8Encoding(false), // No BOM
                    Indent = true,
                    OmitXmlDeclaration = false // Include <?xml version="1.0" encoding="utf-8"?>
                };

                using (XmlWriter writer = XmlWriter.Create(ms, settings))
                {
                    writer.WriteProcessingInstruction("xml", "version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"");
                    doc.WriteTo(writer);
                    writer.Flush();
                }

                AddStreamToPackagePart(ms, pp);
            }
        }

        /// <summary>
        /// Method to add a stream to a package part
        /// </summary>
        /// <param name="stream">Stream to add</param>
        /// <param name="pp">Package part</param>
        internal void AddStreamToPackagePart(MemoryStream stream, PackagePart pp)
        {
            stream.Position = 0;
            stream.CopyTo(pp.GetStream());
            stream.Flush();
        }

    }
}
