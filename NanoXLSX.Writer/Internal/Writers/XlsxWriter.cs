/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Exceptions;
using NanoXLSX.Interfaces;
using NanoXLSX.Interfaces.Writer;
using NanoXLSX.Internal.Structures;
using NanoXLSX.Registry;
using NanoXLSX.Styles;
using NanoXLSX.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
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
        private static readonly DocumentPath WORKBOOK = new DocumentPath("workbook.xml", "xl/");
        private static readonly DocumentPath STYLES = new DocumentPath("styles.xml", "xl/");
        private static readonly DocumentPath APP_PROPERTIES = new DocumentPath("app.xml", "docProps/");
        private static readonly DocumentPath CORE_PROPERTIES = new DocumentPath("core.xml", "docProps/");
        private static readonly DocumentPath SHARED_STRINGS = new DocumentPath("sharedStrings.xml", "xl/");
        private static readonly DocumentPath THEME = new DocumentPath("theme1.xml", "xl/theme/");
        #endregion

        #region privateFields
        private int rootPackageIndex = 1;
        private int xlPackageIndex = 1;
        private Package package = null;

        private readonly List<PackagePartDefinition> packagePartDefinitions = new List<PackagePartDefinition>();
        private readonly Dictionary<string, Dictionary<string, PackagePart>> packageParts = new Dictionary<string, Dictionary<string, PackagePart>>();
        private readonly Dictionary<int, DocumentPath> worksheetPaths = new Dictionary<int, DocumentPath>();
        private readonly HashSet<string> preparedWriterFeatures = new HashSet<string>();
        private readonly Dictionary<string, PackagePart> queuedPackageParts = new Dictionary<string, PackagePart>(StringComparer.Ordinal);

        #endregion

        #region properties
        /// <summary>
        /// Workbook to be saved
        /// </summary>
        public Workbook Workbook { get; }

        /// <summary>
        /// Processing data outside of the workbook (e.g. used for preparation)
        /// </summary>
        public IWriterProcessingData WriterProcessingData { get; set; }

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
            this.Workbook = workbook;
        }
        #endregion

        #region documentCreation_methods

        /// <summary>
        /// Method to save the workbook
        /// </summary>
        /// <exception cref="IOException">Throws IOException in case of an error</exception>
        /// <exception cref="RangeException">Throws a RangeException if the start or end address of a handled cell range was out of range</exception>
        /// <exception cref="Exceptions.FormatException">Throws a FormatException if a handled date cannot be translated to (Excel internal) OADate</exception>
        /// <exception cref="StyleException">Throws a StyleException if one of the styles of the workbook cannot be referenced or is null</exception>
        /// \remark <remarks>The StyleException should never happen in this state if the internally managed style collection was not tampered. </remarks>
        public void Save()
        {
            try
            {
                FileStream fs = new FileStream(Workbook.Filename, FileMode.Create);
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

        /// <summary>
        /// Method to save the workbook as stream.
        /// </summary>
        /// <param name="stream">Writable stream as target</param>
        /// <param name="leaveOpen">Optional parameter to keep the stream open after writing (used for MemoryStreams; default is false)</param>
        /// \remark <remarks>Possible Exceptions are <see cref="IOException">IOException</see>, <see cref="RangeException">RangeException</see>, <see cref="Exceptions.FormatException">FormatException</see> and <see cref="StyleException">StyleException</see>.</remarks>
        public void SaveAsStream(Stream stream, bool leaveOpen = false)
        {
            preparedWriterFeatures.Clear();
            WriterProcessingData = new WriterProcessingData(Workbook, StyleRepository.Instance);
            try
            {
                // preparing processor(s)
                IPluginWriteProcessor preparingProcessor = PlugInLoader.GetPlugIn<IPluginWriteProcessor>(PlugInUUID.PreparingProcessor, new PreparingProcessor());
                preparingProcessor.Init(this, WriterPlugInHandler.HandleInlineQueueProcessorPlugins);
                preparingProcessor.Execute();
                // Compatibility check
                IPluginWriteProcessor compatibilityProcessor = new CompatibilityProcessor(); // This core processor cannot be overwritten
                compatibilityProcessor.Init(this, WriterPlugInHandler.HandleInlineQueueProcessorPlugins);
                compatibilityProcessor.Execute();
                // Workbook can now be written
                RegisterCommonPackageParts();
                HandlePackageRegistryQueuePlugIns();
                HandleQueuePlugIns(PlugInUUID.WriterPrependingQueue);

                using (Package xlsxPackage = Package.Open(stream, FileMode.Create))
                {
                    this.package = xlsxPackage;
                    PreparePackage();
                    PackagePart part;

                    // Workbook
                    IPluginWriter workbookWriter = PlugInLoader.GetPlugIn<IPluginWriter>(PlugInUUID.WorkbookWriter, new WorkbookWriter());
                    workbookWriter.Init(this);
                    workbookWriter.Execute();
                    part = packageParts[WORKBOOK.Path][WORKBOOK.Filename];
                    AppendXmlToPackagePart(workbookWriter.XmlElement, part);

                    // Style
                    IPluginWriter styleWriter = PlugInLoader.GetPlugIn<IPluginWriter>(PlugInUUID.StyleWriter, new StyleWriter());
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
                    if (Workbook.Worksheets.Count > 0)
                    {
                        for (int i = 0; i < Workbook.Worksheets.Count; i++)
                        {
                            Worksheet item = Workbook.Worksheets[i];
                            part = packageParts[worksheetPaths[i].Path][worksheetPaths[i].Filename];
                            worksheetWriter.CurrentWorksheet = item;
                            worksheetWriter.Execute();
                            AppendXmlToPackagePart(worksheetWriter.XmlElement, part);
                            worksheetWriter.ReleaseXmlElement();
                            GC.Collect(1, GCCollectionMode.Optimized); // 
                        }
                    }
                    else
                    {
                        part = packageParts[worksheetPaths[0].Path][worksheetPaths[0].Filename];
                        worksheetWriter.CurrentWorksheet = new Worksheet("sheet1");
                        worksheetWriter.Execute();
                        AppendXmlToPackagePart(worksheetWriter.XmlElement, part);
                        worksheetWriter.ReleaseXmlElement();
                    }

                    // Shared strings - write after collection of strings
                    part = packageParts[SHARED_STRINGS.Path][SHARED_STRINGS.Filename];
                    SharedStringWriter.Execute();
                    AppendXmlToPackagePart(SharedStringWriter.XmlElement, part);

                    // Metadata
                    if (this.Workbook.WorkbookMetadata != null)
                    {
                        IPluginWriter metadataAppWriter = PlugInLoader.GetPlugIn<IPluginWriter>(PlugInUUID.MetadataAppWriter, new MetadataAppWriter());
                        metadataAppWriter.Init(this);
                        metadataAppWriter.Execute();
                        part = packageParts[APP_PROPERTIES.Path][APP_PROPERTIES.Filename];
                        AppendXmlToPackagePart(metadataAppWriter.XmlElement, part);
                        IPluginWriter metadataCoreWriter = PlugInLoader.GetPlugIn<IPluginWriter>(PlugInUUID.MetadataCoreWriter, new MetadataCoreWriter());
                        metadataCoreWriter.Init(this);
                        metadataCoreWriter.Execute();
                        part = packageParts[CORE_PROPERTIES.Path][CORE_PROPERTIES.Filename];
                        AppendXmlToPackagePart(metadataCoreWriter.XmlElement, part);
                    }

                    // Theme
                    if (Workbook.WorkbookTheme != null)
                    {
                        IPluginWriter themeWriter = PlugInLoader.GetPlugIn<IPluginWriter>(PlugInUUID.ThemeWriter, new ThemeWriter());
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
                Workbook.AuxiliaryData.ClearTemporaryData();
            }
            catch (Exception e)
            {
                throw new IOException("An error occurred while saving. See inner exception for details: " + e.Message, e);
            }
        }

        /// <summary>
        /// Method to register the common / mandatory package parts of a XLSX file to be written
        /// </summary>
        private void RegisterCommonPackageParts()
        {
            // Workbook should always be the lowest index
            RegisterPackagePart(PackagePartType.Root, PackagePartDefinition.WORKBOOK_PACKAGE_PART_INDEX, WORKBOOK, @"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
            if (this.Workbook.WorkbookMetadata != null)
            {
                int index = PackagePartDefinition.METADATA_PACKAGE_PART_START_INDEX;
                RegisterPackagePart(PackagePartType.Root, index, CORE_PROPERTIES, @"application/vnd.openxmlformats-package.core-properties+xml", @"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties");
                RegisterPackagePart(PackagePartType.Root, index + 1000, APP_PROPERTIES, @"application/vnd.openxmlformats-officedocument.extended-properties+xml", @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties");
            }
            int worksheetOrderNumber = PackagePartDefinition.WORKSHEET_PACKAGE_PART_START_INDEX;
            if (this.Workbook.Worksheets.Count == 0)
            {
                RegisterPackagePart(PackagePartType.Worksheet, worksheetOrderNumber, "sheet1.xml", "xl/worksheets", @"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
            }
            else
            {
                for (int i = 0; i < this.Workbook.Worksheets.Count; i++)
                {
                    string fileName = "sheet" + ParserUtils.ToString(i + 1) + ".xml";
                    RegisterPackagePart(PackagePartType.Worksheet, worksheetOrderNumber, fileName, "xl/worksheets", @"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
                    worksheetOrderNumber++;
                }
            }
            int postWorksheetOrderNumber = PackagePartDefinition.POST_WORkSHEET_PACKAGE_PART_START_INDEX;
            if (Workbook.WorkbookTheme != null)
            {
                RegisterPackagePart(PackagePartType.Other, postWorksheetOrderNumber, THEME, @"application/vnd.openxmlformats-officedocument.theme+xml", @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme");
                postWorksheetOrderNumber += 1000;
            }
            RegisterPackagePart(PackagePartType.Other, postWorksheetOrderNumber, STYLES, @"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml", @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
            postWorksheetOrderNumber += 1000;
            RegisterPackagePart(PackagePartType.Other, postWorksheetOrderNumber, SHARED_STRINGS, @"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml", @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings");
            // TODO: add themeIndex once if media is embedded

            this.Workbook.AuxiliaryData.SetData(PlugInUUID.WriterPackageRegistryQueue, PlugInUUID.LastPackageOrderNumber, postWorksheetOrderNumber, false);
        }

        /// <summary>
        /// Method to prepare the package as source of the XLSX file to be written. Package parts are to be registered before calling this method
        /// </summary>
        private void PreparePackage()
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
                PackagePart createdPart;
                if (definition.PartType == PackagePartType.Root)
                {
                    createdPart = CreateRootPackagePart(definition.Path, definition.ContentType, definition.RelationshipType);
                }
                else
                {
                    createdPart = CreateXlPackagePart(workbookPart, definition.Path, definition.ContentType, definition.RelationshipType);
                    if (definition.PartType == PackagePartType.Worksheet)
                    {
                        worksheetPaths.Add(definition.GetWorksheetIndex(), definition.Path);
                    }
                }
                if (definition.UniquePackagePartIndex != null)
                {
                    queuedPackageParts.Add(definition.UniquePackagePartIndex, createdPart);
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
        internal PackagePart CreateXlPackagePart(PackagePart parentPart, DocumentPath documentPath, string contentType, string relationshipType)
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
            return part;
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
        /// Method to register a package part for a queued indexed writer
        /// </summary>
        /// <param name="type">Type of the package part, used for handling differentiation</param>
        /// <param name="orderNumber">Order number during registration</param>
        /// <param name="documentPath">Document path with all relevant file and path information</param>
        /// <param name="contentType">Content type of the target file of the part (usually kind of XML)</param>
        /// <param name="relationshipType">Schema URL of the target file of the part (usually kind of XML schema)</param>
        /// <param name="uniquePackagePartIndex">Unique index used by a queued writer to select this package part</param>
        private void RegisterPackagePart(PackagePartType type, int orderNumber, DocumentPath documentPath, string contentType, string relationshipType, string uniquePackagePartIndex)
        {
            this.packagePartDefinitions.Add(new PackagePartDefinition(type, orderNumber, documentPath, contentType, relationshipType, uniquePackagePartIndex));
        }

        #endregion

        #region interface_methodes

        /// <summary>
        /// Marks a feature, defined by its UUID, as prepared, so that a writer can handle it
        /// </summary>
        /// <param name="featureUuid">Feature UUID to be set as prepared</param>
        public void MarkFeatureAsPrepared(string featureUuid)
        {
            ValidateFeatureUuid(featureUuid);
            preparedWriterFeatures.Add(featureUuid);
        }

        /// <summary>
        /// Gets whether a feature UUID was marked as prepared, and therefore can be used in the writing process
        /// </summary>
        /// <param name="featureUuid">Feature UUID to be checked</param>
        /// <returns>True if the feature was defined and was marked as prepared</returns>
        public bool IsFeaturePrepared(string featureUuid)
        {
            ValidateFeatureUuid(featureUuid);
            return preparedWriterFeatures.Contains(featureUuid);
        }

        /// <summary>
        /// Validates a given feature UUID
        /// </summary>
        /// <param name="uuid">UUID to validate</param>
        /// <exception cref="ArgumentException">Thrown if the feature UUID is not valid</exception>
        private void ValidateFeatureUuid(string uuid)
        {
            if (string.IsNullOrWhiteSpace(uuid))
            {
                throw new ArgumentException("The feature UUID must not be null, empty or whitespace");
            }
            // TODO add other validation checks here if applicable (e.g. UUID must be officially registered)
        }

        #endregion

        #region helper_methods
        /// <summary>
        /// Method to handle queue plug-ins
        /// </summary>
        /// <param name="queueUuid">Queue UUID</param>
        private void HandleQueuePlugIns(string queueUuid)
        {
            IPlugin queuePlugIn;
            string lastUuid = null;
            do
            {
                queuePlugIn = PlugInLoader.GetNextQueuePlugIn<IPlugin>(queueUuid, lastUuid, out string currentUuid);
                if (queuePlugIn != null)
                {
                    lastUuid = currentUuid;
                    if (!(queuePlugIn is IPluginWriter queueWriter))
                    {
                        continue;
                    }
                    queueWriter.Init(this);
                    if (queueWriter is IPluginIndexedWriter indexedWriter)
                    {
                        HandleIndexedWriter(indexedWriter, queueUuid);
                    }
                    else
                    {
                        queueWriter.Execute();
                    }
                }
                else
                {
                    lastUuid = null;
                }

            } while (queuePlugIn != null);
        }

        /// <summary>
        /// Executes an indexed writer and appends each generated XML element to its registered package part
        /// </summary>
        /// <param name="indexedWriter">Indexed writer to execute</param>
        /// <param name="queueUuid">UUID of the queue currently being handled</param>
        private void HandleIndexedWriter(IPluginIndexedWriter indexedWriter, string queueUuid)
        {
            int maxIndex = indexedWriter.MaxIndex;
            if (maxIndex < -1)
            {
                throw new IOException("Invalid maximum index in indexed writer plug-in: " + indexedWriter.GetType().Name);
            }
            if (maxIndex == -1)
            {
                return;
            }
            if (queueUuid == PlugInUUID.WriterPrependingQueue)
            {
                throw new IOException("Indexed writer plug-ins cannot be executed in the writer prepending queue: " + indexedWriter.GetType().Name);
            }

            for (int i = 0; i <= maxIndex; i++)
            {
                indexedWriter.CurrentIndex = i;
                indexedWriter.Execute();
                string uniquePackagePartIndex = indexedWriter.CurrentUniquePackagePartIndex;
                if (uniquePackagePartIndex == null)
                {
                    continue;
                }
                if (string.IsNullOrWhiteSpace(uniquePackagePartIndex))
                {
                    throw new IOException("Blank package part index in indexed writer plug-in: " + indexedWriter.GetType().Name);
                }
                if (!queuedPackageParts.TryGetValue(uniquePackagePartIndex, out PackagePart packagePart))
                {
                    throw new IOException("Unknown package part index '" + uniquePackagePartIndex + "' in indexed writer plug-in: " + indexedWriter.GetType().Name);
                }
                if (indexedWriter.XmlElement == null)
                {
                    throw new IOException("Missing XML element in indexed writer plug-in: " + indexedWriter.GetType().Name);
                }
                AppendXmlToPackagePart(indexedWriter.XmlElement, packagePart);
            }
        }

        /// <summary>
        /// Method to handle queue plug-ins that are registering package parts
        /// </summary>
        private void HandlePackageRegistryQueuePlugIns()
        {
            IPlugin queuePlugIn;
            string lastUuid = null;
            do
            {
                queuePlugIn = PlugInLoader.GetNextQueuePlugIn<IPlugin>(PlugInUUID.WriterPackageRegistryQueue, lastUuid, out string currentUuid);
                if (queuePlugIn != null)
                {
                    lastUuid = currentUuid;
                    if (!(queuePlugIn is IPluginPackageRegistry packageRegistry))
                    {
                        continue;
                    }
                    packageRegistry.Init(this);
                    packageRegistry.Execute();
                    int counter = ValidatePackageRegistryPlugin(packageRegistry);
                    for (int i = 0; i < counter; i++)
                    {
                        ValidatePackageRegistryEntry(packageRegistry, i);
                        PackagePartType packagePartType = packageRegistry.ArePackagePartsRoot[i] ? PackagePartType.Root : PackagePartType.Other;
                        RegisterPackagePart(
                            packagePartType,
                            packageRegistry.OrderNumbers[i],
                            new DocumentPath(packageRegistry.PackagePartFileNames[i], packageRegistry.PackagePartPaths[i]),
                            packageRegistry.ContentTypes[i],
                            packageRegistry.RelationshipTypes[i],
                            packageRegistry.UniquePackagePartIndices[i]);
                    }
                }
                else
                {
                    lastUuid = null;
                }

            } while (queuePlugIn != null);
        }

        /// <summary>
        /// Validates the collection structure of a package registry plug-in
        /// </summary>
        /// <param name="plugin">Package registry plug-in to validate</param>
        /// <returns>Number of registered package parts</returns>
        private static int ValidatePackageRegistryPlugin(IPluginPackageRegistry plugin)
        {
            if (plugin.OrderNumbers == null ||
                plugin.ArePackagePartsRoot == null ||
                plugin.ContentTypes == null ||
                plugin.PackagePartFileNames == null ||
                plugin.PackagePartPaths == null ||
                plugin.RelationshipTypes == null ||
                plugin.UniquePackagePartIndices == null)
            {
                throw new IOException("Null collection in package registry plug-in: " + plugin.GetType().Name);
            }

            int count = plugin.OrderNumbers.Count;
            if (plugin.ArePackagePartsRoot.Count != count ||
                plugin.ContentTypes.Count != count ||
                plugin.PackagePartFileNames.Count != count ||
                plugin.PackagePartPaths.Count != count ||
                plugin.RelationshipTypes.Count != count ||
                plugin.UniquePackagePartIndices.Count != count)
            {
                throw new IOException("Inconsistent package registry plug-in detected: " + plugin.GetType().Name);
            }
            return count;
        }

        /// <summary>
        /// Validates one package part definition supplied by a registry plug-in
        /// </summary>
        /// <param name="plugin">Package registry plug-in to validate</param>
        /// <param name="index">Index of the definition to validate</param>
        private void ValidatePackageRegistryEntry(IPluginPackageRegistry plugin, int index)
        {
            string uniquePackagePartIndex = plugin.UniquePackagePartIndices[index];
            if (string.IsNullOrWhiteSpace(uniquePackagePartIndex))
            {
                throw new IOException("Blank package part index in package registry plug-in: " + plugin.GetType().Name);
            }
            if (packagePartDefinitions.Any(definition => string.Equals(definition.UniquePackagePartIndex, uniquePackagePartIndex, StringComparison.Ordinal)))
            {
                throw new IOException("Duplicate package part index '" + uniquePackagePartIndex + "' in package registry plug-in: " + plugin.GetType().Name);
            }
            if (string.IsNullOrWhiteSpace(plugin.PackagePartPaths[index]) ||
                string.IsNullOrWhiteSpace(plugin.PackagePartFileNames[index]) ||
                string.IsNullOrWhiteSpace(plugin.ContentTypes[index]) ||
                string.IsNullOrWhiteSpace(plugin.RelationshipTypes[index]))
            {
                throw new IOException("Invalid package part definition in package registry plug-in: " + plugin.GetType().Name);
            }
        }

        /// <summary>
        /// Method to append XML files to a root package part in the right hierarchy
        /// </summary>
        /// <param name="rootElement">Root element</param>
        /// <param name="pp">Package part</param>
        private void AppendXmlToPackagePart(XmlElement rootElement, PackagePart pp)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                XmlWriterSettings settings = new XmlWriterSettings
                {
                    Encoding = new UTF8Encoding(false), // No BOM
                    Indent = true,
                    OmitXmlDeclaration = true, // Include <?xml version="1.0" encoding="utf-8"?>
                    CloseOutput = false
                };

                using (XmlWriter writer = XmlWriter.Create(ms, settings))
                {
                    writer.WriteProcessingInstruction("xml", "version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"");
                    rootElement.WriteTo(writer);
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

        #endregion
    }
}
