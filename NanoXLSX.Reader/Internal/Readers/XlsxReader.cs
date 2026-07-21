/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Threading.Tasks;
using NanoXLSX.Exceptions;
using NanoXLSX.Interfaces.Reader;
using NanoXLSX.Registry;
using NanoXLSX.Styles;
using IOException = NanoXLSX.Exceptions.IOException;

namespace NanoXLSX.Internal.Readers
{
    /// <summary>
    /// Class representing a reader to decompile XLSX files
    /// </summary>
    public class XlsxReader : IDisposable
    {
        #region privateFields
        private readonly string filePath;
        private readonly Stream inputStream;
        private readonly ReaderOptions readerOptions;
        private MemoryStream memoryStream;
        #endregion

        #region properties
        /// <summary>
        /// Gets the read workbook
        /// </summary>
        public Workbook Workbook { get; internal set; }
        #endregion

        #region constructors
        /// <summary>
        /// Constructor with file path as parameter
        /// </summary>
        /// <param name="options">Reader options to override the automatic approach of the reader. <see cref="ReaderOptions"/> for information about Reader options.</param>
        /// <param name="path">File path of the XLSX file to load</param>
        public XlsxReader(string path, ReaderOptions options = null)
        {
            filePath = path;
            readerOptions = options;
        }

        /// <summary>
        /// Constructor with stream as parameter
        /// </summary>
        /// <param name="options">Reader options to override the automatic approach of the reader. <see cref="ReaderOptions"/> for information about Reader options.</param>
        /// <param name="stream">Stream of the XLSX file to load</param>
        public XlsxReader(Stream stream, ReaderOptions options = null)
        {
            inputStream = stream;
            readerOptions = options;
        }
        #endregion

        #region methods

        /// <summary>
        /// Reads the XLSX file from a file path or a file stream
        /// </summary>
        /// <exception cref="NanoXLSX.Exceptions.IOException">
        /// Throws IOException in case of an error
        /// </exception>
        public void Read()
        {
            try
            {
                using (memoryStream = new MemoryStream())
                {
                    Task.Run(() => ReadInternal()).GetAwaiter().GetResult();
                }
            }
            catch (NotSupportedContentException)
            {
                throw; // rethrow
            }
            catch (IOException)
            {
                throw; // rethrow
            }
            catch (Exception ex)
            {
                throw new IOException("There was an error while reading an XLSX file. Please see the inner exception:", ex);
            }
        }

        /// <summary>
        /// Reads the XLSX file from a file path or a file stream asynchronously
        /// </summary>
        /// <exception cref="NanoXLSX.Exceptions.IOException">
        /// May throw an IOException in case of an error. The asynchronous operation may hide the exception.
        /// </exception>
        /// <returns>Task object (void)</returns>
        public async Task ReadAsync()
        {
            try
            {
                using (memoryStream = new MemoryStream())
                {
                    await ReadInternal();
                }
            }
            catch (IOException)
            {
                throw; // rethrow
            }
            catch (Exception ex)
            {
                throw new IOException("There was an error while reading an XLSX file. Please see the inner exception:", ex);
            }
        }

        /// <summary>
        /// Reads a file or stream asynchronously
        /// </summary>
        /// <returns>Asynchronous task (void)</returns>
        private async Task ReadInternal()
        {
            ZipArchive zf;
            if (inputStream == null && !string.IsNullOrEmpty(filePath))
            {
                using (FileStream fs = new FileStream(filePath, FileMode.Open))
                {
                    await fs.CopyToAsync(memoryStream);
                }
            }
            else if (inputStream != null)
            {
                using (inputStream)
                {
                    await inputStream.CopyToAsync(memoryStream);
                }
            }
            else
            {
                throw new IOException("No valid stream or file path was provided to open");
            }

            memoryStream.Position = 0;
            zf = new ZipArchive(memoryStream, ZipArchiveMode.Read);

            await Task.Run(() =>
            {
                ReadZip(zf);
            }).ConfigureAwait(false);
        }

        /// <summary>
        /// Reads all compressed workbook entries in the provided ZipArchive
        /// </summary>
        /// <param name="zf">Zip archive, containing the workbook</param>
        private void ReadZip(ZipArchive zf)
        {
            MemoryStream ms;
            Workbook wb = new Workbook
            {
                importInProgress = true // Disables checks during load
            };
            Dictionary<string, ZipArchiveEntry> entryLookup = new Dictionary<string, ZipArchiveEntry>(zf.Entries.Count, StringComparer.Ordinal);
            foreach (ZipArchiveEntry entry in zf.Entries)
            {
                entryLookup[entry.FullName] = entry;
            }

            IDiscoveryReader discoveryReader = PlugInLoader.GetPlugIn<IDiscoveryReader>(PlugInUUID.DiscoveryReader, new DiscoveryReader());
            discoveryReader.Init(zf, wb, readerOptions);
            discoveryReader.Execute();
            RelationshipCatalog relationshipCatalog = wb.AuxiliaryData.GetData<RelationshipCatalog>(PlugInUUID.DiscoveryReader, PlugInUUID.DiscoveryCatalogEntity);
            if (relationshipCatalog == null)
            {
                throw new IOException("The relationship discovery reader did not provide a relationship catalog");
            }

            HandleQueuePlugIns(PlugInUUID.ReaderPackageRegistryQueue, entryLookup, relationshipCatalog, ref wb);
            HandleQueuePlugIns(PlugInUUID.ReaderPrependingQueue, entryLookup, relationshipCatalog, ref wb);

            IPluginBaseReader workbookReader = PlugInLoader.GetPlugIn<IPluginBaseReader>(PlugInUUID.WorkbookReader, new WorkbookReader());
            RelationshipInfo workbookRelationship = GetRelationship(relationshipCatalog, string.Empty, GetDocumentType(workbookReader, new WorkbookReader().DocumentType), true);
            string workbookPartPath = workbookRelationship.ResolvedTargetPath;

            ISharedStringReader sharedStringsReader = PlugInLoader.GetPlugIn<ISharedStringReader>(PlugInUUID.SharedStringsReader, new SharedStringsReader());
            RelationshipInfo sharedStringsRelationship = GetRelationship(relationshipCatalog, workbookPartPath, sharedStringsReader.DocumentType, false);
            if (sharedStringsRelationship != null)
            {
                ZipArchiveEntry sharedStringsEntry = GetRequiredEntry(sharedStringsRelationship, entryLookup);
                if (PlugInLoader.HasQueuePlugins(PlugInUUID.SharedStringsInlineReader))
                {
                    // Inline plugins need a seekable stream; buffer so the handler can reset position
                    MemoryStream ssMs = GetEntryStream(sharedStringsRelationship.ResolvedTargetPath, entryLookup);
                    sharedStringsReader.Init(ssMs, wb, readerOptions, ReaderPlugInHandler.HandleInlineQueuePlugins);
                    sharedStringsReader.Execute();
                }
                else
                {
                    // Direct-stream from ZIP entry — no intermediate MemoryStream
                    using (Stream sharedStringsStream = sharedStringsEntry.Open())
                    {
                        sharedStringsReader.Init(sharedStringsStream, wb, readerOptions, ReaderPlugInHandler.HandleInlineQueuePlugins);
                        sharedStringsReader.Execute();
                    }
                }
            }

            IPluginBaseReader themeReader = PlugInLoader.GetPlugIn<IPluginBaseReader>(PlugInUUID.ThemeReader, new ThemeReader());
            // Multiple workbook theme relationships are not defined clearly. Retain every relationship in the
            // discovery catalog, but preserve the previous reader behavior by processing only the first one.
            // "First" is deterministic and means XML order in the discovered workbook relationship part.
            RelationshipInfo themeRelationship = GetRelationship(relationshipCatalog, workbookPartPath, GetDocumentType(themeReader, new ThemeReader().DocumentType), false);
            if (themeRelationship != null)
            {
                ms = GetRequiredEntryStream(themeRelationship, entryLookup);
                themeReader.Init(ms, wb, readerOptions, ReaderPlugInHandler.HandleInlineQueuePlugins);
                themeReader.Execute();
            }

            StyleRepository.Instance.ImportInProgress = true; // TODO: To be checked
            IPluginBaseReader styleReader = PlugInLoader.GetPlugIn<IPluginBaseReader>(PlugInUUID.StyleReader, new StyleReader());
            RelationshipInfo styleRelationship = GetRelationship(relationshipCatalog, workbookPartPath, GetDocumentType(styleReader, new StyleReader().DocumentType), true);
            ms = GetRequiredEntryStream(styleRelationship, entryLookup);
            styleReader.Init(ms, wb, readerOptions, ReaderPlugInHandler.HandleInlineQueuePlugins);
            styleReader.Execute();
            StyleRepository.Instance.ImportInProgress = false;

            ms = GetRequiredEntryStream(workbookRelationship, entryLookup);
            workbookReader.Init(ms, wb, readerOptions, ReaderPlugInHandler.HandleInlineQueuePlugins);
            workbookReader.Execute();

            IPluginBaseReader metadataAppReader = PlugInLoader.GetPlugIn<IPluginBaseReader>(PlugInUUID.MetadataAppReader, new MetadataAppReader());
            RelationshipInfo metadataAppRelationship = GetRelationship(relationshipCatalog, string.Empty, GetDocumentType(metadataAppReader, new MetadataAppReader().DocumentType), false);
            if (metadataAppRelationship != null)
            {
                ms = GetRequiredEntryStream(metadataAppRelationship, entryLookup);
                metadataAppReader.Init(ms, wb, readerOptions, ReaderPlugInHandler.HandleInlineQueuePlugins);
                metadataAppReader.Execute();
            }

            IPluginBaseReader metadataCoreReader = PlugInLoader.GetPlugIn<IPluginBaseReader>(PlugInUUID.MetadataCoreReader, new MetadataCoreReader());
            RelationshipInfo metadataCoreRelationship = GetRelationship(relationshipCatalog, string.Empty, GetDocumentType(metadataCoreReader, new MetadataCoreReader().DocumentType), false);
            if (metadataCoreRelationship != null)
            {
                ms = GetRequiredEntryStream(metadataCoreRelationship, entryLookup);
                metadataCoreReader.Init(ms, wb, readerOptions, ReaderPlugInHandler.HandleInlineQueuePlugins);
                metadataCoreReader.Execute();
            }

            IPluginBaseReader relationships = PlugInLoader.GetPlugIn<IPluginBaseReader>(PlugInUUID.RelationshipReader, new RelationshipReader());
            ms = GetEntryStream(GetRelationshipPartPath(workbookPartPath), entryLookup);
            relationships.Init(ms, wb, readerOptions, ReaderPlugInHandler.HandleInlineQueuePlugins);
            relationships.Execute();

            IWorksheetReader worksheetReader = PlugInLoader.GetPlugIn<IWorksheetReader>(PlugInUUID.WorksheetReader, new WorksheetReader());
            worksheetReader.SharedStrings = sharedStringsReader.SharedStrings;
            int worksheetVisualIndex = 0;
            WorksheetDefinition definition;
            while ((definition = wb.AuxiliaryData.GetData<WorksheetDefinition>(PlugInUUID.WorkbookReader, PlugInUUID.WorksheetDefinitionEntity, worksheetVisualIndex)) != null)
            {
                RelationshipInfo relationship = relationshipCatalog.GetBySourceAndId(workbookPartPath, definition.RelId);
                if (relationship == null)
                {
                    throw new IOException("There was an error while reading an XLSX file. The relationship target of the worksheet with the RelID " + definition.RelId + " was not found");
                }
                if (relationship.TargetMode != System.IO.Packaging.TargetMode.Internal
                    || !entryLookup.TryGetValue(relationship.ResolvedTargetPath, out ZipArchiveEntry worksheetEntry))
                {
                    throw new IOException("There was an error while reading an XLSX file. The worksheet entry '" + relationship.ResolvedTargetPath + "' was not found in the archive");
                }
                worksheetReader.CurrentWorksheetID = worksheetVisualIndex;
                if (PlugInLoader.HasQueuePlugins(PlugInUUID.WorksheetInlineReader))
                {
                    // Inline plugins need a seekable stream; buffer so the handler can reset position
                    MemoryStream wsMs = GetEntryStream(relationship.ResolvedTargetPath, entryLookup);
                    worksheetReader.Init(wsMs, wb, readerOptions, ReaderPlugInHandler.HandleInlineQueuePlugins);
                    worksheetReader.Execute();
                }
                else
                {
                    // Direct-stream from ZIP entry — largest single allocation on big files
                    using (Stream worksheetStream = worksheetEntry.Open())
                    {
                        worksheetReader.Init(worksheetStream, wb, readerOptions, ReaderPlugInHandler.HandleInlineQueuePlugins);
                        worksheetReader.Execute();
                    }
                }
                worksheetVisualIndex++;
            }
            if (wb.Worksheets.Count == 0)
            {
                throw new IOException("No worksheet was found in the workbook");
            }
            HandleQueuePlugIns(PlugInUUID.ReaderAppendingQueue, entryLookup, relationshipCatalog, ref wb);
            wb.importInProgress = false; // Enables checks for runtime
            wb.AuxiliaryData.ClearTemporaryData(); // Remove temporary staging data
            this.Workbook = wb;
        }

        /// <summary>
        /// Gets a buffered MemoryStream of the specified file in the archive (XLSX file).
        /// The buffer is pre-allocated to the entry's uncompressed size for reduced allocation churn.
        /// </summary>
        /// <param name="name">Name of the XML file within the XLSX file</param>
        /// <param name="entryLookup">Pre-built lookup of archive entries by FullName</param>
        /// <returns>MemoryStream object of the specified file, or null if the entry was not found</returns>
        private static MemoryStream GetEntryStream(string name, Dictionary<string, ZipArchiveEntry> entryLookup)
        {
            if (!entryLookup.TryGetValue(name, out ZipArchiveEntry entry))
            {
                return null;
            }
            int capacity = (int)Math.Min(entry.Length, int.MaxValue);
            MemoryStream ms = new MemoryStream(capacity);
            using (Stream src = entry.Open())
            {
                src.CopyTo(ms);
            }
            ms.Position = 0;
            return ms;
        }

        private static RelationshipInfo GetRelationship(RelationshipCatalog catalog, string sourcePartPath, string documentType, bool required)
        {
            foreach (RelationshipInfo relationship in catalog.GetByType(documentType))
            {
                if (relationship.TargetMode == System.IO.Packaging.TargetMode.Internal
                    && string.Equals(relationship.SourcePartPath, sourcePartPath, StringComparison.Ordinal))
                {
                    return relationship;
                }
            }
            if (required)
            {
                string sourceDescription = string.IsNullOrEmpty(sourcePartPath) ? "the package root" : "'" + sourcePartPath + "'";
                throw new IOException("The required relationship type '" + documentType + "' was not found for " + sourceDescription);
            }
            return null;
        }

        private static ZipArchiveEntry GetRequiredEntry(RelationshipInfo relationship, Dictionary<string, ZipArchiveEntry> entryLookup)
        {
            if (relationship == null || string.IsNullOrEmpty(relationship.ResolvedTargetPath)
                || !entryLookup.TryGetValue(relationship.ResolvedTargetPath, out ZipArchiveEntry entry))
            {
                string target = relationship == null ? null : relationship.ResolvedTargetPath;
                throw new IOException("The relationship target entry '" + target + "' was not found in the archive");
            }
            return entry;
        }

        private static MemoryStream GetRequiredEntryStream(RelationshipInfo relationship, Dictionary<string, ZipArchiveEntry> entryLookup)
        {
            GetRequiredEntry(relationship, entryLookup);
            return GetEntryStream(relationship.ResolvedTargetPath, entryLookup);
        }

        private static string GetDocumentType(IPluginBaseReader reader, string defaultDocumentType)
        {
            IDocumentReader documentReader = reader as IDocumentReader;
            return documentReader == null || string.IsNullOrEmpty(documentReader.DocumentType)
                ? defaultDocumentType
                : documentReader.DocumentType;
        }

        private static string GetRelationshipPartPath(string sourcePartPath)
        {
            Uri sourcePartUri = new Uri("/" + sourcePartPath, UriKind.Relative);
            Uri relationshipPartUri = System.IO.Packaging.PackUriHelper.GetRelationshipPartUri(sourcePartUri);
            return relationshipPartUri.OriginalString.TrimStart('/');
        }

        /// <summary>
        /// Method to handle queue plug-ins
        /// </summary>
        /// <param name="queueUuid">Queue UUID</param>
        /// <param name="entryLookup">Pre-built lookup of archive entries by FullName</param>
        /// <param name="relationshipCatalog">Discovered package relationships used for discovery-aware queue readers</param>
        /// <param name="workbook">Workbook reference</param>
        private void HandleQueuePlugIns(string queueUuid, Dictionary<string, ZipArchiveEntry> entryLookup, RelationshipCatalog relationshipCatalog, ref Workbook workbook)
        {
            string lastUuid = null;
            IPluginQueueReader queueReader;
            do
            {
                string currentUuid;
                queueReader = PlugInLoader.GetNextQueuePlugIn<IPluginQueueReader>(queueUuid, lastUuid, out currentUuid);
                MemoryStream ms = null;
                if (queueReader != null)
                {
                    IDiscoveryPackageReader discoveryPackageReader = queueReader as IDiscoveryPackageReader;
                    if (discoveryPackageReader != null)
                    {
                        foreach (RelationshipInfo relationship in relationshipCatalog.GetByType(discoveryPackageReader.DocumentType))
                        {
                            if (relationship.TargetMode != System.IO.Packaging.TargetMode.Internal
                                || string.IsNullOrEmpty(relationship.ResolvedTargetPath)
                                || !entryLookup.ContainsKey(relationship.ResolvedTargetPath))
                            {
                                continue;
                            }
                            discoveryPackageReader.CurrentRelationship = relationship;
                            ms = GetEntryStream(relationship.ResolvedTargetPath, entryLookup);
                            discoveryPackageReader.Init(ms, workbook, this.readerOptions, null);
                            discoveryPackageReader.Execute();
                        }
                        lastUuid = currentUuid;
                        continue;
                    }
                    if (queueReader is IPluginPackageReader)
                    {
                        string streamPartName = (queueReader as IPluginPackageReader).StreamEntryName;
                        if (!string.IsNullOrEmpty(streamPartName))
                        {
                            ms = GetEntryStream(streamPartName, entryLookup);
                            if (ms == null)
                            {
                                lastUuid = currentUuid;
                                continue; // Skip if the stream part name was defined but not found
                            }
                        }
                    }
                    queueReader.Init(ms, workbook, this.readerOptions, null); // stream may be null, inlinePluginAction is not used here
                    queueReader.Execute();
                    lastUuid = currentUuid;
                }
                else
                {
                    lastUuid = null;
                }

            } while (queueReader != null);
        }

        /// <summary>
        /// Disposes the XlsxReader instance
        /// </summary>
        public void Dispose()
        {
            this.inputStream?.Dispose();
            GC.SuppressFinalize(this);
        }


        #endregion
    }
}
