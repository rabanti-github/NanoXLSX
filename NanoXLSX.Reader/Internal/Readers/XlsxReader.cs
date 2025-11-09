/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using NanoXLSX.Exceptions;
using NanoXLSX.Interfaces.Plugin;
using NanoXLSX.Interfaces.Reader;
using NanoXLSX.Registry;
using NanoXLSX.Styles;
using NanoXLSX.Utils;
using IOException = NanoXLSX.Exceptions.IOException;

namespace NanoXLSX.Internal.Readers
{
    /// <summary>
    /// Class representing a reader to decompile XLSX files
    /// </summary>
    public class XlsxReader
    {
        #region privateFields
        private string filePath;
        private Stream inputStream;
        private MemoryStream memoryStream;
        private ReaderOptions readerOptions;
        #endregion

        #region properties
        /// <summary>
        /// Gets the read workbook
        /// </summary>
        public Workbook Workbook { get; internal set; } = null;
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
                    ReadInternal().GetAwaiter().GetResult();
                }
            }
            catch (NotSupportedContentException ex)
            {
                throw ex; // rethrow
            }
            catch (IOException ex)
            {
                throw ex; // rethrow
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
            catch (IOException ex)
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
            Workbook wb = new Workbook();
            wb.importInProgress = true; // Disables checks during load
            HandleQueuePlugIns(PlugInUUID.READER_PREPENDING_QUEUE, zf, ref wb);

            ISharedStringReader sharedStringsReader = PlugInLoader.GetPlugIn<ISharedStringReader>(PlugInUUID.SHARED_STRINGS_READER, new SharedStringsReader());
            ms = GetEntryStream("xl/sharedStrings.xml", zf);
            if (ms != null && ms.Length > 0) // If length == 0, no shared strings are defined (no text in file)
            {
                sharedStringsReader.Init(ms, wb, readerOptions);
                sharedStringsReader.Execute();
            }
            Dictionary<int, string> themeStreamNames = GetSequentialStreamNames("xl/theme/theme", zf);
            if (themeStreamNames.Count > 0)
            {
                // There is not really a definition whether multiple themes can be managed in one workbook.
                // the suffix number (e.g. theme1) indicates it. However, no examples were found and therefore
                // (currently) only the first occurring theme will be read  
                foreach (KeyValuePair<int, string> streamName in themeStreamNames)
                {
                    IPlugInReader themeReader = PlugInLoader.GetPlugIn<IPlugInReader>(PlugInUUID.THEME_READER, new ThemeReader());
                    ms = GetEntryStream(streamName.Value, zf);
                    themeReader.Init(ms, wb, readerOptions);
                    themeReader.Execute();
                    break;
                }
            }
            StyleRepository.Instance.ImportInProgress = true; // TODO: To be checked
            IPlugInReader styleReader = PlugInLoader.GetPlugIn<IPlugInReader>(PlugInUUID.STYLE_READER, new StyleReader());
            ms = GetEntryStream("xl/styles.xml", zf);
            styleReader.Init(ms, wb, readerOptions);
            styleReader.Execute();
            StyleRepository.Instance.ImportInProgress = false;

            ms = GetEntryStream("xl/workbook.xml", zf);
            IPlugInReader workbookReader = PlugInLoader.GetPlugIn<IPlugInReader>(PlugInUUID.WORKBOOK_READER, new WorkbookReader());
            workbookReader.Init(ms, wb, readerOptions);
            workbookReader.Execute();

            ms = GetEntryStream("docProps/app.xml", zf);
            if (ms != null && ms.Length > 0) // If null/length == 0, no docProps/app.xml seems to be defined 
            {
                IPlugInReader metadataAppReader = PlugInLoader.GetPlugIn<IPlugInReader>(PlugInUUID.METADATA_APP_READER, new MetadataAppReader());
                metadataAppReader.Init(ms, wb, readerOptions);
                metadataAppReader.Execute();
            }
            ms = GetEntryStream("docProps/core.xml", zf);
            if (ms != null && ms.Length > 0) // If null/length == 0, no docProps/core.xml seems to be defined 
            {
                IPlugInReader metadataCoreReader = PlugInLoader.GetPlugIn<IPlugInReader>(PlugInUUID.METADATA_CORE_READER, new MetadataCoreReader());
                metadataCoreReader.Init(ms, wb, readerOptions);
                metadataCoreReader.Execute();
            }

            IPlugInReader relationships = PlugInLoader.GetPlugIn<IPlugInReader>(PlugInUUID.RELATIONSHIP_READER, new RelationshipReader());
            ms = GetEntryStream("xl/_rels/workbook.xml.rels", zf);
            relationships.Init(ms, wb, readerOptions);
            relationships.Execute();

            IWorksheetReader worksheetReader = PlugInLoader.GetPlugIn<IWorksheetReader>(PlugInUUID.WORKSHEET_READER, new WorksheetReader());
            worksheetReader.SharedStrings = sharedStringsReader.SharedStrings;
            List<WorksheetDefinition> workshetDefinitions = wb.AuxiliaryData.GetDataList<WorksheetDefinition>(PlugInUUID.WORKBOOK_READER, PlugInUUID.WORKSHEET_DEFINITION_ENTITY);
            List<Relationship> relationshipDefinitions = wb.AuxiliaryData.GetDataList<Relationship>(PlugInUUID.RELATIONSHIP_READER, PlugInUUID.RELATIONSHIP_ENTITY);
            foreach (WorksheetDefinition definition in workshetDefinitions)
            {
                Relationship relationship = relationshipDefinitions.SingleOrDefault(r => r.RID == definition.RelId);
                if (relationship == null)
                {
                    throw new IOException("There was an error while reading an XLSX file. The relationship target of the worksheet with the RelID " + definition.RelId + " was not found");
                }
                ms = GetEntryStream(relationship.Target, zf);
                worksheetReader.Init(ms, wb, readerOptions);
                worksheetReader.CurrentWorksheetID = definition.SheetID;
                worksheetReader.Execute();
            }
            if (wb.Worksheets.Count == 0)
            {
                throw new IOException("No worksheet was found in the workbook");
            }
            HandleQueuePlugIns(PlugInUUID.READER_APPENDING_QUEUE, zf, ref wb);
            wb.importInProgress = false; // Enables checks for runtime
            wb.AuxiliaryData.ClearTemporaryData(); // Remove temporary staging data
            this.Workbook = wb;
        }

        /// <summary>
        /// Gets the memory stream of the specified file in the archive (XLSX file)
        /// </summary>
        /// <param name="name">Name of the XML file within the XLSX file</param>
        /// <param name="archive">Zip file (XLSX)</param>
        /// <returns>MemoryStream object of the specified file</returns>
        private MemoryStream GetEntryStream(string name, ZipArchive archive)
        {
            MemoryStream stream = null;
            for (int i = 0; i < archive.Entries.Count; i++)
            {
                if (archive.Entries[i].FullName == name)
                {
                    MemoryStream ms = new MemoryStream();
                    archive.Entries[i].Open().CopyTo(ms);
                    ms.Position = 0;
                    stream = ms;
                    break;
                }
            }
            return stream;
        }

        /// <summary>
        /// Gets a map of all packed filenames that are matching the given prefix
        /// </summary>
        /// <param name="namePrefix">filename prefix</param>
        /// <param name="archive">Zip archive instance</param>
        /// <returns>Dictionary of filename, where the key is the extracted index of the filename</returns>
        private Dictionary<int, string> GetSequentialStreamNames(string namePrefix, ZipArchive archive)
        {
            Dictionary<int, string> files = new Dictionary<int, string>();
            int index = 1; // Assumption: There is no file that has the index 0 in its name
            MemoryStream ms = null;
            while (true)
            {
                string name = namePrefix + ParserUtils.ToString(index) + ".xml";
                ms = GetEntryStream(name, archive);
                if (ms != null)
                {
                    files.Add(index, name);
                }
                else
                {
                    break;
                }
                index++;
            }
            return files;
        }

        /// <summary>
        /// Method to handle queue plug-ins
        /// </summary>
        /// <param name="queueUuid">Queue UUID</param>
        /// <param name="zf">Zip archive</param>
        /// <param name="workbook">Workbook reference</param>
        private void HandleQueuePlugIns(string queueUuid, ZipArchive zf, ref Workbook workbook)
        {
            IPlugInReader queueReader = null;
            string lastUuid = null;
            do
            {
                string currentUuid;
                queueReader = PlugInLoader.GetNextQueuePlugIn<IPlugInReader>(queueUuid, lastUuid, out currentUuid);
                MemoryStream ms = null;
                if (queueReader != null)
                {
                    if (queueReader is IPlugInPackageReader)
                    {
                        string streamPartName = (queueReader as IPlugInPackageReader).StreamEntryName;
                        if (!string.IsNullOrEmpty(streamPartName))
                        {
                            ms = GetEntryStream(streamPartName, zf);
                            if (ms == null)
                            {
                                lastUuid = currentUuid;
                                continue; // Skip if the stream part name was defined but not found
                            }
                        }
                    }
                    queueReader.Init(ms, workbook, this.readerOptions); // stream may be null
                    queueReader.Execute();
                    lastUuid = currentUuid;
                }
                else
                {
                    lastUuid = null;
                }

            } while (queueReader != null);
        }
        #endregion
    }
}
