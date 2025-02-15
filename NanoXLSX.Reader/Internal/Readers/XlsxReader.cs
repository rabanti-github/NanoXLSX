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
        private Dictionary<int, WorksheetReader> worksheets;
        private MemoryStream memoryStream;
        private WorkbookReader workbook;
        private MetadataCoreReader metadataCoreReader;
        private MetadataAppReader metadataAppReader;
        private ThemeReader themeReader;
        private ReaderOptions readerOptions;
        private StyleReaderContainer styleReaderContainer;
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
            worksheets = new Dictionary<int, WorksheetReader>();
        }

        /// <summary>
        /// Constructor with stream as parameter
        /// </summary>
        /// <param name="options">Reader options to override the automatic approach of the reader. <see cref="ReaderOptions"/> for information about Reader options.</param>
        /// <param name="stream">Stream of the XLSX file to load</param>
        public XlsxReader(Stream stream, ReaderOptions options = null)
        {
            readerOptions = options;
            worksheets = new Dictionary<int, WorksheetReader>();
            inputStream = stream;
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
        /// Resolves the workbook with all worksheets from the loaded file
        /// </summary>
        /// <returns>Workbook object</returns>
        public Workbook GetWorkbook()
        {
            Workbook wb = new Workbook(false);
            wb.importInProgress = true;
            Worksheet ws;
            foreach (KeyValuePair<int, WorksheetReader> reader in worksheets)
            {
                WorkbookReader.WorksheetDefinition definition = workbook.WorksheetDefinitions[reader.Key];
                ws = new Worksheet(definition.WorksheetName, definition.SheetID, wb);
                ws.Hidden = definition.Hidden;
                ws.ViewType = reader.Value.ViewType;
                ws.ShowGridLines = reader.Value.ShowGridLines;
                ws.ShowRowColumnHeaders = reader.Value.ShowRowColHeaders;
                ws.ShowRuler = reader.Value.ShowRuler;
                ws.ZoomFactor = reader.Value.CurrentZoomScale;
                foreach (KeyValuePair<Worksheet.SheetViewType, int> zoomFactor in reader.Value.ZoomFactors)
                {
                    ws.SetZoomFactor(zoomFactor.Key, zoomFactor.Value);
                }
                if (reader.Value.AutoFilterRange.HasValue)
                {
                    ws.SetAutoFilter(reader.Value.AutoFilterRange.Value.StartAddress.Column, reader.Value.AutoFilterRange.Value.EndAddress.Column);
                }
                if (reader.Value.DefaultColumnWidth.HasValue)
                {
                    ws.DefaultColumnWidth = reader.Value.DefaultColumnWidth.Value;
                }
                if (reader.Value.DefaultRowHeight.HasValue)
                {
                    ws.DefaultRowHeight = reader.Value.DefaultRowHeight.Value;
                }
                if (reader.Value.SelectedCells.Count > 0)
                {
                    foreach (Range range in reader.Value.SelectedCells)
                    {
                        ws.AddSelectedCells(range);
                    }
                }
                foreach (Range range in reader.Value.MergedCells)
                {
                    ws.MergeCells(range);
                }
                foreach (KeyValuePair<Worksheet.SheetProtectionValue, int> sheetProtection in reader.Value.WorksheetProtection)
                {
                    ws.SheetProtectionValues.Add(sheetProtection.Key);
                }
                if (reader.Value.WorksheetProtection.Count > 0)
                {
                    ws.UseSheetProtection = true;
                }
                if (reader.Value.PasswordReader.PasswordIsSet())
                {
                    if (reader.Value.PasswordReader is LegacyPasswordReader && (reader.Value.PasswordReader as LegacyPasswordReader).ContemporaryAlgorithmDetected && !readerOptions.IgnoreNotSupportedPasswordAlgorithms)
                    {
                        throw new NotSupportedContentException("A not supported, contemporary password algorithm for the worksheet protection was detected. Check possible packages to add support to NanoXLSX, or ignore this error by a reader option");
                    }
                    ws.SheetProtectionPassword.CopyFrom(reader.Value.PasswordReader);
                }
               // if (!string.IsNullOrEmpty(reader.Value.WorksheetProtectionHash))
               // {
              //      ws.SheetProtectionPasswordHash = reader.Value.WorksheetProtectionHash;
              //  }
                foreach (KeyValuePair<int, WorksheetReader.RowDefinition> row in reader.Value.Rows)
                {
                    if (row.Value.Hidden)
                    {
                        ws.AddHiddenRow(row.Key);
                    }
                    if (row.Value.Height.HasValue)
                    {
                        ws.SetRowHeight(row.Key, row.Value.Height.Value);
                    }
                }
                foreach (Column column in reader.Value.Columns)
                {
                    if (column.Width != Worksheet.DEFAULT_COLUMN_WIDTH)
                    {
                        ws.SetColumnWidth(column.ColumnAddress, column.Width);
                    }
                    if (column.IsHidden)
                    {
                        ws.AddHiddenColumn(column.Number);
                    }
                    if (column.DefaultColumnStyle != null)
                    {
                        ws.SetColumnDefaultStyle(column.ColumnAddress, column.DefaultColumnStyle);
                    }
                }
                foreach (KeyValuePair<string, Cell> cell in reader.Value.Data)
                {
                    if (reader.Value.StyleAssignment.ContainsKey(cell.Key))
                    {
                        Style style = styleReaderContainer.GetStyle(reader.Value.StyleAssignment[cell.Key]);
                        if (style != null)
                        {
                            cell.Value.SetStyle(style);
                        }
                    }
                    ws.AddCell(cell.Value, cell.Key);
                }
                if (reader.Value.PaneSplitValue != null)
                {
                    WorksheetReader.PaneDefinition pane = reader.Value.PaneSplitValue;
                    if (pane.FrozenState)
                    {
                        if (pane.YSplitDefined && !pane.XSplitDefined)
                        {
                            ws.SetHorizontalSplit(pane.PaneSplitRowIndex.Value, pane.FrozenState, pane.TopLeftCell, pane.ActivePane);
                        }
                        if (!pane.YSplitDefined && pane.XSplitDefined)
                        {
                            ws.SetVerticalSplit(pane.PaneSplitColumnIndex.Value, pane.FrozenState, pane.TopLeftCell, pane.ActivePane);
                        }
                        else if (pane.YSplitDefined && pane.XSplitDefined)
                        {
                            ws.SetSplit(pane.PaneSplitColumnIndex.Value, pane.PaneSplitRowIndex.Value, pane.FrozenState, pane.TopLeftCell, pane.ActivePane);
                        }
                    }
                    else
                    {
                        if (pane.YSplitDefined && !pane.XSplitDefined)
                        {
                            ws.SetHorizontalSplit(pane.PaneSplitHeight.Value, pane.TopLeftCell, pane.ActivePane);
                        }
                        if (!pane.YSplitDefined && pane.XSplitDefined)
                        {
                            ws.SetVerticalSplit(pane.PaneSplitWidth.Value, pane.TopLeftCell, pane.ActivePane);
                        }
                        else if (pane.YSplitDefined && pane.XSplitDefined)
                        {
                            ws.SetSplit(pane.PaneSplitWidth, pane.PaneSplitHeight, pane.TopLeftCell, pane.ActivePane);
                        }
                    }
                }
                wb.AddWorksheet(ws);
            }
            if (styleReaderContainer.GetMruColors().Count > 0)
            {
                foreach (string color in styleReaderContainer.GetMruColors())
                {
                    wb.AddMruColor(color);
                }
            }
            wb.Hidden = workbook.Hidden;
            wb.SetSelectedWorksheet(workbook.SelectedWorksheet);
            if (workbook.Protected)
            {
                wb.SetWorkbookProtection(workbook.Protected, workbook.LockWindows, workbook.LockStructure, null);
                if (workbook.PasswordReader is LegacyPasswordReader && (workbook.PasswordReader as LegacyPasswordReader).ContemporaryAlgorithmDetected && !readerOptions.IgnoreNotSupportedPasswordAlgorithms)
                {
                    throw new NotSupportedContentException("A not supported, contemporary password algorithm for the workbook protection was detected. Check possible packages to add support to NanoXLSX, or ignore this error by a reader option");
                }
                wb.WorkbookProtectionPassword.CopyFrom(workbook.PasswordReader);
               // wb.WorkbookProtectionPasswordHash = workbook.PasswordHash;
            }
            wb.WorkbookMetadata.Application = metadataAppReader.Application;
            wb.WorkbookMetadata.ApplicationVersion = metadataAppReader.ApplicationVersion;
            wb.WorkbookMetadata.Company = metadataAppReader.Company;
            wb.WorkbookMetadata.HyperlinkBase = metadataAppReader.HyperlinkBase;
            wb.WorkbookMetadata.Manager = metadataAppReader.Manager;

            wb.WorkbookMetadata.Keywords = metadataCoreReader.Keywords;
            wb.WorkbookMetadata.Subject = metadataCoreReader.Subject;
            wb.WorkbookMetadata.Title = metadataCoreReader.Title;
            wb.WorkbookMetadata.Creator = metadataCoreReader.Creator;
            wb.WorkbookMetadata.Category = metadataCoreReader.Category;
            wb.WorkbookMetadata.ContentStatus = metadataCoreReader.ContentStatus;
            wb.WorkbookMetadata.Description = metadataCoreReader.Description;

            if (themeReader != null)
            {
                wb.WorkbookTheme = themeReader.CurrentTheme;
            }
            wb.importInProgress = false;
            return wb;
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

            SharedStringsReader sharedStrings = new SharedStringsReader(readerOptions);
            ms = GetEntryStream("xl/sharedStrings.xml", zf);
            if (ms != null && ms.Length > 0) // If length == 0, no shared strings are defined (no text in file)
            {
                sharedStrings.Read(ms);
            }
            Dictionary<int, string> themeStreamNames = GetSequentialStreamNames("xl/theme/theme", zf);
            if (themeStreamNames.Count > 0)
            {
                // There is not really a definition whether multiple themes can be managed in one workbook.
                // the suffix number (e.g. theme1) indicates it. However, no examples were found and therefore
                // (currently) only the first occurring theme will be read  
                foreach (KeyValuePair<int, string> streamName in themeStreamNames)
                {
                    themeReader = new ThemeReader();
                    ms = GetEntryStream(streamName.Value, zf);
                    themeReader.Read(ms, streamName.Key);
                    break;
                }
            }
            StyleRepository.Instance.ImportInProgress = true;
            StyleReader styleReader = new StyleReader();
            ms = GetEntryStream("xl/styles.xml", zf);
            styleReader.Read(ms);
            styleReaderContainer = styleReader.StyleReaderContainer;
            StyleRepository.Instance.ImportInProgress = false;

            workbook = new WorkbookReader();
            ms = GetEntryStream("xl/workbook.xml", zf);
            workbook.Read(ms);

            metadataAppReader = new MetadataAppReader();
            ms = GetEntryStream("docProps/app.xml", zf);
            if (ms != null && ms.Length > 0) // If null/length == 0, no docProps/app.xml seems to be defined 
            {
                metadataAppReader.Read(ms);
            }
            metadataCoreReader = new MetadataCoreReader();
            ms = GetEntryStream("docProps/core.xml", zf);
            if (ms != null && ms.Length > 0) // If null/length == 0, no docProps/core.xml seems to be defined 
            {
                metadataCoreReader.Read(ms);
            }

            RelationshipReader relationships = new RelationshipReader();
            relationships.Read(GetEntryStream("xl/_rels/workbook.xml.rels", zf));

            WorksheetReader wr;
            foreach (KeyValuePair<int, WorkbookReader.WorksheetDefinition> definition in workbook.WorksheetDefinitions)
            {
                RelationshipReader.Relationship relationship = relationships.Relationships.SingleOrDefault(r => r.Id == definition.Value.RelId);
                if (relationship == null)
                {
                    throw new IOException("There was an error while reading an XLSX file. The relationship target of the worksheet with the RelID " + definition.Value.RelId + " was not found");
                }
                ms = GetEntryStream(relationship.Target, zf);
                wr = new WorksheetReader(sharedStrings, styleReaderContainer, readerOptions);
                wr.Read(ms);
                worksheets.Add(definition.Key, wr);
            }
            if (this.worksheets.Count == 0)
            {
                throw new IOException("No worksheet was found in the workbook");
            }
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

        #endregion

    }
}
