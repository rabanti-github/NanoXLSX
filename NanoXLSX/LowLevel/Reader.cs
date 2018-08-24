/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2018
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NanoXLSX.Exception;
using IOException = NanoXLSX.Exception.IOException;

namespace NanoXLSX.LowLevel
{
    public class Reader
    {
#region privateFields
        private String filePath;
        private Stream inputStream;
        private Dictionary<int, WorksheetReader> worksheets;
        private MemoryStream memoryStream;
        private WorkbookReader workbook;
#endregion

#region constructors
        /// <summary>
        /// Constructor with file path as parameter
        /// </summary>
        /// <param name="path">File path of the xlsx file to load</param>
        public Reader(String path)
        {
            this.filePath = path;
            this.worksheets = new Dictionary<int, WorksheetReader>();
        }

        /// <summary>
        /// Constructor with stream as parameter
        /// </summary>
        /// <param name="stream">Stream of the xlsx file to load</param>
        public Reader(Stream stream)
        {
            this.worksheets = new Dictionary<int, WorksheetReader>();
            this.inputStream = stream;
        }
        #endregion

        #region methods


        /// <summary>
        /// Reads the XLSX file from a file path or a file stream
        /// </summary>
        /// <exception cref="IOException">
        /// Throws IOException in case of an error
        /// </exception>
        public void Read()
    {
        try
        {
            using (memoryStream = new MemoryStream())
            {

                ZipArchive zf;
                if (inputStream == null || string.IsNullOrEmpty(filePath) == false)
                {
                    using (FileStream fs = new FileStream(filePath, FileMode.Open))
                    {
                        fs.CopyTo(memoryStream);
                    }
                }
                else if (inputStream != null)
                {
                    using (inputStream)
                    {
                        inputStream.CopyTo(memoryStream);
                    }
                }
                else
                {
                    throw new IOException("LoadException", "No valid stream or file path was provided to open");
                }

                memoryStream.Position = 0;
                zf = new ZipArchive(memoryStream, ZipArchiveMode.Read);
                MemoryStream ms;
                SharedStringsReader sharedStrings = new SharedStringsReader();
                ms = GetEntryStream("xl/sharedStrings.xml", zf);
                sharedStrings.Read(ms);

                this.workbook = new WorkbookReader();
                ms = GetEntryStream("xl/workbook.xml", zf);
                this.workbook.Read(ms);

                int worksheetIndex = 1;
                string name, nameTemplate;
                WorksheetReader wr;
                nameTemplate = "sheet" + worksheetIndex.ToString(CultureInfo.InvariantCulture) + ".xml";
                name = "xl/worksheets/" + nameTemplate;
                for (int i = 0; i < this.workbook.WorksheetDefinitions.Count; i++)
                {
                    ms = GetEntryStream(name, zf);
                    wr = new WorksheetReader(sharedStrings, nameTemplate, worksheetIndex);
                    wr.Read(ms);
                    this.worksheets.Add(worksheetIndex - 1, wr);
                    worksheetIndex++;
                    nameTemplate = "sheet" + worksheetIndex.ToString(CultureInfo.InvariantCulture) + ".xml";
                    name = "xl/worksheets/" + nameTemplate;
                }

            }
        }
        catch (System.Exception ex)
        {
            throw new IOException("LoadException", "There was an error while reading an XLSX file. Please see the inner exception:", ex);
        }
    }

        /// <summary>
        /// Resolves the workbook with all worksheets from the loaded file
        /// </summary>
        /// <returns>Workbook object</returns>
        public Workbook GetWorkbook()
        {
            Workbook wb = new Workbook(false);
            WorksheetReader wr;
            Worksheet ws;
            for (int i = 0; i < this.worksheets.Count; i++)
            {
                wr = this.worksheets[i];
                ws = new Worksheet(this.workbook.WorksheetDefinitions[i + 1], i + 1, wb);
                foreach (KeyValuePair<string, Cell> cell in wr.Data)
                {
                    ws.AddCell(cell.Value, cell.Key);
                }
                wb.AddWorksheet(ws);
            }
            return wb;
        }

        /// <summary>
        /// Gets the memory stream of the specified file in the archive (XLSX file)
        /// </summary>
        /// <param name="name">Name of the XML file within the XLSX file</param>
        /// <param name="archive">Zip file (XLSX)</param>
        /// <returns>MemoryStream object of the specified file</returns>
        private MemoryStream GetEntryStream(string name, ZipArchive archive)
        {
            for (int i = 0; i < archive.Entries.Count; i++)
            {
                if (archive.Entries[i].FullName == name)
                {
                    MemoryStream ms = new MemoryStream();
                    archive.Entries[i].Open().CopyTo(ms);
                    ms.Position = 0;
                    return ms;
                }
            }

            return new MemoryStream();
        }

#endregion


    }
}
