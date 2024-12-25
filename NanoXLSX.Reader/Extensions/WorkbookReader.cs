/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2024
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.IO;
using System.Threading.Tasks;
using NanoXLSX.Internal.Readers;
using NanoXLSX.Registry;

namespace NanoXLSX
{
    /// <summary>
    /// Class, providing functions to read Workbooks either by file or stream
    /// </summary>
    public static class WorkbookReader
    {
        static WorkbookReader()
        {
            PackageRegistry.Initialize();
        }

        /// <summary>
        /// Loads a workbook from a file
        /// </summary>
        /// <param name="filename">Filename of the workbook</param>
        /// <param name="options">Import options to override the data types of columns or cells. These options can be used to cope with wrong interpreted data, caused by irregular styles</param>
        /// <returns>Workbook object</returns>
        /// <exception cref="NanoXLSX.Shared.Exceptions.IOException">Throws IOException in case of an error</exception>
        public static Workbook Load(string filename, ReaderOptions options = null)
        {
            XlsxReader r = new XlsxReader(filename, options);
            r.Read();
            return r.GetWorkbook();
        }

        /// <summary>
        /// Loads a workbook from a stream
        /// </summary>
        /// <param name="stream">Stream containing the workbook</param>
        /// /// <param name="options">Import options to override the data types of columns or cells. These options can be used to cope with wrong interpreted data, caused by irregular styles</param>
        /// <returns>Workbook object</returns>
        /// <exception cref="NanoXLSX.Shared.Exceptions.IOException">Throws IOException in case of an error</exception>
        public static Workbook Load(Stream stream, ReaderOptions options = null)
        {
            XlsxReader r = new XlsxReader(stream, options);
            r.Read();
            return r.GetWorkbook();
        }

        /// <summary>
        /// Loads a workbook from a file asynchronously
        /// </summary>
        /// <param name="filename">Filename of the workbook</param>
        /// <param name="options">Import options to override the data types of columns or cells. These options can be used to cope with wrong interpreted data, caused by irregular styles</param>
        /// <returns>Workbook object</returns>
        /// <exception cref="IOException">Throws IOException in case of an error</exception>
        public static async Task<Workbook> LoadAsync(string filename, ReaderOptions options = null)
        {
            XlsxReader r = new XlsxReader(filename, options);
            await r.ReadAsync();
            return r.GetWorkbook();
        }

        /// <summary>
        /// Loads a workbook from a stream asynchronously
        /// </summary>
        /// <param name="stream">Stream containing the workbook</param>
        /// /// <param name="options">Import options to override the data types of columns or cells. These options can be used to cope with wrong interpreted data, caused by irregular styles</param>
        /// <returns>Workbook object</returns>
        /// <exception cref="NanoXLSX.Shared.Exceptions.IOException">Throws IOException in case of an error</exception>
        public static async Task<Workbook> LoadAsync(Stream stream, ReaderOptions options = null)
        {
            XlsxReader r = new XlsxReader(stream, options);
            await r.ReadAsync();
            return r.GetWorkbook();
        }
    }
}
