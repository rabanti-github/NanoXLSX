using System;
using System.ComponentModel;
using System.IO;
using System.Threading.Tasks;
using NanoXLSX.Internal.Writers;

namespace NanoXLSX
{
    [EditorBrowsable(EditorBrowsableState.Never)]
    public static class WorkbookExtension
    {
        /// <summary>
        /// Saves the workbook
        /// </summary>
        /// <exception cref="NanoXLSX.Shared.Exceptions.IOException">Throws IOException in case of an error</exception>
        /// <exception cref="NanoXLSX.Shared.Exceptions.RangeException">Throws a RangeException if the start or end address of a handled cell range was out of range</exception>
        /// <exception cref="NanoXLSX.Shared.Exceptions.FormatException">Throws a FormatException if a handled date cannot be translated to (Excel internal) OADate</exception>
        public static void Save(this Workbook workbook)
        {
            XlsxWriter l = new XlsxWriter(workbook);
            l.Save();
        }

        /// <summary>
        /// Saves the workbook asynchronous.
        /// </summary>
        /// <returns>Task object (void)</returns>
        /// <exception cref="NanoXLSX.Shared.Exceptions.IOException">May throw an IOException in case of an error. The asynchronous operation may hide the exception.</exception>
        /// <exception cref="NanoXLSX.Shared.Exceptions.RangeException">May throw a RangeException if the start or end address of a handled cell range was out of range. The asynchronous operation may hide the exception.</exception>
        /// <exception cref="NanoXLSX.Shared.Exceptions.FormatException">May throw a FormatException if a handled date cannot be translated to (Excel internal) OADate. The asynchronous operation may hide the exception.</exception>
        public static async Task SaveAsync(this Workbook workbook)
        {
            XlsxWriter l = new XlsxWriter(workbook);
            await l.SaveAsync();
        }

        /// <summary>
        /// Saves the workbook with the defined name
        /// </summary>
        /// <param name="filename">filename of the saved workbook</param>
        /// <exception cref="NanoXLSX.Shared.Exceptions.IOException">Throws IOException in case of an error</exception>
        /// <exception cref="NanoXLSX.Shared.Exceptions.RangeException">Throws a RangeException if the start or end address of a handled cell range was out of range</exception>
        /// <exception cref="NanoXLSX.Shared.Exceptions.FormatException">Throws a FormatException if a handled date cannot be translated to (Excel internal) OADate</exception>
        public static void SaveAs(this Workbook workbook, string filename)
        {
            string backup = filename;
            workbook.Filename = filename;
            XlsxWriter l = new XlsxWriter(workbook);
            l.Save();
            workbook.Filename = backup;
        }

        /// <summary>
        /// Saves the workbook with the defined name asynchronous.
        /// </summary>
        /// <param name="fileName">filename of the saved workbook</param>
        /// <returns>Task object (void)</returns>
        /// <exception cref="NanoXLSX.Shared.Exceptions.IOException">May throw an IOException in case of an error. The asynchronous operation may hide the exception.</exception>
        /// <exception cref="NanoXLSX.Shared.Exceptions.RangeException">May throw a RangeException if the start or end address of a handled cell range was out of range. The asynchronous operation may hide the exception.</exception>
        /// <exception cref="NanoXLSX.Shared.Exceptions.FormatException">May throw a FormatException if a handled date cannot be translated to (Excel internal) OADate. The asynchronous operation may hide the exception.</exception>
        public static async Task SaveAsAsync(this Workbook workbook, string fileName)
        {
            string backup = fileName;
            workbook.Filename = fileName;
            XlsxWriter l = new XlsxWriter(workbook);
            await l.SaveAsync();
            workbook.Filename = backup;
        }

        /// <summary>
        /// Save the workbook to a writable stream
        /// </summary>
        /// <param name="stream">Writable stream</param>
        /// <param name="leaveOpen">Optional parameter to keep the stream open after writing (used for MemoryStreams; default is false)</param>
        /// <exception cref="NanoXLSX.Shared.Exceptions.IOException">Throws IOException in case of an error</exception>
        /// <exception cref="NanoXLSX.Shared.Exceptions.RangeException">Throws a RangeException if the start or end address of a handled cell range was out of range</exception>
        /// <exception cref="FormatException">Throws a FormatException if a handled date cannot be translated to (Excel internal) OADate</exception>
        public static void SaveAsStream(this Workbook workbook, Stream stream, bool leaveOpen = false)
        {
            XlsxWriter l = new XlsxWriter(workbook);
            l.SaveAsStream(stream, leaveOpen);
        }

        /// <summary>
        /// Save the workbook to a writable stream asynchronous.
        /// </summary>
        /// <param name="stream">>Writable stream</param>
        /// <param name="leaveOpen">Optional parameter to keep the stream open after writing (used for MemoryStreams; default is false)</param>
        /// <returns>Task object (void)</returns>
        /// <exception cref="NanoXLSX.Shared.Exceptions.IOException">Throws IOException in case of an error. The asynchronous operation may hide the exception.</exception>
        /// <exception cref="NanoXLSX.Shared.Exceptions.RangeException">May throw a RangeException if the start or end address of a handled cell range was out of range. The asynchronous operation may hide the exception.</exception>
        /// <exception cref="FormatException">May throw a FormatException if a handled date cannot be translated to (Excel internal) OADate. The asynchronous operation may hide the exception.</exception>
        public static async Task SaveAsStreamAsync(this Workbook workbook, Stream stream, bool leaveOpen = false)
        {
            XlsxWriter l = new XlsxWriter(workbook);
            await l.SaveAsStreamAsync(stream, leaveOpen);
        }
    }
}
