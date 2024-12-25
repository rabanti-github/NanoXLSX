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
        /** <example> 
         This is an extension method, applied to the <see cref="NanoXLSX.Workbook">Workbook</see> class. An example usage is:
        <code>
            Workbook wb = new Workbook("workboox.xlsx", "worksheet1");
            // do some operations with wb, like adding cells
            wb.Save(); // Will save the workbok as 'workbook.xlsx' in the execution path
         </code> 
         </example> **/
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
        /** <example> 
         This is an extension method, applied to the <see cref="NanoXLSX.Workbook">Workbook</see> class. An example usage is:
        <code>
            private async Task CreateXlsx()
            {
                Workbook wb = new Workbook("workboox.xlsx", "worksheet1");
                // do some operations with wb, like adding cells
                await wb.SaveAsync(); // Will save the workbok as 'workbook.xlsx' in the execution path
            }
         </code> 
         </example> **/
        public static async Task SaveAsync(this Workbook workbook)
        {
            XlsxWriter l = new XlsxWriter(workbook);
            await l.SaveAsync();
        }

        /// <summary>
        /// Saves the workbook with the defined name
        /// </summary>
        /// <param name="filename">filename of the saved workbook</param>
        /// <param name="workbook">Workbook reference</param>
        /// <exception cref="NanoXLSX.Shared.Exceptions.IOException">Throws IOException in case of an error</exception>
        /** <example> 
         This is an extension method, applied to the <see cref="NanoXLSX.Workbook">Workbook</see> class. An example usage is:
        <code>
            Workbook wb = new Workbook("worksheet1");
            // do some operations with wb, like adding cells
            wb.SaveAs("workboox.xlsx"); // Will save the workbok as 'workbook.xlsx' in the execution path
         </code> 
         </example> **/
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
        /// <param name="workbook">Workbook reference</param>
        /// <returns>Task object (void)</returns>
        /// <exception cref="NanoXLSX.Shared.Exceptions.IOException">May throw an IOException in case of an error. The asynchronous operation may hide the exception.</exception>
        /** <example> 
         This is an extension method, applied to the <see cref="NanoXLSX.Workbook">Workbook</see> class. An example usage is:
        <code>
            private async Task CreateXlsx()
            {
                Workbook wb = new Workbook("worksheet1");
                // do some operations with wb, like adding cells
                await wb.SaveAsAsync("workboox.xlsx"); // Will save the workbok as 'workbook.xlsx' in the execution path
            }
         </code> 
         </example> **/
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
        /// <param name="workbook">Workbook reference</param>
        /// <exception cref="NanoXLSX.Shared.Exceptions.IOException">Throws IOException in case of an error</exception>
                /** <example> 
         This is an extension method, applied to the <see cref="NanoXLSX.Workbook">Workbook</see> class. An example usage is:
        <code>
            Workbook wb = new Workbook("worksheet1");
            // do some operations with wb, like adding cells
            using(FileStream fs = new FileStream("workbook.xslx", FileMode.Create))
            {
                wb.SaveAsStream(fs); // Will save the workbok as 'workbook.xlsx' using a FileStream
            }
         </code>
         The stream can also be kept open:
        <code>
            Workbook wb = new Workbook("worksheet1");
            // do some operations with wb, like adding cells
            using(MemoryStream ms = new MemoryStream())
            {
                wb.SaveAsStream(ms, true); // Will save the workbok into a MemoryStream
                ms.Position = 0; // Rewind the stream
                // use ms to do copy or save actions
            }
         </code>
         </example> **/
        public static void SaveAsStream(this Workbook workbook, Stream stream, bool leaveOpen = false)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Position = 0;
            }
            XlsxWriter l = new XlsxWriter(workbook);
            l.SaveAsStream(stream, leaveOpen);
        }

        /// <summary>
        /// Save the workbook to a writable stream asynchronous.
        /// </summary>
        /// <param name="stream">>Writable stream</param>
        /// <param name="leaveOpen">Optional parameter to keep the stream open after writing (used for MemoryStreams; default is false)</param>
        /// <param name="workbook">Workbook reference</param>
        /// <returns>Task object (void)</returns>
        /// <exception cref="NanoXLSX.Shared.Exceptions.IOException">Throws IOException in case of an error. The asynchronous operation may hide the exception.</exception>
                /** <example> 
         This is an extension method, applied to the <see cref="NanoXLSX.Workbook">Workbook</see> class. An example usage is:
        <code>
            private async Task CreateXlsx()
            {
                Workbook wb = new Workbook("worksheet1");
                // do some operations with wb, like adding cells
                using(FileStream fs = new FileStream("workbook.xslx", FileMode.Create))
                {
                    await wb.SaveAsStreamAsync(fs); // Will save the workbok as 'workbook.xlsx' using a FileStream
                }
            }
         </code> 
         The stream can also be kept open:
         <code>
            private async Task CreateXlsx()
            {
                Workbook wb = new Workbook("worksheet1");
                // do some operations with wb, like adding cells
                using(MemoryStream ms = new MemoryStream())
                {
                    await wb.SaveAsStreamAsync(ms, true); // Will save the workbok into a MemoryStream
                    ms.Position = 0; // Rewind the stream
                    // use ms to do copy or save actions
                }
            }
         </code> 
         </example> **/
        public static async Task SaveAsStreamAsync(this Workbook workbook, Stream stream, bool leaveOpen = false)
        {
            XlsxWriter l = new XlsxWriter(workbook);
            await l.SaveAsStreamAsync(stream, leaveOpen);
        }
    }
}
