/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.IO;
using NanoXLSX.Interfaces;
using NanoXLSX.Interfaces.Reader;
using NanoXLSX.Registry;

namespace NanoXLSX.Internal.Readers
{
    /// <summary>
    /// Class for the handling of reader in-line plug-ins
    /// </summary>
    internal static class ReaderPlugInHandler
    {
        /// <summary>
        /// Method to handle in-line queue plug-ins of a specific reader plug-in
        /// </summary>
        /// <param name="stream">MemoryStream to be read</param>
        /// <param name="workbook">Workbook reference</param>
        /// <param name="queueUuid">UUID of the in-line plug-in</param>
        /// <param name="readerOptions">Reader options</param>
        /// <param name="index">Optional index, e.g. for worksheet identification</param>
        internal static void HandleInlineQueuePlugins(MemoryStream stream, Workbook workbook, string queueUuid, IOptions readerOptions, int? index)
        {
            IPluginInlineReader queueReader = null;
            string lastUuid = null;
            do
            {
                string currentUuid;
                queueReader = PlugInLoader.GetNextQueuePlugIn<IPluginInlineReader>(queueUuid, lastUuid, out currentUuid);
                if (queueReader != null)
                {
                    stream.Position = 0;
                    queueReader.Init(stream, workbook, readerOptions, index);
                    queueReader.Execute();
                    lastUuid = currentUuid;
                }
                else
                {
                    lastUuid = null;
                }

            } while (queueReader != null);
        }
    }
}
