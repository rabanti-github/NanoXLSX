/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Diagnostics.CodeAnalysis;
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
        /// <param name="stream">Stream to be read. May be non-seekable; in that case the method buffers it lazily into a MemoryStream when at least one inline plug-in is registered.</param>
        /// <param name="workbook">Workbook reference</param>
        /// <param name="queueUuid">UUID of the in-line plug-in</param>
        /// <param name="readerOptions">Reader options</param>
        /// <param name="index">Optional index, e.g. for worksheet identification</param>
        [ExcludeFromCodeCoverage] // No testable logic, only plug-in handling
        internal static void HandleInlineQueuePlugins(Stream stream, Workbook workbook, string queueUuid, IOptions readerOptions, int? index)
        {
            if (stream == null)
            {
                return;
            }
            // Inline plug-ins reset Position to 0 to re-read the part. If the caller passed a
            // non-seekable stream (e.g. a deflate stream from a ZIP entry), the contract can only
            // be honored by buffering once. Skip the buffer when no inline plug-ins exist.
            Stream working = stream;
            MemoryStream owned = null;
            if (!stream.CanSeek && PlugInLoader.HasQueuePlugins(queueUuid))
            {
                owned = new MemoryStream();
                stream.CopyTo(owned);
                owned.Position = 0;
                working = owned;
            }
            try
            {
                IPluginInlineReader queueReader = null;
                string lastUuid = null;
                do
                {
                    string currentUuid;
                    queueReader = PlugInLoader.GetNextQueuePlugIn<IPluginInlineReader>(queueUuid, lastUuid, out currentUuid);
                    if (queueReader != null)
                    {
                        if (working.CanSeek)
                        {
                            working.Position = 0;
                        }
                        queueReader.Init(working, workbook, readerOptions, index);
                        queueReader.Execute();
                        lastUuid = currentUuid;
                    }
                    else
                    {
                        lastUuid = null;
                    }

                } while (queueReader != null);
            }
            finally
            {
                owned?.Dispose();
            }
        }
    }
}
