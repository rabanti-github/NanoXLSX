/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */


using System.IO;
using NanoXLSX.Interfaces.Reader;
using NanoXLSX.Registry;

namespace NanoXLSX.Internal.Readers
{
    /// <summary>
    /// Class for the handling of reader inline plug-ins
    /// </summary>
    internal static class RederPlugInHandler
    {
        internal static void HandleInlineQueuePlugins(ref MemoryStream stream, Workbook workbook, string queueUuid)
        {
            IInlinePlugInReader queueReader = null;
            string lastUuid = null;
            do
            {
                string currentUuid;
                queueReader = PlugInLoader.GetNextQueuePlugIn<IInlinePlugInReader>(queueUuid, lastUuid, out currentUuid);
                if (queueReader != null)
                {
                    stream.Position = 0;
                    queueReader.Init(ref stream, workbook);
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
