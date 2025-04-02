/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Interfaces.Writer;
using NanoXLSX.Registry;
using NanoXLSX.Utils.Xml;

namespace NanoXLSX.Internal.Writers
{
    /// <summary>
    /// Class for the handling of writer inline plug-ins
    /// </summary>
    internal static class WriterPlugInHandler
    {
        /// <summary>
        /// Method to handle inline queue plug-ins of a specific writer plug-in
        /// </summary>
        /// <param name="workbook">Workbook reference</param>
        /// <param name="rootElement">Reference to te root element of the base writer plug-in</param>
        /// <param name="queueUuid">UUID of the inline plug-in</param>
        /// <returns>XML element instance. If no plug-ins were processes, the root element is passed back unaltered</returns>
        internal static void HandleInlineQueuePlugins(ref XmlElement rootElement, Workbook workbook, string queueUuid)
        {
            IInlinePlugInWriter queueWriter = null;
            string lastUuid = null;
            do
            {
                string currentUuid;
                queueWriter = PlugInLoader.GetNextQueuePlugIn<IInlinePlugInWriter>(queueUuid, lastUuid, out currentUuid);
                if (queueWriter != null)
                {
                    queueWriter.Init(ref rootElement, workbook);
                    queueWriter.Execute();
                    lastUuid = currentUuid;
                }
                else
                {
                    lastUuid = null;
                }

            } while (queueWriter != null);
        }
    }
}
