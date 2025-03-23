/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.IO;

namespace NanoXLSX.Interfaces.Reader
{
    /// <summary>
    /// Interface, used by XML reader classes 
    /// </summary>
    public interface IPluginReader
    {
        /// <summary>
        /// Interface function to read an XML file within an XLSX file
        /// </summary>
        /// <param name="stream">Stream of the XML file</param>
        void Read(MemoryStream stream);
        /// <summary>
        /// Method that is called before the <see cref="Read(MemoryStream)"/> method is executed
        /// </summary>
        /// <param name="stream">Stream of the XML file. The stream must be reset in this method at the end, if any stream operation was performed</param>
        void PreRead(MemoryStream stream);
        /// <summary>
        /// Method that is called after the <see cref="Read(MemoryStream)"/> method is executed
        /// </summary>
        /// <param name="stream">Stream of the XML file. The stream must be reset in this method before any stream operation is performed</param>
        void PostRead(MemoryStream stream);
    }
}
