/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

namespace NanoXLSX.Interfaces.Reader
{
    /// <summary>
    /// Interface for readers that process a document identified by an OOXML relationship type.
    /// </summary>
    internal interface IDocumentReader : IPluginBaseReader
    {
        /// <summary>
        /// Gets the complete, case-sensitive relationship type URI handled by the reader.
        /// </summary>
        string DocumentType { get; }
    }
}
