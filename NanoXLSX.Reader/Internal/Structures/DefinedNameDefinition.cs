/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

namespace NanoXLSX.Internal.Readers
{
    /// <summary>
    /// Class for defined-name meta-data on import. The information is captured by <see cref="WorkbookReader"/> and
    /// finalized into <see cref="DefinedName"/> instances by <see cref="WorksheetReader"/> once all worksheets are bound.
    /// </summary>
    internal class DefinedNameDefinition
    {
        /// <summary>
        /// Name of the defined name
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// Verbatim reference text
        /// </summary>
        public string Reference { get; set; }
        /// <summary>
        /// Optional local sheet index (zero-based, into the visible worksheet list). Null for workbook scope.
        /// </summary>
        public int? LocalSheetId { get; set; }
        /// <summary>
        /// Optional comment
        /// </summary>
        public string Comment { get; set; }
    }
}
