/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Diagnostics.CodeAnalysis;

namespace NanoXLSX.Internal.Readers
{
    /// <summary>
    /// Class to represent a workbook relation (do not use)
    /// </summary>
    [Obsolete("Will be removed with next major version")]
    public class Relationship
    {
        /// <summary>
        /// ID of the relation
        /// </summary>
        [ExcludeFromCodeCoverage]
        public string RID { get; set; }
        /// <summary>
        /// Type of the relation
        /// </summary>
        [ExcludeFromCodeCoverage]
        public string Type { get; set; }
        /// <summary>
        /// Target of the relation
        /// </summary>
        [ExcludeFromCodeCoverage]
        public string Target { get; set; }
    }
}
