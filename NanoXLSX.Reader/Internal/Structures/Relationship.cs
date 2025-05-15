/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;

namespace NanoXLSX.Internal.Readers
{
        /// <summary>
        /// Class to represent a workbook relation
        /// </summary>
        public class Relationship
        {
            /// <summary>
            /// ID of the relation
            /// </summary>
            public string RID { get; set; }
            /// <summary>
            /// Type of the relation
            /// </summary>
            public string Type { get; set; }
            /// <summary>
            /// Target of the relation
            /// </summary>
            public string Target { get; set; }

        /// <summary>
        /// Gets the numeric (1-based) ID of the relationship 
        /// </summary>
        /// <returns>1-based ID</returns>
        /// /remark <remarks>There is no exception handling. If this method fails, something bad happened anyway</remarks>
        internal int GetID()
        {
            string idPart = RID.Substring(3);
            return int.Parse(idPart);
        }
    }
}
