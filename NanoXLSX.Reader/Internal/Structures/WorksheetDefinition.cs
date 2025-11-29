/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

namespace NanoXLSX.Internal.Readers
{
    /// <summary>
    /// Class for worksheet Meta-data on import
    /// </summary>
    internal class WorksheetDefinition
    {
        /// <summary>
        /// Worksheet name
        /// </summary>
        public string WorksheetName { get; set; }
        /// <summary>
        /// Hidden state of the worksheet
        /// </summary>
        public bool Hidden { get; set; }
        /// <summary>
        /// Internal worksheet ID
        /// </summary>
        public int SheetID { get; set; }
        /// <summary>
        /// Reference ID
        /// </summary>
        public string RelId { get; set; }
        /// <summary>
        /// Default constructor with parameters
        /// </summary>
        /// <param name="id">Internal ID</param>
        /// <param name="name">Worksheet name</param>
        /// <param name="relId">Relation ID</param>
        public WorksheetDefinition(int id, string name, string relId)
        {
            this.SheetID = id;
            this.WorksheetName = name;
            this.RelId = relId;
        }
    }

}
