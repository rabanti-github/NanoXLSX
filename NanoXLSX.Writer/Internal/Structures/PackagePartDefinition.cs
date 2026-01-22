/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Collections.Generic;
using System.Linq;

namespace NanoXLSX.Internal.Structures
{
    /// <summary>
    /// Class to manage package parts of a XLSX file to be written
    /// </summary>
    /// \remark <remarks>This class is only for internal use. Use the high level API (e.g. class Workbook) to manipulate data and create Excel files</remarks>
    internal class PackagePartDefinition
    {

        #region staticFields
        /// <summary>
        /// Package part definition index (for sorting), designated to the workbook
        /// </summary>
        internal const int WORKBOOK_PACKAGE_PART_INDEX = 0;
        /// <summary>
        /// Package part definition start index (for sorting), designated for metadata and other root parts
        /// </summary>
        internal const int METADATA_PACKAGE_PART_START_INDEX = 1000;
        /// <summary>
        /// Package part definition start index (for sorting), designated for worksheet parts. These numbers shall not be used for other instances until <see cref="POST_WORSHEET_PACKAGE_PART_START_INDEX"/>
        /// </summary>
        internal const int WORKSHEET_PACKAGE_PART_START_INDEX = 10000;
        /// <summary>
        /// Package part definition start index (for sorting), designated for other non-root package parts
        /// </summary>
        internal const int POST_WORSHEET_PACKAGE_PART_START_INDEX = 2000000;
        #endregion

        #region enums
        /// <summary>
        /// Enum to define the relation ID of package parts
        /// </summary>
        internal enum PackagePartType
        {
            /// <summary>
            /// Package part is a root part (e.g. workbook, metadata)
            /// </summary>
            Root,
            /// <summary>
            /// Package part is a worksheet
            /// </summary>
            Worksheet,
            /// <summary>
            /// Package part is a non-root part (e.g. style, sharedStrings)
            /// </summary>
            Other,
        }
        #endregion

        /// <summary>
        /// Document path of the package part
        /// </summary>
        public DocumentPath Path { get; private set; }
        /// <summary>
        /// Type of the package part, used for handling differentiation
        /// </summary>
        public PackagePartType PartType { get; private set; }
        /// <summary>
        /// Order number during registration. The order number can be used to place package part into specific orders and therefore to enforce specific rIDs for the XML part 
        /// </summary>
        public int OrderNumber { get; private set; }
        /// <summary>
        /// Content type of the target file of the part (usually kind of XML)
        /// </summary>
        public string ContentType { get; private set; }
        /// <summary>
        /// Schema URL of the target file of the part (usually kind of XML schema)
        /// </summary>
        public string RelationshipType { get; private set; }

        /// <summary>
        /// Constructor with all fields
        /// </summary>
        /// <param name="type">Type of the package part, used for handling differentiation</param>
        /// <param name="orderNumber">Order number during registration</param>
        /// <param name="fileNameInPackage">Relative file name of the target file of the package part, without path</param>
        /// <param name="pathInPackage">Relative path to the file of the package part</param>
        /// <param name="contentType">Content type of the target file of the part (usually kind of XML)</param>
        /// <param name="relationshipType">Schema URL of the target file of the part (usually kind of XML schema)</param>
        public PackagePartDefinition(PackagePartType type, int orderNumber, string fileNameInPackage, string pathInPackage, string contentType, string relationshipType)
            : this(type, orderNumber, new DocumentPath(fileNameInPackage, pathInPackage), contentType, relationshipType)
        { }

        /// <summary>
        /// Constructor with definition of a document path
        /// </summary>
        /// <param name="type">Type of the package part, used for handling differentiation</param>
        /// <param name="orderNumber">Order number during registration</param>
        /// <param name="documentPath">Document path with all relevant file and path information</param>
        /// <param name="contentType">Content type of the target file of the part (usually kind of XML)</param>
        /// <param name="relationshipType">Schema URL of the target file of the part (usually kind of XML schema)</param>
        internal PackagePartDefinition(PackagePartType type, int orderNumber, DocumentPath documentPath, string contentType, string relationshipType)
        {
            this.PartType = type;
            this.OrderNumber = orderNumber;
            this.Path = documentPath;
            this.ContentType = contentType;
            this.RelationshipType = relationshipType;
        }

        /// <summary>
        /// Function to get the zero-based index of a worksheet
        /// </summary>
        /// <returns>Zero-based index of the worksheet, represented by this package part</returns>
        /// \remark <remarks>This method should only be used if the package part represents a worksheet. It will return invalid indices otherwise</remarks>
        internal int GetWorksheetIndex()
        {
            return this.OrderNumber - WORKSHEET_PACKAGE_PART_START_INDEX;
        }

        /// <summary>
        /// Static method to sort a list of package part definitions, based on the order number
        /// </summary>
        /// <param name="packagePartDefinitions">List to sort</param>
        /// <returns>Sorted list</returns>
        internal static List<PackagePartDefinition> Sort(List<PackagePartDefinition> packagePartDefinitions)
        {
            return packagePartDefinitions.OrderBy(p => p.OrderNumber).ToList();
        }
    }

}
