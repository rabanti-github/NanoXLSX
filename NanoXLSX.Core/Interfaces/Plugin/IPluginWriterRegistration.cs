/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
namespace NanoXLSX.Interfaces.Writer
{
    /// <summary>
    /// Interface, used by classes to register package parts prior to the XLSX creation process
    /// </summary>
    internal interface IPlugInWriterRegistration : IPlugIn
    {
        /// <summary>
        /// Order number of the package part (for sorting purpose during registration)
        /// </summary>
        int OrderNumber { get; set; }
        /// <summary>
        /// Relative path of the package part
        /// </summary>
        string PackagePartPath { get; set; }
        /// <summary>
        /// File name of the package part
        /// </summary>
        string PackagePartFileName { get; set; }
        /// <summary>
        /// Content type of the target file of the part (usually kind of XML)
        /// </summary>
        string ContentType { get; set; }
        /// <summary>
        /// Schema URL of the target file of the part (usually kind of XML schema)
        /// </summary>
        string RelationshipType { get; set; }
        /// <summary>
        /// If true, the package part is in the root directory, otherwise in the 'xl' sub-directory (with various sub-sub-directories)
        /// </summary>
        bool IsRootPackagaePart { get; set; }
    }
}
