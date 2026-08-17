/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Collections.Generic;

namespace NanoXLSX.Interfaces.Writer
{
    /// <summary>
    /// Interface, used by classes to register package parts at the beginning of the writer process
    /// </summary>
    /// \remark <remarks>All collection entries must contain the <b>same number of elements</b>. Package part indices must be unique within a writing operation and are compared ordinally.</remarks>
    internal interface IPluginPackageRegistry : IPlugin
    {
        /// <summary>
        /// Gets or replaces the workbook instance, defined by the constructor
        /// </summary>
        Workbook Workbook { get; set; }

        /// <summary>
        /// Initializes the registry for the current writing operation
        /// </summary>
        /// <param name="baseWriter">Base writer instance that holds the current writing context</param>
        void Init(IBaseWriter baseWriter);

        /// <summary>
        /// List of order numbers of the package parts (for sorting purpose during registration)
        /// </summary>
        IReadOnlyList<int> OrderNumbers { get; }
        /// <summary>
        /// List of relative paths of the package parts
        /// </summary>
        IReadOnlyList<string> PackagePartPaths { get; }
        /// <summary>
        /// List of the file names of the package parts
        /// </summary>
        IReadOnlyList<string> PackagePartFileNames { get; }
        /// <summary>
        /// List of the content types of the target file of the parts (usually kind of XML)
        /// </summary>
        IReadOnlyList<string> ContentTypes { get; }
        /// <summary>
        /// List of the schema URLs of the target file of the parts (usually kind of XML schema)
        /// </summary>
        IReadOnlyList<string> RelationshipTypes { get; }
        /// <summary>
        /// List of location statement. If true, the package part is in the root directory, otherwise in the 'xl' sub-directory (with various sub-sub-directories)
        /// </summary>
        IReadOnlyList<bool> ArePackagePartsRoot { get; }
        /// <summary>
        /// List of unique index indicators that can be used from <see cref="IPluginIndexedWriter"/> instances to write package parts
        /// </summary>
        IReadOnlyList<string> UniquePackagePartIndices { get; }
    }
}
