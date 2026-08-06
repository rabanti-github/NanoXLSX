/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Internal;

namespace NanoXLSX.Interfaces.Reader
{
    /// <summary>
    /// Defines a queued reader that is dispatched to package parts by their discovered OOXML relationship type.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The package relationship discovery step runs before this reader. For every internal relationship whose
    /// <see cref="IDocumentReader.DocumentType"/> exactly matches the relationship type and whose target exists
    /// in the package, the reader is initialized with a fresh stream of that target and then executed once.
    /// <see cref="CurrentRelationship"/> identifies the relationship for the current execution.
    /// </para>
    /// <para>
    /// This contract is intended for relationship targets with non-fixed names. For example, a reader whose
    /// document type is <c>http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink</c>
    /// is dispatched to <c>xl/externalLinks/externalLink1.xml</c>, <c>externalLink2.xml</c>, and every other
    /// matching internal target discovered in the package.
    /// </para>
    /// <para>
    /// This reader does not perform relationship discovery and does not receive the corresponding
    /// <c>*.rels</c> stream. Relationship parts have already been parsed into the discovery catalog.
    /// The inherited <see cref="IPluginPackageReader.StreamEntryName"/> is not used for discovery-based
    /// dispatch and should return <see langword="null"/>.
    /// </para>
    /// </remarks>
    internal interface IDiscoveryPackageReader : IPluginPackageReader, IDocumentReader
    {
        /// <summary>
        /// Gets or sets the discovered relationship whose resolved target stream is supplied for the current execution.
        /// </summary>
        RelationshipInfo CurrentRelationship { get; set; }
    }
}
