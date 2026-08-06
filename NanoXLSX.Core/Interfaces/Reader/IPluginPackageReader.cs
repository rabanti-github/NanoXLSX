/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

namespace NanoXLSX.Interfaces.Reader
{
    /// <summary>
    /// Defines a queued package reader that optionally reads one ZIP entry with a fixed, known path.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The reader is executed once when its queue is processed. If <see cref="StreamEntryName"/> contains a path,
    /// that exact ZIP entry is opened and supplied to the reader. If the entry does not exist, the reader is skipped.
    /// If the property is <see langword="null"/> or empty, the reader is executed with a <see langword="null"/> stream.
    /// For example, returning <c>xl/custom.xml</c> requests that fixed package part.
    /// </para>
    /// <para>
    /// Use <see cref="IDiscoveryPackageReader"/> instead when the part path is obtained from OOXML relationships
    /// and may contain counters or otherwise vary between packages, such as
    /// <c>xl/externalLinks/externalLink1.xml</c> and <c>externalLink2.xml</c>.
    /// </para>
    /// </remarks>
    internal interface IPluginPackageReader : IPluginQueueReader
    {
        /// <summary>
        /// Gets the exact, case-sensitive path of the ZIP entry to read, or <see langword="null"/> to execute without an entry stream.
        /// </summary>
        string StreamEntryName { get; }
    }
}
