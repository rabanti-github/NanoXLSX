/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.IO.Compression;
using NanoXLSX.Interfaces;
using NanoXLSX.Interfaces.Reader;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;

namespace NanoXLSX.Internal.Readers
{
    /// <summary>
    /// Discovers relationship information in an OOXML package before document readers execute.
    /// </summary>
    /// \remark <remarks>The relationship parsing and validation logic is introduced in the next discovery implementation step.</remarks>
    [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.DiscoveryReader)]
    public class DiscoveryReader : IDiscoveryReader
    {
        private ZipArchive archive;

        /// <summary>
        /// Gets or sets the workbook that receives the temporary discovery catalog.
        /// </summary>
        public Workbook Workbook { get; set; }

        /// <summary>
        /// Gets or sets the reader options used for discovery validation.
        /// </summary>
        public IOptions Options { get; set; }

        /// <summary>
        /// Initializes a new discovery reader.
        /// </summary>
        public DiscoveryReader()
        {
        }

        /// <summary>
        /// Initializes discovery with a caller-owned ZIP archive.
        /// </summary>
        public void Init(ZipArchive archive, Workbook workbook, IOptions readerOptions)
        {
            this.archive = archive;
            Workbook = workbook;
            Options = readerOptions;
        }

        /// <summary>
        /// Prepares the temporary relationship catalog.
        /// </summary>
        /// <exception cref="InvalidOperationException">Thrown when the reader was not initialized with a ZIP archive or workbook.</exception>
        public void Execute()
        {
            if (archive == null)
            {
                throw new InvalidOperationException("The discovery reader was not initialized with a ZIP archive.");
            }
            if (Workbook == null)
            {
                throw new InvalidOperationException("The discovery reader was not initialized with a workbook.");
            }
            RelationshipCatalog catalog = new RelationshipCatalog();
            Workbook.AuxiliaryData.SetData(PlugInUUID.DiscoveryReader, PlugInUUID.DiscoveryCatalogEntity, catalog);
        }
    }
}
