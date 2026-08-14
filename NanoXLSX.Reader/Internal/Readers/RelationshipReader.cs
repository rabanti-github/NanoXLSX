/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.IO;
using NanoXLSX.Interfaces;
using NanoXLSX.Interfaces.Reader;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;

namespace NanoXLSX.Internal.Readers
{
    // TODO V4 (next major version): remove
    /// <summary>
    /// Class representing the legacy workbook relationship reader of XLSX files.
    /// </summary>
    /// \remark <remarks>This reader is retained for plug-in and inline-hook compatibility. Relationship discovery is authoritative for document resolution. Reconsider removal only in the next major version.</remarks>
    [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.RelationshipReader)]
    [Obsolete("Will be removed with the next major version")]
    public partial class RelationshipReader : IPluginBaseReader
    {
        private Stream stream;

        #region properties

        /// <summary>
        /// Workbook reference where read data is stored (should not be null)
        /// </summary>
        public Workbook Workbook { get; set; }
        /// <summary>
        /// Reader options
        /// </summary>
        public IOptions Options { get; set; }
        /// <summary>
        /// Reference to the <see cref="ReaderPlugInHandler"/>, to be used for post operations in the <see cref="Execute"/> method
        /// </summary>
        public Action<Stream, Workbook, string, IOptions, int?> InlinePluginHandler { get; set; }

        #endregion

        #region constructor 
        /// <summary>
        /// Default constructor - Must be defined for instantiation of the plug-ins
        /// </summary>
        public RelationshipReader()
        {
        }
        #endregion

        #region functions
        /// <summary>
        /// Initialization method (interface implementation)
        /// </summary>
        /// <param name="stream">MemoryStream to be read</param>
        /// <param name="workbook">Workbook reference</param>
        /// <param name="readerOptions">Reader options (NoOp)</param>
        /// <param name="inlinePluginHandler">Reference to the a handler action, to be used for post operations in reader methods</param>
        public void Init(Stream stream, Workbook workbook, IOptions readerOptions, Action<Stream, Workbook, string, IOptions, int?> inlinePluginHandler)
        {
            this.stream = stream;
            this.Workbook = workbook;
            this.Options = readerOptions;
            this.InlinePluginHandler = inlinePluginHandler;
        }

        /// <summary>
        /// Executes legacy relationship inline plug-ins without duplicating discovery parsing.
        /// </summary>
        public void Execute()
        {
            // V4 TODO (next major version): Replace this compatibility staging reader after its UUID and inline plug-in contracts can be retired.
            if (stream == null) return;
            using (stream) // Close after processing
            {
                InlinePluginHandler?.Invoke(stream, Workbook, PlugInUUID.RelationshipInlineReader, Options, null);
            }
        }
        #endregion
    }
}
