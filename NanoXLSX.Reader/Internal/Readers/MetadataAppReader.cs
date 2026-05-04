/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.IO;
using System.Xml;
using NanoXLSX.Interfaces;
using NanoXLSX.Interfaces.Reader;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;
using NanoXLSX.Utils.Xml;

namespace NanoXLSX.Internal.Readers
{
    /// <summary>
    /// Class representing a reader for the App metadata file (docProps) embedded in XLSX files
    /// </summary>
    [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.MetadataAppReader)]
    public class MetadataAppReader : IPluginBaseReader
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

        #region constructors
        /// <summary>
        /// Default constructor - Must be defined for instantiation of the plug-ins
        /// </summary>
        internal MetadataAppReader()
        {
        }

        #endregion

        #region methods
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
        /// Method to execute the main logic of the plug-in (interface implementation)
        /// </summary>
        /// <exception cref="Exceptions.IOException">Throws an IOException in case of a error during reading</exception>
        public void Execute()
        {
            try
            {
                using (stream) // Close after processing
                {
                    Metadata metadata = Workbook.WorkbookMetadata;
                    using (XmlReader reader = XmlReader.Create(stream, XmlStreamUtils.CreateSettings()))
                    {
                        while (reader.Read())
                        {
                            if (reader.NodeType != XmlNodeType.Element)
                            {
                                continue;
                            }
                            if (XmlStreamUtils.IsElement(reader, "Application"))
                            {
                                metadata.Application = XmlStreamUtils.ReadElementText(reader);
                            }
                            else if (XmlStreamUtils.IsElement(reader, "AppVersion"))
                            {
                                metadata.ApplicationVersion = XmlStreamUtils.ReadElementText(reader);
                            }
                            else if (XmlStreamUtils.IsElement(reader, "Company"))
                            {
                                metadata.Company = XmlStreamUtils.ReadElementText(reader);
                            }
                            else if (XmlStreamUtils.IsElement(reader, "Manager"))
                            {
                                metadata.Manager = XmlStreamUtils.ReadElementText(reader);
                            }
                            else if (XmlStreamUtils.IsElement(reader, "HyperlinkBase"))
                            {
                                metadata.HyperlinkBase = XmlStreamUtils.ReadElementText(reader);
                            }
                        }
                        InlinePluginHandler?.Invoke(stream, Workbook, PlugInUUID.MetadataAppInlineReader, Options, null);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new NanoXLSX.Exceptions.IOException("The XML entry could not be read from the input stream. Please see the inner exception:", ex);
            }
        }
        #endregion
    }
}
