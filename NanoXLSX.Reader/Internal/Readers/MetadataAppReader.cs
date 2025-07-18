/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.IO;
using System.Xml;
using NanoXLSX.Interfaces.Plugin;
using NanoXLSX.Interfaces.Reader;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;

namespace NanoXLSX.Internal.Readers
{
    /// <summary>
    /// Class representing a reader for the App metadata file (docProps) embedded in XLSX files
    /// </summary>
    [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.METADATA_APP_READER)]
    public class MetadataAppReader : IPlugInReader
    {
        private MemoryStream stream;

        #region properties
        /// <summary>
        /// Workbook reference where read data is stored (should not be null)
        /// </summary>
        public Workbook Workbook { get; set; }
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
        public void Init(MemoryStream stream, Workbook workbook, IOptions readerOptions)
        {
            this.stream = stream;
            this.Workbook = workbook;
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

                    XmlDocument xr = new XmlDocument();
                    xr.XmlResolver = null;
                    xr.Load(stream);
                    foreach (XmlNode node in xr.DocumentElement.ChildNodes)
                    {
                        if (node.LocalName.Equals("Application", StringComparison.InvariantCultureIgnoreCase))
                        {
                            metadata.Application = node.InnerText;
                        }
                        else if (node.LocalName.Equals("AppVersion", StringComparison.InvariantCultureIgnoreCase))
                        {
                            metadata.ApplicationVersion = node.InnerText;
                        }
                        else if (node.LocalName.Equals("Company", StringComparison.InvariantCultureIgnoreCase))
                        {
                            metadata.Company = node.InnerText;
                        }
                        else if (node.LocalName.Equals("Manager", StringComparison.InvariantCultureIgnoreCase))
                        {
                            metadata.Manager = node.InnerText;
                        }
                        else if (node.LocalName.Equals("HyperlinkBase", StringComparison.InvariantCultureIgnoreCase))
                        {
                            metadata.HyperlinkBase = node.InnerText;
                        }
                    }
                    RederPlugInHandler.HandleInlineQueuePlugins(ref stream, Workbook, PlugInUUID.METADATA_APP_INLINE_READER);
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
