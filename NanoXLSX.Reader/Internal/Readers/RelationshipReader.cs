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
using NanoXLSX.Utils;
using IOException = NanoXLSX.Exceptions.IOException;

namespace NanoXLSX.Internal.Readers
{

    /// <summary>
    /// Class representing a reader for relationship of XLSX files
    /// </summary>
    [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.RelationshipReader)]
    public partial class RelationshipReader : IPluginBaseReader
    {
        private MemoryStream stream;

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
        public Action<MemoryStream, Workbook, string, IOptions, int?> InlinePluginHandler { get; set; }

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
        public void Init(MemoryStream stream, Workbook workbook, IOptions readerOptions, Action<MemoryStream, Workbook, string, IOptions, int?> inlinePluginHandler)
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
            if (stream == null) return;
            try
            {
                XmlDocument xr;
                using (stream) // Close after processing
                {
                    xr = new XmlDocument
                    {
                        XmlResolver = null
                    };
                    using (XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings() { XmlResolver = null }))
                    {
                        xr.Load(reader);
                        XmlNodeList relationships = xr.GetElementsByTagName("Relationship");
                        foreach (XmlNode relationship in relationships)
                        {
                            string id = ReaderUtils.GetAttribute(relationship, "Id");
                            string type = ReaderUtils.GetAttribute(relationship, "Type");
                            string target = ReaderUtils.GetAttribute(relationship, "Target");
                            if (ParserUtils.StartsWith(target, "/"))
                            {
                                target = target.TrimStart('/');
                            }
                            if (ParserUtils.NotStartsWith(target, "xl/"))
                            {
                                target = "xl/" + target;
                            }
                            Relationship rel = new Relationship
                            {
                                RID = id,
                                Type = type,
                                Target = target,
                            };
                            Workbook.AuxiliaryData.SetData(PlugInUUID.RelationshipReader, PlugInUUID.RelationshipEntity, id, rel);
                        }
                        InlinePluginHandler?.Invoke(stream, Workbook, PlugInUUID.RelationshipInlineReader, Options, null);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new IOException("The XML entry could not be read from the input stream. Please see the inner exception:", ex);
            }
        }
        #endregion
    }
}
