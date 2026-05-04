/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.IO;
using System.Xml;
using NanoXLSX.Exceptions;
using NanoXLSX.Interfaces;
using NanoXLSX.Interfaces.Reader;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;
using NanoXLSX.Utils;
using NanoXLSX.Utils.Xml;
using static NanoXLSX.Enums.Password;
using IOException = NanoXLSX.Exceptions.IOException;

namespace NanoXLSX.Internal.Readers
{
    /// <summary>
    /// Class representing a reader to decompile a workbook in an XLSX files
    /// </summary>
    [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.WorkbookReader)]
    public partial class WorkbookReader : IPluginBaseReader
    {
        private Stream stream;
        private IPasswordReader passwordReader;
        // private ReaderOptions readerOptions;

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
        internal WorkbookReader()
        {
        }
        #endregion

        #region methods
        /// <summary>
        /// Initialization method (interface implementation)
        /// </summary>
        /// <param name="stream">MemoryStream to be read</param>
        /// <param name="workbook">Workbook reference</param>
        /// <param name="readerOptions">Reader options</param>
        /// <param name="inlinePluginHandler">Reference to the a handler action, to be used for post operations in reader methods</param>
        public void Init(Stream stream, Workbook workbook, IOptions readerOptions, Action<Stream, Workbook, string, IOptions, int?> inlinePluginHandler)
        {
            this.stream = stream;
            this.Workbook = workbook;
            this.Options = readerOptions;
            this.InlinePluginHandler = inlinePluginHandler;
            this.passwordReader = PlugInLoader.GetPlugIn<IPasswordReader>(PlugInUUID.PasswordReader, new LegacyPasswordReader());
            this.passwordReader.Init(PasswordType.WorkbookProtection, (ReaderOptions)readerOptions);
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
                    using (XmlReader reader = XmlReader.Create(stream, XmlStreamUtils.CreateSettings()))
                    {
                        while (reader.Read())
                        {
                            if (reader.NodeType != XmlNodeType.Element)
                            {
                                continue;
                            }
                            if (XmlStreamUtils.IsElement(reader, "sheets"))
                            {
                                GetWorksheetInformation(reader);
                            }
                            else if (XmlStreamUtils.IsElement(reader, "bookViews"))
                            {
                                GetViewInformation(reader);
                            }
                            else if (XmlStreamUtils.IsElement(reader, "workbookProtection"))
                            {
                                GetProtectionInformation(reader);
                            }
                        }
                        InlinePluginHandler?.Invoke(stream, Workbook, PlugInUUID.WorkbookInlineReader, Options, null);
                    }
                }
            }
            catch (NotSupportedContentException)
            {
                throw; // rethrow
            }
            catch (Exception ex)
            {
                throw new IOException("The XML entry could not be read from the input stream. Please see the inner exception:", ex);
            }
        }

        /// <summary>
        /// Gets the workbook protection information. Because the password reader interface requires an
        /// XmlNode, the protection element is captured via ReadOuterXml and loaded into a minimal
        /// temporary XmlDocument — a single-element DOM allocation that is negligible in practice.
        /// </summary>
        /// <param name="reader">Reader positioned on the workbookProtection start element</param>
        private void GetProtectionInformation(XmlReader reader)
        {
            bool lockStructure = false;
            bool lockWindows = false;
            string attribute = reader.GetAttribute("lockWindows");
            if (attribute != null)
            {
                lockWindows = ParserUtils.ParseBinaryBool(attribute) == 1;
            }
            attribute = reader.GetAttribute("lockStructure");
            if (attribute != null)
            {
                lockStructure = ParserUtils.ParseBinaryBool(attribute) == 1;
            }
            Workbook.SetWorkbookProtection(true, lockWindows, lockStructure, null);
            string outerXml;
            using (XmlReader subtree = reader.ReadSubtree())
            {
                subtree.MoveToContent();
                outerXml = subtree.ReadOuterXml();
            }
            XmlDocument miniDoc = new XmlDocument { XmlResolver = null };
            miniDoc.LoadXml(outerXml);
            passwordReader.ReadXmlAttributes(miniDoc.DocumentElement);
            if (passwordReader.PasswordIsSet())
            {
                Workbook.WorkbookProtectionPassword.CopyFrom(passwordReader);
            }
        }

        /// <summary>
        /// Gets the workbook view information
        /// </summary>
        /// <param name="reader">Reader positioned on the bookViews start element</param>
        private void GetViewInformation(XmlReader reader)
        {
            using (XmlReader subtree = reader.ReadSubtree())
            {
                subtree.Read(); // consume the bookViews open tag
                while (subtree.Read())
                {
                    if (!XmlStreamUtils.IsElement(subtree, "workbookView"))
                    {
                        continue;
                    }
                    string attribute = subtree.GetAttribute("visibility");
                    if (attribute != null && ParserUtils.ToLower(attribute) == "hidden")
                    {
                        Workbook.Hidden = true;
                    }
                    attribute = subtree.GetAttribute("activeTab");
                    if (!string.IsNullOrEmpty(attribute))
                    {
                        Workbook.AuxiliaryData.SetData(PlugInUUID.WorkbookReader, PlugInUUID.SelectedWorksheetEntity, ParserUtils.ParseInt(attribute));
                    }
                }
            }
        }

        /// <summary>
        /// Gets the worksheet information
        /// </summary>
        /// <param name="reader">Reader positioned on the sheets start element</param>
        private void GetWorksheetInformation(XmlReader reader)
        {
            int visibleWorksheetOrder = 0;
            using (XmlReader subtree = reader.ReadSubtree())
            {
                subtree.Read(); // consume the sheets open tag
                while (subtree.Read())
                {
                    if (!XmlStreamUtils.IsElement(subtree, "sheet"))
                    {
                        continue;
                    }
                    try
                    {
                        string sheetName = subtree.GetAttribute("name") ?? "worksheet1";
                        int id = ParserUtils.ParseInt(subtree.GetAttribute("sheetId")); // null will rightly throw
                        string relId = subtree.GetAttribute("r:id");
                        string state = subtree.GetAttribute("state");
                        bool hidden = state != null && ParserUtils.ToLower(state) == "hidden";
                        WorksheetDefinition definition = new WorksheetDefinition(id, sheetName, relId)
                        {
                            Hidden = hidden
                        };
                        Workbook.AuxiliaryData.SetData(PlugInUUID.WorkbookReader, PlugInUUID.WorksheetDefinitionEntity, visibleWorksheetOrder, definition);
                        visibleWorksheetOrder++;
                    }
                    catch (Exception e)
                    {
                        throw new IOException("The workbook information could not be resolved. Please see the inner exception:", e);
                    }
                }
            }
        }
        #endregion
    }
}
