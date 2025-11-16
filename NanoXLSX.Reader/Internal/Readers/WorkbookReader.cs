/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.IO;
using System.Xml;
using NanoXLSX.Exceptions;
using NanoXLSX.Interfaces.Plugin;
using NanoXLSX.Interfaces.Reader;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;
using NanoXLSX.Utils;
using NanoXLSX.Utils.Xml;
using static NanoXLSX.Internal.Enums.ReaderPassword;
using IOException = NanoXLSX.Exceptions.IOException;

namespace NanoXLSX.Internal.Readers
{
    /// <summary>
    /// Class representing a reader to decompile a workbook in an XLSX files
    /// </summary>
    [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.WorkbookReader)]
    public partial class WorkbookReader : IPlugInReader
    {
        private MemoryStream stream;
        private IPasswordReader passwordReader;
        private ReaderOptions readerOptions;

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
        public void Init(MemoryStream stream, Workbook workbook, IOptions readerOptions)
        {
            this.stream = stream;
            this.Workbook = workbook;
            this.readerOptions = (ReaderOptions)readerOptions;
            this.passwordReader = PlugInLoader.GetPlugIn<IPasswordReader>(PlugInUUID.PasswordReader, new LegacyPasswordReader());
            this.passwordReader.Init(PasswordType.WorkbookProtection, this.readerOptions);
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
                    XmlDocument xr = new XmlDocument() { XmlResolver = null };
                    using (XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings() { XmlResolver = null }))
                    {
                        xr.Load(reader);
                        foreach (XmlNode node in xr.DocumentElement.ChildNodes)
                        {
                            if (node.LocalName.Equals("sheets", StringComparison.OrdinalIgnoreCase) && node.HasChildNodes)
                            {
                                GetWorksheetInformation(node.ChildNodes);
                            }
                            else if (node.LocalName.Equals("bookViews", StringComparison.OrdinalIgnoreCase) && node.HasChildNodes)
                            {
                                GetViewInformation(node.ChildNodes);
                            }
                            else if (node.LocalName.Equals("workbookProtection", StringComparison.OrdinalIgnoreCase))
                            {
                                GetProtectionInformation(node);
                            }
                        }
                        RederPlugInHandler.HandleInlineQueuePlugins(ref stream, Workbook, PlugInUUID.WorkbookInlineReader);
                    }
                }
            }
            catch (NotSupportedContentException ex)
            {
                throw ex; // rethrow
            }
            catch (Exception ex)
            {
                throw new IOException("The XML entry could not be read from the input stream. Please see the inner exception:", ex);
            }
        }

        /// <summary>
        /// Gets the workbook protection information
        /// </summary>
        /// <param name="node">Root node to check</param>
        private void GetProtectionInformation(XmlNode node)
        {
            bool lockStructure = false;
            bool lockWindows = false;
            //this.Protected = true;
            string attribute = ReaderUtils.GetAttribute(node, "lockWindows");
            if (attribute != null)
            {
                int value = ParserUtils.ParseBinaryBool(attribute);
                //this.LockWindows = value == 1;
                lockWindows = value == 1;
            }
            attribute = ReaderUtils.GetAttribute(node, "lockStructure");
            if (attribute != null)
            {
                int value = ParserUtils.ParseBinaryBool(attribute);
                //this.LockStructure = value == 1;
                lockStructure = value == 1;
            }
            Workbook.SetWorkbookProtection(true, lockWindows, lockStructure, null);
            passwordReader.ReadXmlAttributes(node);
            if (passwordReader.PasswordIsSet())
            {
                Workbook.WorkbookProtectionPassword.CopyFrom(passwordReader);
            }
        }

        /// <summary>
        /// Gets the workbook view information
        /// </summary>
        /// <param name="nodes">View nodes to check</param>
        private void GetViewInformation(XmlNodeList nodes)
        {
            foreach (XmlNode node in nodes)
            {
                if (node.LocalName.Equals("workbookView", StringComparison.OrdinalIgnoreCase))
                {
                    string attribute = ReaderUtils.GetAttribute(node, "visibility");
                    if (attribute != null && ParserUtils.ToLower(attribute) == "hidden")
                    {
                        this.Workbook.Hidden = true;
                    }
                    attribute = ReaderUtils.GetAttribute(node, "activeTab");
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
        /// <param name="nodes">Sheet nodes to check</param>
        private void GetWorksheetInformation(XmlNodeList nodes)
        {
            foreach (XmlNode node in nodes)
            {
                if (node.LocalName.Equals("sheet", StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        string sheetName = ReaderUtils.GetAttribute(node, "name", "worksheet1");
                        int id = ParserUtils.ParseInt(ReaderUtils.GetAttribute(node, "sheetId")); // Default will rightly throw an exception
                        string relId = ReaderUtils.GetAttribute(node, "r:id");
                        string state = ReaderUtils.GetAttribute(node, "state");
                        bool hidden = false;
                        if (state != null && ParserUtils.ToLower(state) == "hidden")
                        {
                            hidden = true;
                        }
                        WorksheetDefinition definition = new WorksheetDefinition(id, sheetName, relId);
                        definition.Hidden = hidden;
                        Workbook.AuxiliaryData.SetData(PlugInUUID.WorkbookReader, PlugInUUID.WorksheetDefinitionEntity, id, definition);
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
