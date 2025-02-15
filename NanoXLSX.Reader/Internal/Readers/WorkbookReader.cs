﻿/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using NanoXLSX.Interfaces.Reader;
using NanoXLSX.Utils;
using IOException = NanoXLSX.Exceptions.IOException;

namespace NanoXLSX.Internal.Readers
{
    /// <summary>
    /// Class representing a reader to decompile a workbook in an XLSX files
    /// </summary>
    public class WorkbookReader : IPluginReader
    {

        #region properties

        /// <summary>
        /// Dictionary of worksheet definitions. The key is the worksheet number and the value is a WorksheetDefinition object with name, hidden state and other information
        /// </summary>
        /// <value>
        /// Dictionary with worksheet definitions
        /// </value>
        public Dictionary<int, WorksheetDefinition> WorksheetDefinitions { get; private set; }

        /// <summary>
        /// Hidden state of the workbook
        /// </summary>
        public bool Hidden { get; private set; }
        /// <summary>
        /// Selected worksheet of the workbook
        /// </summary>
        public int SelectedWorksheet { get; private set; }
        /// <summary>
        /// Protection state of the workbook
        /// </summary>
        public bool Protected { get; private set; }
        /// <summary>
        /// Lock state of the windows
        /// </summary>
        public bool LockWindows { get; private set; }
        /// <summary>
        /// Lock state of the structural elements
        /// </summary>
        public bool LockStructure { get; private set; }

        public IPasswordReader PasswordReader { get; internal set; }

        #endregion

        #region constructors

        /// <summary>
        /// Default constructor
        /// </summary>
        public WorkbookReader()
        {
            WorksheetDefinitions = new Dictionary<int, WorksheetDefinition>();
            // TODO add hook to overwrite password reader
            PasswordReader = new LegacyPasswordReader(LegacyPasswordReader.PasswordType.WORKBOOK_PROTECTION);
        }

        #endregion

        #region functions

        /// <summary>
        /// Reads the XML file form the passed stream and processes the workbook information
        /// </summary>
        /// <param name="stream">Stream of the XML file</param>
        /// \remark <remarks>This method is virtual. Plug-in packages may override it</remarks>
        /// <exception cref="NanoXLSX.Exceptions.IOException">Throws IOException in case of an error</exception>
        public virtual void Read(MemoryStream stream)
        {
            PreRead(stream);
            try
            {
                using (stream) // Close after processing
                {
                    XmlDocument xr = new XmlDocument();
                    xr.XmlResolver = null;
                    xr.Load(stream);
                    foreach (XmlNode node in xr.DocumentElement.ChildNodes)
                    {
                        if (node.LocalName.Equals("sheets", StringComparison.InvariantCultureIgnoreCase) && node.HasChildNodes)
                        {
                            GetWorksheetInformation(node.ChildNodes);
                        }
                        else if (node.LocalName.Equals("bookViews", StringComparison.InvariantCultureIgnoreCase) && node.HasChildNodes)
                        {
                            GetViewInformation(node.ChildNodes);
                        }
                        else if (node.LocalName.Equals("workbookProtection", StringComparison.InvariantCultureIgnoreCase))
                        {
                            GetProtectionInformation(node);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new IOException("The XML entry could not be read from the input stream. Please see the inner exception:", ex);
            }
            PostRead(stream);
        }

        /// <summary>
        /// Method that is called before the <see cref="Read(MemoryStream)"/> method is executed. 
        /// This virtual method is empty by default and can be overridden by a plug-in package
        /// </summary>
        /// <param name="stream">Stream of the XML file. The stream must be reset in this method at the end, if any stream opeartion was performed</param>
        public virtual void PreRead(MemoryStream stream)
        {
            // NoOp - replaced by plugin
        }

        /// <summary>
        /// Method that is called after the <see cref="Read(MemoryStream)"/> method is executed. 
        /// This virtual method is empty by default and can be overridden by a plug-in package
        /// </summary>
        /// <param name="stream">Stream of the XML file. The stream must be reset in this method before any stream operation is performed</param>
        public virtual void PostRead(MemoryStream stream)
        {
            // NoOp - replaced by plugin
        }

        /// <summary>
        /// Gets the workbook protection information
        /// </summary>
        /// <param name="node">Root node to check</param>
        private void GetProtectionInformation(XmlNode node)
        {
            this.Protected = true;
            string attribute = ReaderUtils.GetAttribute(node, "lockWindows");
            if (attribute != null)
            {
                int value = ParserUtils.ParseBinaryBool(attribute);
                this.LockWindows = value == 1;
            }
            attribute = ReaderUtils.GetAttribute(node, "lockStructure");
            if (attribute != null)
            {
                int value = ParserUtils.ParseBinaryBool(attribute);
                this.LockStructure = value == 1;
            }
            PasswordReader.ReadXmlAttributes(node);
        }

        /// <summary>
        /// Gets the workbook view information
        /// </summary>
        /// <param name="nodes">View nodes to check</param>
        private void GetViewInformation(XmlNodeList nodes)
        {
            foreach (XmlNode node in nodes)
            {
                if (node.LocalName.Equals("workbookView", StringComparison.InvariantCultureIgnoreCase))
                {
                    string attribute = ReaderUtils.GetAttribute(node, "visibility");
                    if (attribute != null && attribute.ToLower() == "hidden")
                    {
                        this.Hidden = true;
                    }
                    attribute = ReaderUtils.GetAttribute(node, "activeTab");
                    if (!string.IsNullOrEmpty(attribute))
                    {
                        this.SelectedWorksheet = ParserUtils.ParseInt(attribute);
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
                if (node.LocalName.Equals("sheet", StringComparison.InvariantCultureIgnoreCase))
                {
                    try
                    {
                        string sheetName = ReaderUtils.GetAttribute(node, "name", "worksheet1");
                        int id = ParserUtils.ParseInt(ReaderUtils.GetAttribute(node, "sheetId")); // Default will rightly throw an exception
                        string relId = ReaderUtils.GetAttribute(node, "r:id");
                        string state = ReaderUtils.GetAttribute(node, "state");
                        bool hidden = false;
                        if (state != null && state.ToLower() == "hidden")
                        {
                            hidden = true;
                        }
                        WorksheetDefinition definition = new WorksheetDefinition(id, sheetName, relId);
                        definition.Hidden = hidden;
                        WorksheetDefinitions.Add(id, definition);
                    }
                    catch (Exception e)
                    {
                        throw new IOException("The workbook information could not be resolved. Please see the inner exception:", e);
                    }
                }
            }
        }

        #endregion

        #region subclasses

        /// <summary>
        /// Class for worksheet Mata-data on import
        /// </summary>
        public class WorksheetDefinition
        {
            /// <summary>
            /// Worksheet name
            /// </summary>
            public string WorksheetName { get; set; }
            /// <summary>
            /// Hidden state of the worksheet
            /// </summary>
            public bool Hidden { get; set; }
            /// <summary>
            /// Internal worksheet ID
            /// </summary>
            public int SheetID { get; set; }
            /// <summary>
            /// Reference ID
            /// </summary>
            public string RelId { get; set; }
            /// <summary>
            /// Default constructor with parameters
            /// </summary>
            /// <param name="id">Internal ID</param>
            /// <param name="name">Worksheet name</param>
            /// <param name="relId">Relation ID</param>
            public WorksheetDefinition(int id, string name, string relId)
            {
                this.SheetID = id;
                this.WorksheetName = name;
                this.RelId = relId;
            }
        }

        #endregion

    }
}
