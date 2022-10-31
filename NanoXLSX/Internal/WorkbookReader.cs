/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2022
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Shared.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using IOException = NanoXLSX.Shared.Exceptions.IOException;

namespace NanoXLSX.Internal
{
    /// <summary>
    /// Class representing a reader to decompile a workbook in an XLSX files
    /// </summary>
    public class WorkbookReader
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
        /// <summary>
        /// Password hash, if available
        /// </summary>
        public string PasswordHash { get; private set; }

        #endregion

        #region constructors

        /// <summary>
        /// Default constructor
        /// </summary>
        public WorkbookReader()
        {
            WorksheetDefinitions = new Dictionary<int, WorksheetDefinition>();
        }

        #endregion

        #region functions

        /// <summary>
        /// Reads the XML file form the passed stream and processes the workbook information
        /// </summary>
        /// <param name="stream">Stream of the XML file</param>
        /// <exception cref="NanoXLSX.Shared.Exceptions.IOException">Throws IOException in case of an error</exception>
        public void Read(MemoryStream stream)
        {
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
        }

        /// <summary>
        /// Gets the workbook protection information
        /// </summary>
        /// <param name="node">Root node to check</param>
        private void GetProtectionInformation(XmlNode node)
        {
            this.Protected = true;
            string attribute = ReaderUtils.GetAttribute(node, "lockWindows");
            if (attribute != null && attribute == "1")
            {
                this.LockWindows = true;
            }
            attribute = ReaderUtils.GetAttribute(node, "lockStructure");
            if (attribute != null && attribute == "1")
            {
                this.LockStructure = true;
            }
            attribute = ReaderUtils.GetAttribute(node, "workbookPassword");
            if (attribute != null)
            {
                this.PasswordHash = attribute;
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
            foreach(XmlNode node in nodes)
            {
                if (node.LocalName.Equals("sheet", StringComparison.InvariantCultureIgnoreCase))
                {
                    try
                    {
                        string sheetName = ReaderUtils.GetAttribute(node, "name", "worksheet1");
                        int id = ParserUtils.ParseInt(ReaderUtils.GetAttribute(node, "sheetId")); // Default will rightly throw an exception
                        string state = ReaderUtils.GetAttribute(node, "state");
                        bool hidden = false;
                        if (state != null && state.ToLower() == "hidden")
                        {
                            hidden = true;
                        }
                        WorksheetDefinition definition = new WorksheetDefinition(id, sheetName);
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
            /// Default constructor with parameters
            /// </summary>
            /// <param name="id">Internal ID</param>
            /// <param name="name">Worksheet name</param>
            public WorksheetDefinition(int id, string name)
            {
                this.SheetID = id;
                this.WorksheetName = name;
            }
        }

        #endregion

    }
}
