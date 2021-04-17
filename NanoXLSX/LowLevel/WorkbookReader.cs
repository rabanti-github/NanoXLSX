/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2021
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Xml;
using NanoXLSX.Exceptions;
using IOException = NanoXLSX.Exceptions.IOException;

namespace NanoXLSX.LowLevel
{
    /// <summary>
    /// Class representing a reader to decompile a workbook in an XLSX files
    /// </summary>
    public class WorkbookReader
    {

        #region properties

        /// <summary>
        /// Dictionary of worksheet definitions. The key is the worksheet number and the value is the worksheet name
        /// </summary>
        /// <value>
        /// Dictionary with worksheet definitions
        /// </value>
        public Dictionary<int, string> WorksheetDefinitions { get; private set; }

        #endregion

        #region constructors

        /// <summary>
        /// Default constructor
        /// </summary>
        public WorkbookReader()
        {
            WorksheetDefinitions = new Dictionary<int, string>();
        }

        #endregion

        #region functions

        /// <summary>
        /// Reads the XML file form the passed stream and processes the workbook information
        /// </summary>
        /// <param name="stream">Stream of the XML file</param>
        /// <exception cref="Exceptions.IOException">Throws IOException in case of an error</exception>
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
                        GetWorkbookInformation(node);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new IOException("XMLStreamException", "The XML entry could not be read from the input stream. Please see the inner exception:", ex);
            }
        }

        /// <summary>
        /// Finds the workbook information recursively
        /// </summary>
        /// <param name="node">Root node to check</param>
        private void GetWorkbookInformation(XmlNode node)
        {
            if (node.LocalName.Equals("sheet", StringComparison.InvariantCultureIgnoreCase))
            {
                try
                {
                    string sheetName = ReaderUtils.GetAttribute("name", node, "worksheet1");
                    int id = int.Parse(ReaderUtils.GetAttribute("sheetId", node), CultureInfo.InvariantCulture); // Default will rightly throw an exception
                    WorksheetDefinitions.Add(id, sheetName);
                }
                catch (Exception e)
                {
                    throw new IOException("XMLStreamException", "The workbook information could not be resolved. Please see the inner exception:", e);
                }
            }

            if (node.HasChildNodes)
            {
                foreach (XmlNode childNode in node.ChildNodes)
                {
                    GetWorkbookInformation(childNode);
                }
            }
        }

        #endregion

    }
}
