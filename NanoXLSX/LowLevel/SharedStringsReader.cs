/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2021
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using NanoXLSX.Exceptions;
using IOException = NanoXLSX.Exceptions.IOException;

namespace NanoXLSX.LowLevel
{
    /// <summary>
    /// Class representing a reader for the shared strings table of XLSX files
    /// </summary>
    public class SharedStringsReader
    {

        #region properties

        /// <summary>
        /// List of shared string entries
        /// </summary>
        /// <value>
        /// String entry, sorted by its internal index of the table
        /// </value>
        public List<string> SharedStrings { get; private set; }

        /// <summary>
        /// Gets whether the workbook contains shared strings
        /// </summary>
        /// <value>
        ///   True if at least one shared string object exists in the workbook
        /// </value>
        public bool HasElements
        {
            get
            {
                return SharedStrings.Count > 0;
            }
        }

        /// <summary>
        /// Gets the value of the shared string table by its index
        /// </summary>
        /// <param name="index">Index of the stared string entry</param>
        /// <returns>Determined shared string value. Returns null in case of a invalid index</returns>
        public string GetString(int index)
        {
            if (!HasElements || index > SharedStrings.Count - 1)
            {
                return null;
            }
            return SharedStrings[index];
        }
        #endregion

        #region constructors

        /// <summary>
        /// Default constructor
        /// </summary>
        public SharedStringsReader()
        {
            SharedStrings = new List<string>();
        }
        #endregion

        #region methods

        /// <summary>
        /// Reads the XML file form the passed stream and processes the shared strings table
        /// </summary>
        /// <param name="stream">Stream of the XML file</param>
        /// <exception cref="Exceptions.IOException">Throws IOException in case of an error</exception>
        public void Read(Stream stream)
        {
            try
            {
                using (stream) // Close after processing
                {
                    XmlDocument xr = new XmlDocument();
                    xr.XmlResolver = null;
                    xr.Load(stream);
                    StringBuilder sb = new StringBuilder();
                    foreach (XmlNode node in xr.DocumentElement.ChildNodes)
                    {
                        if (node.LocalName.Equals("si", StringComparison.InvariantCultureIgnoreCase))
                        {
                            sb.Clear();
                            GetTextToken(node, ref sb);
                            SharedStrings.Add(sb.ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new IOException("XMLStreamException", "The XML entry could not be read from the " + nameof(stream) +  ". Please see the inner exception:", ex);
            }
        }

        /// <summary>
        /// Function collects text tokens recursively in case of a split by formatting
        /// </summary>
        /// <param name="node">Root node to process</param>
        /// <param name="sb">StringBuilder reference</param>
        private void GetTextToken(XmlNode node, ref StringBuilder sb)
        {
            if (node.LocalName.Equals("t", StringComparison.InvariantCultureIgnoreCase) && !string.IsNullOrEmpty(node.InnerText))
            {
                sb.Append(node.InnerText);
            }
            if (node.HasChildNodes)
            {
                foreach (XmlNode childNode in node.ChildNodes)
                {
                    GetTextToken(childNode, ref sb);
                }
            }
        }

        #endregion

    }
}
