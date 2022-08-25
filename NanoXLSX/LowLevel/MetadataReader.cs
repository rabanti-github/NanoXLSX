/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2022
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.IO;
using System.Xml;

namespace NanoXLSX.LowLevel
{
    /// <summary>
    /// Class representing a reader for the metadata files (docProps) of XLSX files
    /// </summary>
    public class MetadataReader
    {
        #region properties
        /// <summary>
        /// Application that has created an XLSX file. This is an arbitrary text and the default of this library is "NanoXLSX"
        /// </summary>
        public string Application { get; private set; }
        #endregion

        #region methods
        /// <summary>
        /// Reads the XML file form the passed stream and processes the AppData section
        /// </summary>
        /// <param name="stream">Stream of the XML file</param>
        /// <exception cref="Exceptions.IOException">Throws IOException in case of an error</exception>
        public void ReadAppData(MemoryStream stream)
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
                        if (node.LocalName.Equals("Application", StringComparison.InvariantCultureIgnoreCase))
                        {
                            this.Application = node.InnerText;
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
        /// Reads the XML file form the passed stream and processes the Core section
        /// </summary>
        /// <param name="stream">Stream of the XML file</param>
        /// <exception cref="Exceptions.IOException">Throws IOException in case of an error</exception>
        public void ReadCoreData(MemoryStream stream)
        {
           // throw new NotImplementedException();
        }

        #endregion
    }
}
