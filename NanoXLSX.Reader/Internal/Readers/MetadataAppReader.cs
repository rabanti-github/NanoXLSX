/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.IO;
using System.Xml;
using NanoXLSX.Interfaces.Reader;

namespace NanoXLSX.Internal.Readers
{
    /// <summary>
    /// Class representing a reader for the App metadata file (docProps) embedded in XLSX files
    /// </summary>
    public class MetadataAppReader : IPluginReader
    {
        #region properties
        /// <summary>
        /// Application that has created an XLSX file. This is an arbitrary text and the default of this library is "NanoXLSX"
        /// </summary>
        public string Application { get; private set; }
        /// <summary>
        /// Version of the application that has created an XLSX file
        /// </summary>
        public string ApplicationVersion { get; private set; }
        /// <summary>
        /// Responsible company of an XLSX file
        /// </summary>
        public string Company { get; private set; }
        /// <summary>
        /// Manager (responsible) of the XLSX file
        /// </summary>
        public string Manager { get; private set; }
        /// <summary>
        /// Hyperlink base of the XLSX file
        /// </summary>
        public string HyperlinkBase { get; private set; }
        #endregion

        #region methods
        /// <summary>
        /// Reads the XML file form the passed stream and processes the AppData section. The existence of the stream should be checked previously
        /// </summary>
        /// \remark <remarks>This method is virtual. Plug-in packages may override it</remarks>
        /// <param name="stream">Stream of the XML file</param>
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
                        if (node.LocalName.Equals("Application", StringComparison.InvariantCultureIgnoreCase))
                        {
                            this.Application = node.InnerText;
                        }
                        else if (node.LocalName.Equals("AppVersion", StringComparison.InvariantCultureIgnoreCase))
                        {
                            this.ApplicationVersion = node.InnerText;
                        }
                        else if (node.LocalName.Equals("Company", StringComparison.InvariantCultureIgnoreCase))
                        {
                            this.Company = node.InnerText;
                        }
                        else if (node.LocalName.Equals("Manager", StringComparison.InvariantCultureIgnoreCase))
                        {
                            this.Manager = node.InnerText;
                        }
                        else if (node.LocalName.Equals("HyperlinkBase", StringComparison.InvariantCultureIgnoreCase))
                        {
                            this.HyperlinkBase = node.InnerText;
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

        #endregion
    }
}
