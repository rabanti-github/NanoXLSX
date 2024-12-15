/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2024
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.IO;
using System.Xml;

namespace NanoXLSX.Internal.Readers
{
    /// <summary>
    /// Class representing a reader for the Core metadata file (docProps) embedded in XLSX files
    /// </summary>
    public class MetadataCoreReader : IPluginReader
    {
        #region properties
        /// <summary>
        /// Document category of an XLSX file
        /// </summary>
        public string Category { get; private set; }
        /// <summary>
        /// Content status of an XLSX file
        /// </summary>
        public string ContentStatus { get; private set; }
        /// <summary>
        /// Creator of an XLSX file
        /// </summary>
        public string Creator { get; private set; }
        /// <summary>
        /// Description of the XLSX file
        /// </summary>
        public string Description { get; private set; }
        /// <summary>
        /// Keywords of the XLSX file
        /// </summary>
        public string Keywords { get; private set; }
        /// <summary>
        /// Subject of the XLSX file
        /// </summary>
        public string Subject { get; private set; }
        /// <summary>
        /// Title of the XLSX file
        /// </summary>
        public string Title { get; private set; }
        #endregion

        #region methods
        /// <summary>
        /// Reads the XML file form the passed stream and processes the Core section
        /// </summary>
        /// <param name="stream">Stream of the XML file</param>
        /// \remark <remarks>This method is virtual. Plug-in packages may override it</remarks>
        /// <exception cref="NanoXLSX.Shared.Exceptions.IOException">Throws IOException in case of an error</exception>
        public virtual void Read(MemoryStream stream)
        {
            PreRead(stream);
            if (stream == null)
            {
                // No metadata available in xlsx file
                return;
            }
            try
            {
                using (stream) // Close after processing
                {
                    XmlDocument xr = new XmlDocument();
                    xr.XmlResolver = null;
                    xr.Load(stream);
                    foreach (XmlNode node in xr.DocumentElement.ChildNodes)
                    {
                        if (node.LocalName.Equals("Category", StringComparison.InvariantCultureIgnoreCase))
                        {
                            this.Category = node.InnerText;
                        }
                        else if (node.LocalName.Equals("ContentStatus", StringComparison.InvariantCultureIgnoreCase))
                        {
                            this.ContentStatus = node.InnerText;
                        }
                        else if (node.LocalName.Equals("Creator", StringComparison.InvariantCultureIgnoreCase))
                        {
                            this.Creator = node.InnerText;
                        }
                        else if (node.LocalName.Equals("Description", StringComparison.InvariantCultureIgnoreCase))
                        {
                            this.Description = node.InnerText;
                        }
                        else if (node.LocalName.Equals("Keywords", StringComparison.InvariantCultureIgnoreCase))
                        {
                            this.Keywords = node.InnerText;
                        }
                        else if (node.LocalName.Equals("Subject", StringComparison.InvariantCultureIgnoreCase))
                        {
                            this.Subject = node.InnerText;
                        }
                        else if (node.LocalName.Equals("Title", StringComparison.InvariantCultureIgnoreCase))
                        {
                            this.Title = node.InnerText;
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
