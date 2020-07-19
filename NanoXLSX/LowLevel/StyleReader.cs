/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2020
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using NanoXLSX.Exceptions;
using Styles;
using IOException = NanoXLSX.Exceptions.IOException;

namespace NanoXLSX.LowLevel
{
    /// <summary>
    /// Class representing a reader for style definitions of XLSX files
    /// </summary>
    public class StyleReader
    {

        #region properties

        /// <summary>
        /// Container for raw style components of the reader. 
        /// </summary>
        public StyleReaderContainer StyleReaderContainer { get; set; }

        #endregion

        #region constructors

        /// <summary>
        /// Default constructor
        /// </summary>
        public StyleReader()
        {
            StyleReaderContainer = new StyleReaderContainer();
        }
        #endregion

        #region functions
        /// <summary>
        /// Reads the XML file form the passed stream and processes the style information
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
                    xr.Load(stream);
                    XmlNodeList nodes = xr.DocumentElement.ChildNodes;

                    foreach (XmlNode node in xr.DocumentElement.ChildNodes)
                    {
                        if (node.LocalName.ToLower() == "numfmts")
                        {
                            GetGetNumberFormats(node);
                        }
                        // TODO: Implement other style components
                    }
                    foreach (XmlNode node in xr.DocumentElement.ChildNodes) // Redo for composition after all style parts are gathered
                    {
                         if (node.LocalName.ToLower() == "cellxfs")
                        {
                            GetGetCellXfs(node);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new IOException("XMLStreamException", "The XML entry could not be read from the input stream. Please see the inner exception:", ex);
            }
        }

        private void GetGetNumberFormats(XmlNode node)
        {
            try
            {
                foreach (XmlNode childNode in node.ChildNodes)
                {
                    if (childNode.LocalName.ToLower() == "numfmt")
                    {
                        NumberFormat numberFormat = new NumberFormat();
                        int id = int.Parse(ReaderUtils.GetAttribute("numFmtId", childNode)); // Default will rightly throw an exception
                        string code = ReaderUtils.GetAttribute("formatCode", childNode, string.Empty);

                        if (id < NumberFormat.CUSTOMFORMAT_START_NUMBER)
                        {
                            if (Enum.IsDefined(typeof(NumberFormat.FormatNumber), id))
                            {
                                numberFormat.Number = (NumberFormat.FormatNumber)Enum.ToObject(typeof(NumberFormat.FormatNumber), id);
                            }
                            else
                            {
                                numberFormat.CustomFormatID = id;
                                numberFormat.Number = NumberFormat.FormatNumber.custom;
                            }

                        }
                        else
                        {
                            numberFormat.CustomFormatID = id;
                            numberFormat.Number = NumberFormat.FormatNumber.custom;
                        }
                        numberFormat.InternalID = StyleReaderContainer.GetNextNumberFormatId();
                        numberFormat.CustomFormatCode = code;
                        StyleReaderContainer.AddStyleComponent(numberFormat);
                    }
                }
            }
            catch(Exception ex)
            {
                throw new IOException("XMLStreamException", "The style information could not be resolved. Please see the inner exception:", ex);
            }
        }

        private void GetGetCellXfs(XmlNode node)
        {
            try
            {
                foreach (XmlNode childNode in node.ChildNodes)
                {


                    if (childNode.LocalName.ToLower() == "xf")
                    {
                        Style style = new Style();
                        // Default (null) of any reference value of an attribute will rightly throw an exception
                        int id = int.Parse(ReaderUtils.GetAttribute("numFmtId", childNode));
                        NumberFormat format = StyleReaderContainer.GetNumberFormat(id, true);
                        if (format == null)
                        {
                            // TODO: What here?
                        }
                        else
                        {
                            style.CurrentNumberFormat = format;
                        }
                        // TODO: Implement other style information
                        StyleReaderContainer.AddStyleComponent(style);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new IOException("XMLStreamException", "The style information could not be resolved. Please see the inner exception:", ex);
            }

        }


            #endregion
        }
}
