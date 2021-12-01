/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2021
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Linq;
using System.IO;
using System.Xml;
using NanoXLSX.Exceptions;
using NanoXLSX.Styles;
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
                    xr.XmlResolver = null;
                    xr.Load(stream);
                    foreach (XmlNode node in xr.DocumentElement.ChildNodes)
                    {
                        if (node.LocalName.Equals("numfmts", StringComparison.InvariantCultureIgnoreCase)) // Handles custom number formats
                        {
                            GetNumberFormats(node);
                        }
                        else if (node.LocalName.Equals("borders", StringComparison.InvariantCultureIgnoreCase)) // Handles borders
                        {
                            GetBorders(node);
                        }
                        // TODO: Implement other style components
                    }
                    foreach (XmlNode node in xr.DocumentElement.ChildNodes) // Redo for composition after all style parts are gathered; standard number formats
                    {
                        if (node.LocalName.Equals("cellxfs", StringComparison.InvariantCultureIgnoreCase))
                        {
                            GetCellXfs(node);
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
        /// Determines the number formats in an XML node of the style document
        /// </summary>
        /// <param name="node">Number formats root node</param>
        /// <exception cref="Exceptions.IOException">Throws IOException in case of an error</exception>
        private void GetNumberFormats(XmlNode node)
        {
            try
            {
                foreach (XmlNode childNode in node.ChildNodes)
                {
                    if (childNode.LocalName.Equals("numfmt", StringComparison.InvariantCultureIgnoreCase))
                    {
                        NumberFormat numberFormat = new NumberFormat();
                        int id = int.Parse(ReaderUtils.GetAttribute("numFmtId", childNode)); // Default will (justified) throw an exception
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
                        numberFormat.InternalID = id;
                        numberFormat.CustomFormatCode = code;
                        StyleReaderContainer.AddStyleComponent(numberFormat);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new IOException("The style information could not be resolved. Please see the inner exception:", ex);
            }
        }

        /// <summary>
        /// Determines the borders in an XML node of the style document
        /// </summary>
        /// <param name="node">Border root node</param>
        private void GetBorders(XmlNode node)
        {
            foreach (XmlNode border in node.ChildNodes)
            {
                Border borderStyle = new Border();
                    string diagonalDown = ReaderUtils.GetAttribute("diagonalDown", border);
                    string diagonalUp = ReaderUtils.GetAttribute("diagonalUp", border);
                    if (diagonalDown == "1")
                    {
                        borderStyle.DiagonalDown = true;
                    }
                    if (diagonalUp == "1")
                    {
                        borderStyle.DiagonalUp = true;
                    }
                    Border.StyleValue styleType;
                    XmlNode innerNode = ReaderUtils.GetChildNode(border, "diagonal");
                    if (innerNode != null)
                    {
                        string styleValue = ReaderUtils.GetAttribute("style", innerNode);
                        if (styleValue != null && Enum.TryParse<Border.StyleValue>(styleValue, out styleType))
                        {
                            borderStyle.DiagonalStyle = styleType;
                        }
                        borderStyle.DiagonalColor = GetColor(innerNode, Border.DEFAULT_COLOR);
                    }
                    innerNode = ReaderUtils.GetChildNode(border, "top");
                    if (innerNode != null)
                    {
                        string styleValue = ReaderUtils.GetAttribute("style", innerNode);
                        if (styleValue != null && Enum.TryParse<Border.StyleValue>(styleValue, out styleType))
                        {
                            borderStyle.TopStyle = styleType;
                        }
                        borderStyle.TopColor = GetColor(innerNode, Border.DEFAULT_COLOR);
                    }
                    innerNode = ReaderUtils.GetChildNode(border, "bottom");
                    if (innerNode != null)
                    {
                        string styleValue = ReaderUtils.GetAttribute("style", innerNode);
                        if (styleValue != null && Enum.TryParse<Border.StyleValue>(styleValue, out styleType))
                        {
                            borderStyle.BottomStyle = styleType;
                        }
                        borderStyle.BottomColor = GetColor(innerNode, Border.DEFAULT_COLOR);
                    }
                    innerNode = ReaderUtils.GetChildNode(border, "left");
                    if (innerNode != null)
                    {
                        string styleValue = ReaderUtils.GetAttribute("style", innerNode);
                        if (styleValue != null && Enum.TryParse<Border.StyleValue>(styleValue, out styleType))
                        {
                            borderStyle.LeftStyle = styleType;
                        }
                        borderStyle.LeftColor = GetColor(innerNode, Border.DEFAULT_COLOR);
                    }
                    innerNode = ReaderUtils.GetChildNode(border, "right");
                    if (innerNode != null)
                    {
                        string styleValue = ReaderUtils.GetAttribute("style", innerNode);
                        if (styleValue != null && Enum.TryParse<Border.StyleValue>(styleValue, out styleType))
                        {
                            borderStyle.RightStyle = styleType;
                        }
                        borderStyle.RightColor = GetColor(innerNode, Border.DEFAULT_COLOR);
                    }
                borderStyle.InternalID = StyleReaderContainer.GetNextBorderId();
                StyleReaderContainer.AddStyleComponent(borderStyle);
            }
        }

        /// <summary>
        /// Determines the cell XF entries in an XML node of the style document
        /// </summary>
        /// <param name="node">Cell XF root node</param>
        private void GetCellXfs(XmlNode node)
        {
            try
            {
                foreach (XmlNode childNode in node.ChildNodes)
                {
                    if (ReaderUtils.IsNode(childNode, "xf"))
                    {
                        Style style = new Style();
                        int id = int.Parse(ReaderUtils.GetAttribute("numFmtId", childNode));
                        NumberFormat format = StyleReaderContainer.GetNumberFormat(id, true);
                        if (format == null)
                        {
                            NumberFormat.FormatNumber formatNumber;
                            NumberFormat.TryParseFormatNumber(id, out formatNumber); // Validity is neglected here to prevent unhandled crashes. If invalid, the format will be declared as 'none'
                            // Invalid values should not occur at all (malformed Excel files). 
                            // Undefined values may occur if the file was saved by an Excel version that has implemented yet unknown format numbers (undefined in NanoXLSX) 
                            format = new NumberFormat();
                            format.Number = formatNumber;
                            format.InternalID = StyleReaderContainer.GetNextNumberFormatId();
                            StyleReaderContainer.AddStyleComponent(format);
                        }
                        id = int.Parse(ReaderUtils.GetAttribute("borderId", childNode));
                        Border border = StyleReaderContainer.GetBorder(id, true);
                        if (border == null)
                        {
                            border = new Border();
                            border.InternalID = StyleReaderContainer.GetNextBorderId();
                        }
                        
                        // TODO: Implement other style information
                        style.CurrentNumberFormat = format;
                        style.CurrentBorder = border;
                        style.InternalID = StyleReaderContainer.GetNextStyleId();

                        StyleReaderContainer.AddStyleComponent(style);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new IOException("The style information could not be resolved. Please see the inner exception:", ex);
            }
        }

        /// <summary>
        /// Resolves a color value from an XML node, when a rgb attribute exists
        /// </summary>
        /// <param name="node">Node to check</param>
        /// <param name="fallback">Fallback value if the color could not be resolved</param>
        /// <returns>RGB value as string or the fallback</returns>
        private static string GetColor(XmlNode node, string fallback)
        {
            XmlNode childNode = ReaderUtils.GetChildNode(node, "color");
            if (childNode != null)
            {
                return ReaderUtils.GetAttribute("rgb", childNode);
            }
            return fallback;
        }

        #endregion
    }
}
