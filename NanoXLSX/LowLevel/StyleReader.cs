/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2024
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
        public StyleReaderContainer StyleReaderContainer { get; private set; }

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
                        else if (node.LocalName.Equals("fills", StringComparison.InvariantCultureIgnoreCase)) // Handles fills
                        {
                            GetFills(node);
                        }
                        else if (node.LocalName.Equals("fonts", StringComparison.InvariantCultureIgnoreCase)) // Handles fonts
                        {
                            GetFonts(node);
                        }
                        else if (node.LocalName.Equals("colors", StringComparison.InvariantCultureIgnoreCase)) // Handles MRU colors
                        {
                            GetColors(node);
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
        /// <param name="node">Number formats root name</param>
        private void GetNumberFormats(XmlNode node)
        {
                foreach (XmlNode childNode in node.ChildNodes)
                {
                    if (childNode.LocalName.Equals("numfmt", StringComparison.InvariantCultureIgnoreCase))
                    {
                        NumberFormat numberFormat = new NumberFormat();
                        int id = ReaderUtils.ParseInt(ReaderUtils.GetAttribute(childNode, "numFmtId")); // Default will (justified) throw an exception
                        string code = ReaderUtils.GetAttribute(childNode, "formatCode", string.Empty); // Code is not un-escaped
                        numberFormat.CustomFormatID = id;
                        numberFormat.Number = NumberFormat.FormatNumber.custom;
                        numberFormat.InternalID = id;
                        numberFormat.CustomFormatCode = code;
                        StyleReaderContainer.AddStyleComponent(numberFormat);
                    }
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
                string diagonalDown = ReaderUtils.GetAttribute(border, "diagonalDown");
                string diagonalUp = ReaderUtils.GetAttribute(border, "diagonalUp");
                if (diagonalDown != null)
                {
                    int value = ReaderUtils.ParseBinaryBool(diagonalDown);
                    if (value == 1)
                    {
                        borderStyle.DiagonalDown = true;
                    }
                }
                if (diagonalUp != null)
                {
                    int value = ReaderUtils.ParseBinaryBool(diagonalUp);
                    if (value == 1)
                    {
                        borderStyle.DiagonalUp = true;
                    }
                }
                XmlNode innerNode = ReaderUtils.GetChildNode(border, "diagonal");
                if (innerNode != null)
                {
                    borderStyle.DiagonalStyle = ParseBorderStyle(innerNode);
                    borderStyle.DiagonalColor = GetColor(innerNode, Border.DEFAULT_COLOR);
                }
                innerNode = ReaderUtils.GetChildNode(border, "top");
                if (innerNode != null)
                {
                    borderStyle.TopStyle = ParseBorderStyle(innerNode);
                    borderStyle.TopColor = GetColor(innerNode, Border.DEFAULT_COLOR);
                }
                innerNode = ReaderUtils.GetChildNode(border, "bottom");
                if (innerNode != null)
                {
                    borderStyle.BottomStyle = ParseBorderStyle(innerNode);
                    borderStyle.BottomColor = GetColor(innerNode, Border.DEFAULT_COLOR);
                }
                innerNode = ReaderUtils.GetChildNode(border, "left");
                if (innerNode != null)
                {
                    borderStyle.LeftStyle = ParseBorderStyle(innerNode);
                    borderStyle.LeftColor = GetColor(innerNode, Border.DEFAULT_COLOR);
                }
                innerNode = ReaderUtils.GetChildNode(border, "right");
                if (innerNode != null)
                {
                    borderStyle.RightStyle = ParseBorderStyle(innerNode);
                    borderStyle.RightColor = GetColor(innerNode, Border.DEFAULT_COLOR);
                }
                borderStyle.InternalID = StyleReaderContainer.GetNextBorderId();
                StyleReaderContainer.AddStyleComponent(borderStyle);
                }
        }

            /// <summary>
            /// Tries to parse a border style
            /// </summary>
            /// <param name="innerNode">Border sub-node</param>
            /// <returns>Border type or non if parsing was not successful</returns>
         private static Border.StyleValue ParseBorderStyle(XmlNode innerNode)
        {
            string value = ReaderUtils.GetAttribute(innerNode, "style");
            if (value != null)
            {
                if (value.Equals("double", StringComparison.InvariantCultureIgnoreCase))
                {
                    return Border.StyleValue.s_double; // special handling, since double is not a valid enum value
                }
                Border.StyleValue styleType;
                if (Enum.TryParse(value, out styleType))
                {
                    return styleType;
                }
            }
            return Border.StyleValue.none;
        }

        /// <summary>
        /// Determines the fills in an XML node of the style document
        /// </summary>
        /// <param name="node">Fill root node</param>
        private void GetFills(XmlNode node)
        {
            string attribute;
            foreach (XmlNode fill in node.ChildNodes)
            {
                Fill fillStyle = new Fill();
                XmlNode innerNode = ReaderUtils.GetChildNode(fill, "patternFill");
                if (innerNode != null)
                {
                    string pattern = ReaderUtils.GetAttribute(innerNode, "patternType");
                    Fill.PatternValue patternValue;
                    if (Enum.TryParse<Fill.PatternValue>(pattern, out patternValue))
                    {
                        fillStyle.PatternFill = patternValue;
                    }
                    if (ReaderUtils.GetAttributeOfChild(innerNode, "fgColor", "rgb", out attribute))
                    {
                        if (!string.IsNullOrEmpty(attribute))
                        {
                            fillStyle.ForegroundColor = attribute;
                        }
                    }
                    XmlNode backgroundNode = ReaderUtils.GetChildNode(innerNode, "bgColor");
                    if (backgroundNode != null)
                    {
                        string backgroundArgb = ReaderUtils.GetAttribute(backgroundNode, "rgb");
                        if (!string.IsNullOrEmpty(backgroundArgb))
                        {
                            fillStyle.BackgroundColor = backgroundArgb;
                        }
                        string backgroundIndex = ReaderUtils.GetAttribute(backgroundNode, "indexed");
                        if (!string.IsNullOrEmpty(backgroundIndex))
                        {
                            fillStyle.IndexedColor = ReaderUtils.ParseInt(backgroundIndex);
                        }
                    }
                }

                fillStyle.InternalID = StyleReaderContainer.GetNextFillId();
                StyleReaderContainer.AddStyleComponent(fillStyle);
            }
        }

        /// <summary>
        /// Determines the fonts in an XML node of the style document
        /// </summary>
        /// <param name="node">Font root node</param>
        private void GetFonts(XmlNode node)
        {
            string attribute;
            foreach (XmlNode font in node.ChildNodes)
            {
                Font fontStyle = new Font();
                XmlNode boldNode = ReaderUtils.GetChildNode(font, "b");
                if (boldNode != null)
                {
                    fontStyle.Bold = true;
                }
                XmlNode italicdNode = ReaderUtils.GetChildNode(font, "i");
                if (italicdNode != null)
                {
                    fontStyle.Italic = true;
                }
                XmlNode strikeNode = ReaderUtils.GetChildNode(font, "strike");
                if (strikeNode != null)
                {
                    fontStyle.Strike = true;
                }
                if (ReaderUtils.GetAttributeOfChild(font, "u", "val", out attribute))
                {
                    fontStyle.Underline = Font.UnderlineValue.u_single; // default
                    switch (attribute)
                    {
                        case "double":
                            fontStyle.Underline = Font.UnderlineValue.u_double;
                            break;
                        case "singleAccounting":
                            fontStyle.Underline = Font.UnderlineValue.singleAccounting;
                            break;
                        case "doubleAccounting":
                            fontStyle.Underline = Font.UnderlineValue.doubleAccounting;
                            break;
                    }
                }
                if (ReaderUtils.GetAttributeOfChild(font, "vertAlign", "val", out attribute))
                {
                    Font.VerticalAlignValue vertAlignValue;
                    if (Enum.TryParse<Font.VerticalAlignValue>(attribute, out vertAlignValue))
                    {
                        fontStyle.VerticalAlign = vertAlignValue;
                    }
                }
                if (ReaderUtils.GetAttributeOfChild(font, "sz", "val", out attribute))
                {
                    fontStyle.Size = ReaderUtils.ParseFloat(attribute);
                }
                XmlNode colorNode = ReaderUtils.GetChildNode(font, "color");
                if (colorNode != null)
                {
                    attribute = ReaderUtils.GetAttribute(colorNode, "theme");
                    if (attribute != null)
                    {
                        fontStyle.ColorTheme = ReaderUtils.ParseInt(attribute);
                    }
                    attribute = ReaderUtils.GetAttribute(colorNode, "rgb");
                   if (attribute != null)
                    {
                        fontStyle.ColorValue = attribute;
                    }
                }
                if (ReaderUtils.GetAttributeOfChild(font, "name", "val", out attribute))
                {
                    fontStyle.Name = attribute;
                }
                if (ReaderUtils.GetAttributeOfChild(font, "family", "val", out attribute))
                {
                    fontStyle.Family = attribute;
                }
                if (ReaderUtils.GetAttributeOfChild(font, "scheme", "val", out attribute))
                {
                    switch (attribute)
                    {
                        case "major":
                            fontStyle.Scheme = Font.SchemeValue.major;
                            break;
                        case "minor":
                            fontStyle.Scheme = Font.SchemeValue.minor;
                            break;
                    }
                }
                if (ReaderUtils.GetAttributeOfChild(font, "charset", "val", out attribute))
                {
                    fontStyle.Charset = attribute;
                }

                fontStyle.InternalID = StyleReaderContainer.GetNextFontId();
                StyleReaderContainer.AddStyleComponent(fontStyle);
            }
        }


        /// <summary>
        /// Determines the cell XF entries in an XML node of the style document
        /// </summary>
        /// <param name="node">Cell XF root node</param>
        private void GetCellXfs(XmlNode node)
        {
                foreach (XmlNode childNode in node.ChildNodes)
                {
                    if (ReaderUtils.IsNode(childNode, "xf"))
                    {
                    CellXf cellXfStyle = new CellXf();
                    string attribute = ReaderUtils.GetAttribute(childNode, "applyAlignment");
                    if (attribute != null)
                    {
                        int value = ReaderUtils.ParseBinaryBool(attribute);
                        cellXfStyle.ForceApplyAlignment = value == 1;
                    }
                    XmlNode alignmentNode = ReaderUtils.GetChildNode(childNode, "alignment");
                    if (alignmentNode != null)
                    {
                        attribute = ReaderUtils.GetAttribute(alignmentNode, "shrinkToFit");
                        if (attribute != null)
                        {
                            int value = ReaderUtils.ParseBinaryBool(attribute);
                            if (value == 1)
                            {
                                cellXfStyle.Alignment = CellXf.TextBreakValue.shrinkToFit;
                            }
                                
                        }
                        attribute = ReaderUtils.GetAttribute(alignmentNode, "wrapText");
                        if (attribute != null)
                        {
                            int value = ReaderUtils.ParseBinaryBool(attribute);
                            if (value == 1)
                            {
                                cellXfStyle.Alignment = CellXf.TextBreakValue.wrapText;
                            } 
                        }
                        attribute = ReaderUtils.GetAttribute(alignmentNode, "horizontal");
                        CellXf.HorizontalAlignValue horizontalAlignValue;
                        if (Enum.TryParse<CellXf.HorizontalAlignValue>(attribute, out horizontalAlignValue))
                        {
                            cellXfStyle.HorizontalAlign = horizontalAlignValue;
                        }
                        attribute = ReaderUtils.GetAttribute(alignmentNode, "vertical");
                        CellXf.VerticalAlignValue verticalAlignValue;
                        if (Enum.TryParse<CellXf.VerticalAlignValue>(attribute, out verticalAlignValue))
                        {
                            cellXfStyle.VerticalAlign = verticalAlignValue;
                        }
                        attribute = ReaderUtils.GetAttribute(alignmentNode, "indent");
                        if (attribute != null)
                        {
                            cellXfStyle.Indent = ReaderUtils.ParseInt(attribute);
                        }
                        attribute = ReaderUtils.GetAttribute(alignmentNode, "textRotation");
                        if (attribute != null)
                        {
                            int rotation = ReaderUtils.ParseInt(attribute);
                            cellXfStyle.TextRotation = rotation > 90 ? 90 - rotation : rotation;
                        }
                    }
                    XmlNode protectionNode = ReaderUtils.GetChildNode(childNode, "protection");
                    if (protectionNode != null)
                    {
                        attribute = ReaderUtils.GetAttribute(protectionNode, "hidden");
                        if (attribute != null)
                        {
                            int value = ReaderUtils.ParseBinaryBool(attribute);
                            if (value == 1)
                            {
                                cellXfStyle.Hidden = true;
                            }
                        }
                        attribute = ReaderUtils.GetAttribute(protectionNode, "locked");
                        if (attribute != null)
                        {
                            int value = ReaderUtils.ParseBinaryBool(attribute);
                            if (value  == 1)
                            {
                                cellXfStyle.Locked = true;
                            }
                        }
                    }

                    cellXfStyle.InternalID = StyleReaderContainer.GetNextCellXFId();
                    StyleReaderContainer.AddStyleComponent(cellXfStyle);

                    Style style = new Style();
                    int id = 0;
                    bool hasId;

                    hasId = ReaderUtils.TryParseInt(ReaderUtils.GetAttribute(childNode, "numFmtId"), out id);
                    NumberFormat format = StyleReaderContainer.GetNumberFormat(id);
                    if (!hasId || format == null)
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
                    hasId = ReaderUtils.TryParseInt(ReaderUtils.GetAttribute(childNode, "borderId"), out id);
                    Border border = StyleReaderContainer.GetBorder(id);
                    if (!hasId || border == null)
                    {
                        border = new Border();
                        border.InternalID = StyleReaderContainer.GetNextBorderId();
                    }
                    hasId = ReaderUtils.TryParseInt(ReaderUtils.GetAttribute(childNode, "fillId"), out id);
                    Fill fill = StyleReaderContainer.GetFill(id);
                    if (!hasId || fill == null)
                    {
                        fill = new Fill();
                        fill.InternalID = StyleReaderContainer.GetNextFillId();
                    }
                    hasId = ReaderUtils.TryParseInt(ReaderUtils.GetAttribute(childNode, "fontId"), out id);
                    Font font = StyleReaderContainer.GetFont(id);
                    if (!hasId || font == null)
                    {
                        font = new Font();
                        font.InternalID = StyleReaderContainer.GetNextFontId();
                    }

                    // TODO: Implement other style information
                    style.CurrentNumberFormat = format;
                    style.CurrentBorder = border;
                    style.CurrentFill = fill;
                    style.CurrentFont = font;
                    style.CurrentCellXf = cellXfStyle;
                    style.InternalID = StyleReaderContainer.GetNextStyleId();

                    StyleReaderContainer.AddStyleComponent(style);
                    }
                }
        }

        /// <summary>
        /// Determines the MRU colors in an XML node of the style document
        /// </summary>
        /// <param name="node">Color root node</param>
        private void GetColors(XmlNode node)
        {
            foreach (XmlNode color in node.ChildNodes)
            {
                XmlNode mruColor = ReaderUtils.GetChildNode(color, "color");
                if (color.Name.Equals("mruColors") && mruColor != null)
                {
                    foreach (XmlNode value in color.ChildNodes)
                    {
                        string attribute = ReaderUtils.GetAttribute(value, "rgb");
                        if (attribute != null)
                        {
                            StyleReaderContainer.AddMruColor(attribute);
                        }
                   }
                }
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
                return ReaderUtils.GetAttribute(childNode, "rgb");
            }
            return fallback;
        }

        #endregion
    }
}
