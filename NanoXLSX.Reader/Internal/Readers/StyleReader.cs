/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2024
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

namespace NanoXLSX.Internal.Readers
{
    using System;
    using System.IO;
    using System.Xml;
    using NanoXLS.Schemes;
    using NanoXLSX.Shared.Utils;
    using NanoXLSX.Styles;
    using IOException = NanoXLSX.Shared.Exceptions.IOException;

    /// <summary>
    /// Class representing a reader for style definitions of XLSX files.
    /// </summary>
    public class StyleReader : IPluginReader
    {
        /// <summary>
        /// Gets the StyleReaderContainer
        /// Container for raw style components of the reader..
        /// </summary>
        public StyleReaderContainer StyleReaderContainer { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="StyleReader"/> class.
        /// </summary>
        public StyleReader()
        {
            StyleReaderContainer = new StyleReaderContainer();
        }

        /// <summary>
        /// Reads the XML file form the passed stream and processes the style information.
        /// </summary>
        /// <remarks>This method is virtual. Plug-in packages may override it</remarks>
        /// <param name="stream">Stream of the XML file.</param>
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
        /// Determines the number formats in an XML node of the style document.
        /// </summary>
        /// <param name="node">Number formats root name.</param>
        private void GetNumberFormats(XmlNode node)
        {
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName.Equals("numfmt", StringComparison.InvariantCultureIgnoreCase))
                {
                    NumberFormat numberFormat = new NumberFormat();
                    int id = ParserUtils.ParseInt(ReaderUtils.GetAttribute(childNode, "numFmtId")); // Default will (justified) throw an exception
                    string code = ReaderUtils.GetAttribute(childNode, "formatCode", string.Empty);
                    numberFormat.CustomFormatID = id;
                    numberFormat.Number = FormatNumber.custom;
                    numberFormat.InternalID = id;
                    numberFormat.CustomFormatCode = code;
                    StyleReaderContainer.AddStyleComponent(numberFormat);
                }
            }
        }

        /// <summary>
        /// Determines the borders in an XML node of the style document.
        /// </summary>
        /// <param name="node">Border root node.</param>
        private void GetBorders(XmlNode node)
        {
            foreach (XmlNode border in node.ChildNodes)
            {
                Border borderStyle = new Border();
                string diagonalDown = ReaderUtils.GetAttribute(border, "diagonalDown");
                string diagonalUp = ReaderUtils.GetAttribute(border, "diagonalUp");
                if (diagonalDown != null)
                {
                    int value = ParserUtils.ParseBinaryBool(diagonalDown);
                    if (value == 1)
                    {
                        borderStyle.DiagonalDown = true;
                    }
                }
                if (diagonalUp != null)
                {
                    int value = ParserUtils.ParseBinaryBool(diagonalUp);
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
        /// Tries to parse a border style.
        /// </summary>
        /// <param name="innerNode">Border sub-node.</param>
        /// <returns>Border type or none if parsing was not successful.</returns>
        private static StyleValue ParseBorderStyle(XmlNode innerNode)
        {
            string value = ReaderUtils.GetAttribute(innerNode, "style");
            if (value != null)
            {
                if (value.Equals("double", StringComparison.InvariantCultureIgnoreCase))
                {
                    return StyleValue.s_double; // special handling, since double is not a valid enum value
                }
                StyleValue styleType;
                if (Enum.TryParse(value, out styleType))
                {
                    return styleType;
                }
            }
            return StyleValue.none;
        }

        /// <summary>
        /// Determines the fills in an XML node of the style document.
        /// </summary>
        /// <param name="node">Fill root node.</param>
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
                    PatternValue patternValue;
                    if (Enum.TryParse<PatternValue>(pattern, out patternValue))
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
                            fillStyle.IndexedColor = ParserUtils.ParseInt(backgroundIndex);
                        }
                    }
                }

                fillStyle.InternalID = StyleReaderContainer.GetNextFillId();
                StyleReaderContainer.AddStyleComponent(fillStyle);
            }
        }

        /// <summary>
        /// Determines the fonts in an XML node of the style document.
        /// </summary>
        /// <param name="node">Font root node.</param>
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
                    fontStyle.Underline = UnderlineValue.u_single; // default
                    switch (attribute)
                    {
                        case "double":
                            fontStyle.Underline = UnderlineValue.u_double;
                            break;
                        case "singleAccounting":
                            fontStyle.Underline = UnderlineValue.singleAccounting;
                            break;
                        case "doubleAccounting":
                            fontStyle.Underline = UnderlineValue.doubleAccounting;
                            break;
                    }
                }
                if (ReaderUtils.GetAttributeOfChild(font, "vertAlign", "val", out attribute))
                {
                    VerticalTextAlignValue vertAlignValue;
                    if (Enum.TryParse<VerticalTextAlignValue>(attribute, out vertAlignValue))
                    {
                        fontStyle.VerticalAlign = vertAlignValue;
                    }
                }
                if (ReaderUtils.GetAttributeOfChild(font, "sz", "val", out attribute))
                {
                    fontStyle.Size = ParserUtils.ParseFloat(attribute);
                }
                XmlNode colorNode = ReaderUtils.GetChildNode(font, "color");
                if (colorNode != null)
                {

                    attribute = ReaderUtils.GetAttribute(colorNode, "theme");
                    if (attribute != null)
                    {
                        switch (attribute)
                        {
                            case "0":
                                fontStyle.ColorTheme = ColorSchemeElement.dark1;
                                break;
                            case "1":
                                fontStyle.ColorTheme = ColorSchemeElement.light1;
                                break;
                            case "2":
                                fontStyle.ColorTheme = ColorSchemeElement.dark2;
                                break;
                            case "3":
                                fontStyle.ColorTheme = ColorSchemeElement.light2;
                                break;
                            case "4":
                                fontStyle.ColorTheme = ColorSchemeElement.accent1;
                                break;
                            case "5":
                                fontStyle.ColorTheme = ColorSchemeElement.accent2;
                                break;
                            case "6":
                                fontStyle.ColorTheme = ColorSchemeElement.accent3;
                                break;
                            case "7":
                                fontStyle.ColorTheme = ColorSchemeElement.accent4;
                                break;
                            case "8":
                                fontStyle.ColorTheme = ColorSchemeElement.accent5;
                                break;
                            case "9":
                                fontStyle.ColorTheme = ColorSchemeElement.accent6;
                                break;
                            case "10":
                                fontStyle.ColorTheme = ColorSchemeElement.hyperlink;
                                break;
                            case "11":
                                fontStyle.ColorTheme = ColorSchemeElement.followedHyperlink;
                                break;
                        }
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
                    switch (attribute)
                    {

                        case "0":
                            fontStyle.Family = FontFamilyValue.NotApplicable;
                            break;
                        case "1":
                            fontStyle.Family = FontFamilyValue.Roman;
                            break;
                        case "2":
                            fontStyle.Family = FontFamilyValue.Swiss;
                            break;
                        case "3":
                            fontStyle.Family = FontFamilyValue.Modern;
                            break;
                        case "4":
                            fontStyle.Family = FontFamilyValue.Script;
                            break;
                        case "5":
                            fontStyle.Family = FontFamilyValue.Decorative;
                            break;
                        case "6":
                            fontStyle.Family = FontFamilyValue.Reserved1;
                            break;
                        case "7":
                            fontStyle.Family = FontFamilyValue.Reserved2;
                            break;
                        case "8":
                            fontStyle.Family = FontFamilyValue.Reserved3;
                            break;
                        case "9":
                            fontStyle.Family = FontFamilyValue.Reserved4;
                            break;
                        case "10":
                            fontStyle.Family = FontFamilyValue.Reserved5;
                            break;
                        case "11":
                            fontStyle.Family = FontFamilyValue.Reserved6;
                            break;
                        case "12":
                            fontStyle.Family = FontFamilyValue.Reserved7;
                            break;
                        case "13":
                            fontStyle.Family = FontFamilyValue.Reserved8;
                            break;
                        case "14":
                            fontStyle.Family = FontFamilyValue.Reserved9;
                            break;
                    }
                }
                if (ReaderUtils.GetAttributeOfChild(font, "scheme", "val", out attribute))
                {
                    switch (attribute)
                    {
                        case "major":
                            fontStyle.Scheme = SchemeValue.major;
                            break;
                        case "minor":
                            fontStyle.Scheme = SchemeValue.minor;
                            break;
                    }
                }
                if (ReaderUtils.GetAttributeOfChild(font, "charset", "val", out attribute))
                {
                    switch (attribute)
                    {
                        case "0":
                            fontStyle.Charset = CharsetValue.ANSI;
                            break;
                        case "1":
                            fontStyle.Charset = CharsetValue.Default;
                            break;
                        case "2":
                            fontStyle.Charset = CharsetValue.Symbols;
                            break;
                        case "77":
                            fontStyle.Charset = CharsetValue.Macintosh;
                            break;
                        case "128":
                            fontStyle.Charset = CharsetValue.JIS;
                            break;
                        case "129":
                            fontStyle.Charset = CharsetValue.Hangul;
                            break;
                        case "130":
                            fontStyle.Charset = CharsetValue.Johab;
                            break;
                        case "134":
                            fontStyle.Charset = CharsetValue.GKB;
                            break;
                        case "136":
                            fontStyle.Charset = CharsetValue.Big5;
                            break;
                        case "161":
                            fontStyle.Charset = CharsetValue.Greek;
                            break;
                        case "162":
                            fontStyle.Charset = CharsetValue.Turkish;
                            break;
                        case "163":
                            fontStyle.Charset = CharsetValue.Vietnamese;
                            break;
                        case "177":
                            fontStyle.Charset = CharsetValue.Hebrew;
                            break;
                        case "178":
                            fontStyle.Charset = CharsetValue.Arabic;
                            break;
                        case "186":
                            fontStyle.Charset = CharsetValue.Baltic;
                            break;
                        case "204":
                            fontStyle.Charset = CharsetValue.Russian;
                            break;
                        case "222":
                            fontStyle.Charset = CharsetValue.Thai;
                            break;
                        case "238":
                            fontStyle.Charset = CharsetValue.EasternEuropean;
                            break;
                        case "255":
                            fontStyle.Charset = CharsetValue.OEM;
                            break;
                        default:
                            fontStyle.Charset = CharsetValue.ApplicationDefined;
                            break;
                    }
                }

                fontStyle.InternalID = StyleReaderContainer.GetNextFontId();
                StyleReaderContainer.AddStyleComponent(fontStyle);
            }
        }

        /// <summary>
        /// Determines the cell XF entries in an XML node of the style document.
        /// </summary>
        /// <param name="node">Cell XF root node.</param>
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
                        int value = ParserUtils.ParseBinaryBool(attribute);
                        cellXfStyle.ForceApplyAlignment = value == 1;
                    }
                    XmlNode alignmentNode = ReaderUtils.GetChildNode(childNode, "alignment");
                    if (alignmentNode != null)
                    {
                        attribute = ReaderUtils.GetAttribute(alignmentNode, "shrinkToFit");
                        if (attribute != null)
                        {
                            int value = ParserUtils.ParseBinaryBool(attribute);
                            if (value == 1)
                            {
                                cellXfStyle.Alignment = TextBreakValue.shrinkToFit;
                            }
                        }
                        attribute = ReaderUtils.GetAttribute(alignmentNode, "wrapText");
                        if (attribute != null && attribute == "1")
                        {
                            cellXfStyle.Alignment = TextBreakValue.wrapText;
                        }
                        attribute = ReaderUtils.GetAttribute(alignmentNode, "horizontal");
                        HorizontalAlignValue horizontalAlignValue;
                        if (Enum.TryParse<HorizontalAlignValue>(attribute, out horizontalAlignValue))
                        {
                            cellXfStyle.HorizontalAlign = horizontalAlignValue;
                        }
                        attribute = ReaderUtils.GetAttribute(alignmentNode, "vertical");
                        VerticalAlignValue verticalAlignValue;
                        if (Enum.TryParse<VerticalAlignValue>(attribute, out verticalAlignValue))
                        {
                            cellXfStyle.VerticalAlign = verticalAlignValue;
                        }
                        attribute = ReaderUtils.GetAttribute(alignmentNode, "indent");
                        if (attribute != null)
                        {
                            cellXfStyle.Indent = ParserUtils.ParseInt(attribute);
                        }
                        attribute = ReaderUtils.GetAttribute(alignmentNode, "textRotation");
                        if (attribute != null)
                        {
                            int rotation = ParserUtils.ParseInt(attribute);
                            cellXfStyle.TextRotation = rotation > 90 ? 90 - rotation : rotation;
                        }
                    }
                    XmlNode protectionNode = ReaderUtils.GetChildNode(childNode, "protection");
                    if (protectionNode != null)
                    {
                        attribute = ReaderUtils.GetAttribute(protectionNode, "hidden");
                        if (attribute != null && attribute == "1")
                        {
                            cellXfStyle.Hidden = true;
                        }
                        attribute = ReaderUtils.GetAttribute(protectionNode, "locked");
                        if (attribute != null && attribute == "1")
                        {
                            cellXfStyle.Locked = true;
                        }
                    }

                    cellXfStyle.InternalID = StyleReaderContainer.GetNextCellXFId();
                    StyleReaderContainer.AddStyleComponent(cellXfStyle);

                    Style style = new Style();
                    int id;
                    bool hasId;

                    hasId = ParserUtils.TryParseInt(ReaderUtils.GetAttribute(childNode, "numFmtId"), out id);
                    NumberFormat format = StyleReaderContainer.GetNumberFormat(id);
                    if (!hasId || format == null)
                    {
                        FormatNumber formatNumber;
                        NumberFormat.TryParseFormatNumber(id, out formatNumber); // Validity is neglected here to prevent unhandled crashes. If invalid, the format will be declared as 'none'
                                                                                 // Invalid values should not occur at all (malformed Excel files). 
                                                                                 // Undefined values may occur if the file was saved by an Excel version that has implemented yet unknown format numbers (undefined in NanoXLSX) 
                        format = new NumberFormat();
                        format.Number = formatNumber;
                        format.InternalID = id;
                        StyleReaderContainer.AddStyleComponent(format);
                    }
                    hasId = ParserUtils.TryParseInt(ReaderUtils.GetAttribute(childNode, "borderId"), out id);
                    Border border = StyleReaderContainer.GetBorder(id);
                    if (!hasId || border == null)
                    {
                        border = new Border();
                        border.InternalID = StyleReaderContainer.GetNextBorderId();
                    }
                    hasId = ParserUtils.TryParseInt(ReaderUtils.GetAttribute(childNode, "fillId"), out id);
                    Fill fill = StyleReaderContainer.GetFill(id);
                    if (!hasId || fill == null)
                    {
                        fill = new Fill();
                        fill.InternalID = StyleReaderContainer.GetNextFillId();
                    }
                    hasId = ParserUtils.TryParseInt(ReaderUtils.GetAttribute(childNode, "fontId"), out id);
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
        /// Determines the MRU colors in an XML node of the style document.
        /// </summary>
        /// <param name="node">Color root node.</param>
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
        /// Resolves a color value from an XML node, when a rgb attribute exists.
        /// </summary>
        /// <param name="node">Node to check.</param>
        /// <param name="fallback">Fallback value if the color could not be resolved.</param>
        /// <returns>RGB value as string or the fallback.</returns>
        private static string GetColor(XmlNode node, string fallback)
        {
            XmlNode childNode = ReaderUtils.GetChildNode(node, "color");
            if (childNode != null)
            {
                return ReaderUtils.GetAttribute(childNode, "rgb");
            }
            return fallback;
        }
    }
}
