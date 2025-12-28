/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

namespace NanoXLSX.Internal.Readers
{
    using System;
    using System.IO;
    using System.Xml;
    using NanoXLSX.Colors;
    using NanoXLSX.Interfaces.Plugin;
    using NanoXLSX.Interfaces.Reader;
    using NanoXLSX.Registry;
    using NanoXLSX.Registry.Attributes;
    using NanoXLSX.Styles;
    using NanoXLSX.Themes;
    using NanoXLSX.Utils;
    using static NanoXLSX.Styles.Border;
    using static NanoXLSX.Styles.CellXf;
    using static NanoXLSX.Styles.Font;
    using static NanoXLSX.Styles.NumberFormat;
    using static NanoXLSX.Themes.Theme;
    using IOException = Exceptions.IOException;

    /// <summary>
    /// Class representing a reader for style definitions of XLSX files.
    /// </summary>
    [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.StyleReader)]
    public class StyleReader : IPlugInReader
    {

        private MemoryStream stream;
        private StyleReaderContainer styleReaderContainer;

        #region properties
        /// <summary>
        /// Workbook reference where read data is stored (should not be null)
        /// </summary>
        public Workbook Workbook { get; set; }
        #endregion

        #region constructors
        /// <summary>
        /// Initializes a new instance of the <see cref="StyleReader"/> class.
        /// </summary>
        public StyleReader()
        {
        }
        #endregion

        #region methods
        /// <summary>
        /// Default constructor - Must be defined for instantiation of the plug-ins
        /// </summary>
        public void Init(MemoryStream stream, Workbook workbook, IOptions readerOptions)
        {
            this.stream = stream;
            this.Workbook = workbook;
        }

        /// <summary>
        /// Method to execute the main logic of the plug-in (interface implementation)
        /// </summary>
        /// <exception cref="Exceptions.IOException">Throws an IOException in case of a error during reading</exception>
        public void Execute()
        {
            this.styleReaderContainer = new StyleReaderContainer();
            try
            {
                using (stream) // Close after processing
                {
                    XmlDocument xr = new XmlDocument() { XmlResolver = null };
                    using (XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings() { XmlResolver = null }))
                    {
                        xr.Load(reader);
                        foreach (XmlNode node in xr.DocumentElement.ChildNodes)
                        {
                            if (node.LocalName.Equals("numfmts", StringComparison.OrdinalIgnoreCase)) // Handles custom number formats
                            {
                                GetNumberFormats(node);
                            }
                            else if (node.LocalName.Equals("borders", StringComparison.OrdinalIgnoreCase)) // Handles borders
                            {
                                GetBorders(node);
                            }
                            else if (node.LocalName.Equals("fills", StringComparison.OrdinalIgnoreCase)) // Handles fills
                            {
                                GetFills(node);
                            }
                            else if (node.LocalName.Equals("fonts", StringComparison.OrdinalIgnoreCase)) // Handles fonts
                            {
                                GetFonts(node);
                            }
                            else if (node.LocalName.Equals("colors", StringComparison.OrdinalIgnoreCase)) // Handles MRU colors
                            {
                                GetColors(node);
                            }
                            // TODO: Implement other style components
                        }
                        foreach (XmlNode node in xr.DocumentElement.ChildNodes) // Redo for composition after all style parts are gathered; standard number formats
                        {
                            if (node.LocalName.Equals("cellxfs", StringComparison.OrdinalIgnoreCase))
                            {
                                GetCellXfs(node);
                            }
                        }
                        HandleMruColors();
                        RederPlugInHandler.HandleInlineQueuePlugins(ref stream, Workbook, PlugInUUID.StyleInlineReader);
                    }
                }
                Workbook.AuxiliaryData.SetData(PlugInUUID.StyleReader, PlugInUUID.StyleEntity, styleReaderContainer);
            }
            catch (Exception ex)
            {
                throw new IOException("The XML entry could not be read from the input stream. Please see the inner exception:", ex);
            }
        }

        /// <summary>
        /// handles MRU colors, if defined
        /// </summary>
        private void HandleMruColors()
        {
            if (styleReaderContainer.GetMruColors().Count > 0)
            {
                foreach (string color in styleReaderContainer.GetMruColors())
                {
                    Workbook.AddMruColor(color);
                }
            }
        }

        /// <summary>
        /// Determines the number formats in an XML node of the style document.
        /// </summary>
        /// <param name="node">Number formats root name.</param>
        private void GetNumberFormats(XmlNode node)
        {
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName.Equals("numfmt", StringComparison.OrdinalIgnoreCase))
                {
                    NumberFormat numberFormat = new NumberFormat();
                    int id = ParserUtils.ParseInt(ReaderUtils.GetAttribute(childNode, "numFmtId")); // Default will (justified) throw an exception
                    string code = ReaderUtils.GetAttribute(childNode, "formatCode", string.Empty);
                    numberFormat.CustomFormatID = id;
                    numberFormat.Number = FormatNumber.Custom;
                    numberFormat.InternalID = id;
                    numberFormat.CustomFormatCode = code;
                    this.styleReaderContainer.AddStyleComponent(numberFormat);
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
                    borderStyle.DiagonalColor = GetColor(innerNode, Border.DefaultBorderColor);
                }
                innerNode = ReaderUtils.GetChildNode(border, "top");
                if (innerNode != null)
                {
                    borderStyle.TopStyle = ParseBorderStyle(innerNode);
                    borderStyle.TopColor = GetColor(innerNode, Border.DefaultBorderColor);
                }
                innerNode = ReaderUtils.GetChildNode(border, "bottom");
                if (innerNode != null)
                {
                    borderStyle.BottomStyle = ParseBorderStyle(innerNode);
                    borderStyle.BottomColor = GetColor(innerNode, Border.DefaultBorderColor);
                }
                innerNode = ReaderUtils.GetChildNode(border, "left");
                if (innerNode != null)
                {
                    borderStyle.LeftStyle = ParseBorderStyle(innerNode);
                    borderStyle.LeftColor = GetColor(innerNode, Border.DefaultBorderColor);
                }
                innerNode = ReaderUtils.GetChildNode(border, "right");
                if (innerNode != null)
                {
                    borderStyle.RightStyle = ParseBorderStyle(innerNode);
                    borderStyle.RightColor = GetColor(innerNode, Border.DefaultBorderColor);
                }
                borderStyle.InternalID = this.styleReaderContainer.GetNextBorderId();
                this.styleReaderContainer.AddStyleComponent(borderStyle);
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
                return Border.GetStyleEnum(value);
            }
            return StyleValue.None;
        }

        /// <summary>
        /// Determines the fills in an XML node of the style document.
        /// </summary>
        /// <param name="node">Fill root node.</param>
        private void GetFills(XmlNode node)
        {
            foreach (XmlNode fill in node.ChildNodes)
            {
                Fill fillStyle = new Fill();
                XmlNode innerNode = ReaderUtils.GetChildNode(fill, "patternFill");
                if (innerNode != null)
                {
                    string pattern = ReaderUtils.GetAttribute(innerNode, "patternType", string.Empty);
                    fillStyle.PatternFill = Fill.GetPatternEnum(pattern);

                    // Read fgColor
                    XmlNode fgColorNode = ReaderUtils.GetChildNode(innerNode, "fgColor");
                    if (fgColorNode != null)
                    {
                        fillStyle.ForegroundColor = ReadColorFromNode(fgColorNode);
                    }

                    // Read bgColor
                    XmlNode bgColorNode = ReaderUtils.GetChildNode(innerNode, "bgColor");
                    if (bgColorNode != null)
                    {
                        fillStyle.BackgroundColor = ReadColorFromNode(bgColorNode);
                    }
                }
                fillStyle.InternalID = this.styleReaderContainer.GetNextFillId();
                this.styleReaderContainer.AddStyleComponent(fillStyle);
            }
        }

        /// <summary>
        /// Reads a CT_Color from an XML node (fgColor or bgColor element)
        /// </summary>
        /// <param name="colorNode">The color XML node</param>
        /// <returns>Color object representing the CT_Color</returns>
        private Color ReadColorFromNode(XmlNode colorNode)
        {
            // Check for auto attribute
            string autoAttr = ReaderUtils.GetAttribute(colorNode, "auto");
            if (!string.IsNullOrEmpty(autoAttr) && ParserUtils.ParseBinaryBool(autoAttr) == 1)
            {
                    return Color.CreateAuto();
            }

            // Check for rgb attribute
            string rgbAttr = ReaderUtils.GetAttribute(colorNode, "rgb");
            if (!string.IsNullOrEmpty(rgbAttr))
            {
                return Color.CreateRgb(rgbAttr);
            }

            // Check for indexed attribute
            string indexedAttr = ReaderUtils.GetAttribute(colorNode, "indexed");
            if (!string.IsNullOrEmpty(indexedAttr))
            {
                return Color.CreateIndexed(ParserUtils.ParseInt(indexedAttr));
            }

            // Check for theme attribute
            string themeAttr = ReaderUtils.GetAttribute(colorNode, "theme");
            if (!string.IsNullOrEmpty(themeAttr))
            {
                int themeIndex = ParserUtils.ParseInt(themeAttr);
                // Check for optional tint attribute
                string tintAttr = ReaderUtils.GetAttribute(colorNode, "tint");
                double? tint = null;
                if (!string.IsNullOrEmpty(tintAttr))
                {
                    tint = ParserUtils.ParseDouble(tintAttr); // Or Convert.ToDouble with InvariantCulture
                }
                return Color.CreateTheme((Theme.ColorSchemeElement)themeIndex, tint);
            }

            // Check for system attribute (if supported)
            string systemAttr = ReaderUtils.GetAttribute(colorNode, "system");
            if (!string.IsNullOrEmpty(systemAttr))
            {
                SystemColor sysColor = new SystemColor(SystemColor.MapStringToValue(systemAttr));
                return Color.CreateSystem(sysColor);
            }

            // No color defined
            return Color.CreateNone();
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
                    if (attribute == null)
                    {
                        fontStyle.Underline = Font.UnderlineValue.Single; // Default value
                    }
                    else
                    {
                        fontStyle.Underline = Font.GetUnderlineEnum(attribute);
                    }
                }
                if (ReaderUtils.GetAttributeOfChild(font, "vertAlign", "val", out attribute))
                {
                    fontStyle.VerticalAlign = Font.GetVerticalTextAlignEnum(attribute);
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
                                fontStyle.ColorTheme = ColorSchemeElement.Dark1;
                                break;
                            case "1":
                                fontStyle.ColorTheme = ColorSchemeElement.Light1;
                                break;
                            case "2":
                                fontStyle.ColorTheme = ColorSchemeElement.Dark2;
                                break;
                            case "3":
                                fontStyle.ColorTheme = ColorSchemeElement.Light2;
                                break;
                            case "4":
                                fontStyle.ColorTheme = ColorSchemeElement.Accent1;
                                break;
                            case "5":
                                fontStyle.ColorTheme = ColorSchemeElement.Accent2;
                                break;
                            case "6":
                                fontStyle.ColorTheme = ColorSchemeElement.Accent3;
                                break;
                            case "7":
                                fontStyle.ColorTheme = ColorSchemeElement.Accent4;
                                break;
                            case "8":
                                fontStyle.ColorTheme = ColorSchemeElement.Accent5;
                                break;
                            case "9":
                                fontStyle.ColorTheme = ColorSchemeElement.Accent6;
                                break;
                            case "10":
                                fontStyle.ColorTheme = ColorSchemeElement.Hyperlink;
                                break;
                            case "11":
                                fontStyle.ColorTheme = ColorSchemeElement.FollowedHyperlink;
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
                            fontStyle.Scheme = SchemeValue.Major;
                            break;
                        case "minor":
                            fontStyle.Scheme = SchemeValue.Minor;
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
                            fontStyle.Charset = CharsetValue.GBK;
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

                fontStyle.InternalID = this.styleReaderContainer.GetNextFontId();
                this.styleReaderContainer.AddStyleComponent(fontStyle);
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
                                cellXfStyle.Alignment = TextBreakValue.ShrinkToFit;
                            }
                        }
                        attribute = ReaderUtils.GetAttribute(alignmentNode, "wrapText");
                        if (attribute != null && attribute == "1")
                        {
                            cellXfStyle.Alignment = TextBreakValue.WrapText;
                        }
                        attribute = ReaderUtils.GetAttribute(alignmentNode, "horizontal", string.Empty);
                        cellXfStyle.HorizontalAlign = CellXf.GetHorizontalAlignEnum(attribute);
                        attribute = ReaderUtils.GetAttribute(alignmentNode, "vertical", string.Empty);
                        cellXfStyle.VerticalAlign = CellXf.GetVerticalAlignEnum(attribute);
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
                        if (attribute != null && attribute == "0")
                        {
                            cellXfStyle.Locked = false;
                        }
                        // else - NoOp - No need to set locked value, since true by default
                    }

                    cellXfStyle.InternalID = this.styleReaderContainer.GetNextCellXFId();
                    this.styleReaderContainer.AddStyleComponent(cellXfStyle);

                    Style style = new Style();
                    int id;
                    bool hasId;

                    hasId = ParserUtils.TryParseInt(ReaderUtils.GetAttribute(childNode, "numFmtId"), out id);
                    NumberFormat format = this.styleReaderContainer.GetNumberFormat(id);
                    if (!hasId || format == null)
                    {
                        FormatNumber formatNumber;
                        NumberFormat.TryParseFormatNumber(id, out formatNumber); // Validity is neglected here to prevent unhandled crashes. If invalid, the format will be declared as 'none'.
                                                                                 // Invalid values should not occur at all (malformed Excel files). 
                                                                                 // Undefined values may occur if the file was saved by an Excel version that has implemented yet unknown format numbers (undefined in NanoXLSX) 
                        format = new NumberFormat
                        {
                            Number = formatNumber,
                            InternalID = id
                        };
                        this.styleReaderContainer.AddStyleComponent(format);
                    }
                    hasId = ParserUtils.TryParseInt(ReaderUtils.GetAttribute(childNode, "borderId"), out id);
                    Border border = this.styleReaderContainer.GetBorder(id);
                    if (!hasId || border == null)
                    {
                        border = new Border
                        {
                            InternalID = this.styleReaderContainer.GetNextBorderId()
                        };
                    }
                    hasId = ParserUtils.TryParseInt(ReaderUtils.GetAttribute(childNode, "fillId"), out id);
                    Fill fill = this.styleReaderContainer.GetFill(id);
                    if (!hasId || fill == null)
                    {
                        fill = new Fill
                        {
                            InternalID = this.styleReaderContainer.GetNextFillId()
                        };
                    }
                    hasId = ParserUtils.TryParseInt(ReaderUtils.GetAttribute(childNode, "fontId"), out id);
                    Font font = this.styleReaderContainer.GetFont(id);
                    if (!hasId || font == null)
                    {
                        font = new Font
                        {
                            InternalID = this.styleReaderContainer.GetNextFontId()
                        };
                    }

                    // TODO: Implement other style information
                    style.CurrentNumberFormat = format;
                    style.CurrentBorder = border;
                    style.CurrentFill = fill;
                    style.CurrentFont = font;
                    style.CurrentCellXf = cellXfStyle;
                    style.InternalID = this.styleReaderContainer.GetNextStyleId();

                    this.styleReaderContainer.AddStyleComponent(style);
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
                if (color.Name.Equals("mruColors", StringComparison.Ordinal) && mruColor != null)
                {
                    foreach (XmlNode value in color.ChildNodes)
                    {
                        string attribute = ReaderUtils.GetAttribute(value, "rgb");
                        if (attribute != null)
                        {
                            this.styleReaderContainer.AddMruColor(attribute);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Resolves a color value from an XML node, when a RGB attribute exists. If the value is null, the fallback will be returned
        /// </summary>
        /// <param name="node">Node to check.</param>
        /// <param name="fallback">Fallback value if the color could not be resolved.</param>
        /// <returns>RGB value as string or the fallback.</returns>
        private static string GetColor(XmlNode node, string fallback)
        {
            XmlNode childNode = ReaderUtils.GetChildNode(node, "color");
            if (childNode != null)
            {
                string color = ReaderUtils.GetAttribute(childNode, "rgb");
                if (color != null)
                {
                    return color;
                }
            }
            return fallback;
        }
        #endregion
    }
}
