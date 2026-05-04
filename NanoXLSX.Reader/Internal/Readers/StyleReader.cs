/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

namespace NanoXLSX.Internal.Readers
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Xml;
    using NanoXLSX.Colors;
    using NanoXLSX.Interfaces;
    using NanoXLSX.Interfaces.Reader;
    using NanoXLSX.Registry;
    using NanoXLSX.Registry.Attributes;
    using NanoXLSX.Styles;
    using NanoXLSX.Themes;
    using NanoXLSX.Utils;
    using NanoXLSX.Utils.Xml;
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
    public class StyleReader : IPluginBaseReader
    {

        private Stream stream;
        private StyleReaderContainer styleReaderContainer;

        #region properties
        /// <summary>
        /// Workbook reference where read data is stored (should not be null)
        /// </summary>
        public Workbook Workbook { get; set; }
        /// <summary>
        /// Reader options
        /// </summary>
        public IOptions Options { get; set; }
        /// <summary>
        /// Reference to the <see cref="ReaderPlugInHandler"/>, to be used for post operations in the <see cref="Execute"/> method
        /// </summary>
        public Action<Stream, Workbook, string, IOptions, int?> InlinePluginHandler { get; set; }
        #endregion

        #region constructors
        /// <summary>
        /// Default constructor - Must be defined for instantiation of the plug-ins
        /// </summary>
        public StyleReader()
        {
        }
        #endregion

        #region methods

        /// <summary>
        /// Initialization method (interface implementation)
        /// </summary>
        /// <param name="stream">MemoryStream to be read</param>
        /// <param name="workbook">Workbook reference</param>
        /// <param name="readerOptions">Reader options</param>
        /// <param name="inlinePluginHandler">Reference to the a handler action, to be used for post operations in reader methods</param>
        public void Init(Stream stream, Workbook workbook, IOptions readerOptions, Action<Stream, Workbook, string, IOptions, int?> inlinePluginHandler)
        {
            this.stream = stream;
            this.Workbook = workbook;
            this.Options = readerOptions;
            this.InlinePluginHandler = inlinePluginHandler;
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
                    using (XmlReader reader = XmlReader.Create(stream, XmlStreamUtils.CreateSettings()))
                    {
                        while (reader.Read())
                        {
                            if (reader.NodeType != XmlNodeType.Element)
                            {
                                continue;
                            }
                            switch (reader.LocalName.ToLowerInvariant())
                            {
                                case "numfmts":
                                    GetNumberFormats(reader);
                                    break;
                                case "borders":
                                    GetBorders(reader);
                                    break;
                                case "fills":
                                    GetFills(reader);
                                    break;
                                case "fonts":
                                    GetFonts(reader);
                                    break;
                                case "colors":
                                    GetColors(reader);
                                    break;
                                case "cellxfs":
                                    GetCellXfs(reader);
                                    break;
                            }
                        }
                        HandleMruColors();
                        InlinePluginHandler?.Invoke(stream, Workbook, PlugInUUID.StyleInlineReader, Options, null);
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
            List<string> mruColors = styleReaderContainer.GetMruColors();
            foreach (string color in mruColors)
            {
                Workbook.AddMruColor(color);
            }
        }

        /// <summary>
        /// Determines the number formats in an XML node of the style document.
        /// </summary>
        /// <param name="reader">Reader positioned on the numFmts start element.</param>
        private void GetNumberFormats(XmlReader reader)
        {
            using (XmlReader subtree = reader.ReadSubtree())
            {
                subtree.Read(); // consume <numFmts>
                while (subtree.Read())
                {
                    if (!XmlStreamUtils.IsElement(subtree, "numFmt"))
                    {
                        continue;
                    }
                    NumberFormat numberFormat = new NumberFormat();
                    int id = ParserUtils.ParseInt(subtree.GetAttribute("numFmtId")); // null will rightly throw
                    string code = subtree.GetAttribute("formatCode") ?? string.Empty;
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
        /// <param name="reader">Reader positioned on the borders start element.</param>
        private void GetBorders(XmlReader reader)
        {
            using (XmlReader subtree = reader.ReadSubtree())
            {
                subtree.Read(); // consume <borders>
                while (subtree.Read())
                {
                    if (!XmlStreamUtils.IsElement(subtree, "border"))
                    {
                        continue;
                    }
                    Border borderStyle = new Border();
                    string diagonalDown = subtree.GetAttribute("diagonalDown");
                    if (diagonalDown != null && ParserUtils.ParseBinaryBool(diagonalDown) == 1)
                    {
                        borderStyle.DiagonalDown = true;
                    }
                    string diagonalUp = subtree.GetAttribute("diagonalUp");
                    if (diagonalUp != null && ParserUtils.ParseBinaryBool(diagonalUp) == 1)
                    {
                        borderStyle.DiagonalUp = true;
                    }
                    StyleValue sideStyle;
                    string sideColor;
                    using (XmlReader borderSubtree = subtree.ReadSubtree())
                    {
                        borderSubtree.Read(); // consume <border>
                        while (borderSubtree.Read())
                        {
                            if (borderSubtree.NodeType != XmlNodeType.Element)
                            {
                                continue;
                            }
                            switch (borderSubtree.LocalName.ToLowerInvariant())
                            {
                                case "diagonal":
                                    ReadBorderSide(borderSubtree, out sideStyle, out sideColor);
                                    borderStyle.DiagonalStyle = sideStyle;
                                    borderStyle.DiagonalColor = sideColor;
                                    break;
                                case "top":
                                    ReadBorderSide(borderSubtree, out sideStyle, out sideColor);
                                    borderStyle.TopStyle = sideStyle;
                                    borderStyle.TopColor = sideColor;
                                    break;
                                case "bottom":
                                    ReadBorderSide(borderSubtree, out sideStyle, out sideColor);
                                    borderStyle.BottomStyle = sideStyle;
                                    borderStyle.BottomColor = sideColor;
                                    break;
                                case "left":
                                    ReadBorderSide(borderSubtree, out sideStyle, out sideColor);
                                    borderStyle.LeftStyle = sideStyle;
                                    borderStyle.LeftColor = sideColor;
                                    break;
                                case "right":
                                    ReadBorderSide(borderSubtree, out sideStyle, out sideColor);
                                    borderStyle.RightStyle = sideStyle;
                                    borderStyle.RightColor = sideColor;
                                    break;
                            }
                        }
                    }
                    borderStyle.InternalID = this.styleReaderContainer.GetNextBorderId();
                    this.styleReaderContainer.AddStyleComponent(borderStyle);
                }
            }
        }

        /// <summary>
        /// Reads the style and color from a border side element (left, right, top, bottom, diagonal).
        /// Uses ReadSubtree so the outer reader advances past the element on return.
        /// </summary>
        private static void ReadBorderSide(XmlReader reader, out StyleValue style, out string color)
        {
            style = StyleValue.None;
            color = Border.DefaultBorderColor;
            string styleAttr = reader.GetAttribute("style");
            if (styleAttr != null)
            {
                style = Border.GetStyleEnum(styleAttr);
            }
            if (reader.IsEmptyElement)
            {
                return;
            }
            using (XmlReader subtree = reader.ReadSubtree())
            {
                subtree.Read(); // consume the side element open tag
                while (subtree.Read())
                {
                    if (subtree.NodeType == XmlNodeType.Element
                        && subtree.LocalName.Equals("color", StringComparison.OrdinalIgnoreCase))
                    {
                        string rgb = subtree.GetAttribute("rgb");
                        if (rgb != null)
                        {
                            color = rgb;
                        }
                        break;
                    }
                }
            }
        }

        /// <summary>
        /// Determines the fills in an XML node of the style document.
        /// </summary>
        /// <param name="reader">Reader positioned on the fills start element.</param>
        private void GetFills(XmlReader reader)
        {
            using (XmlReader subtree = reader.ReadSubtree())
            {
                subtree.Read(); // consume <fills>
                while (subtree.Read())
                {
                    if (!XmlStreamUtils.IsElement(subtree, "fill"))
                    {
                        continue;
                    }
                    Fill fillStyle = new Fill();
                    using (XmlReader fillSubtree = subtree.ReadSubtree())
                    {
                        fillSubtree.Read(); // consume <fill>
                        while (fillSubtree.Read())
                        {
                            if (!XmlStreamUtils.IsElement(fillSubtree, "patternFill"))
                            {
                                continue;
                            }
                            fillStyle.PatternFill = Fill.GetPatternEnum(fillSubtree.GetAttribute("patternType") ?? string.Empty);
                            using (XmlReader patternSubtree = fillSubtree.ReadSubtree())
                            {
                                patternSubtree.Read(); // consume <patternFill>
                                while (patternSubtree.Read())
                                {
                                    if (patternSubtree.NodeType != XmlNodeType.Element)
                                    {
                                        continue;
                                    }
                                    if (patternSubtree.LocalName.Equals("fgColor", StringComparison.OrdinalIgnoreCase))
                                    {
                                        fillStyle.ForegroundColor = ReadColorFromNode(patternSubtree);
                                    }
                                    else if (patternSubtree.LocalName.Equals("bgColor", StringComparison.OrdinalIgnoreCase))
                                    {
                                        fillStyle.BackgroundColor = ReadColorFromNode(patternSubtree);
                                    }
                                }
                            }
                        }
                    }
                    fillStyle.InternalID = this.styleReaderContainer.GetNextFillId();
                    this.styleReaderContainer.AddStyleComponent(fillStyle);
                }
            }
        }

        /// <summary>
        /// Reads a CT_Color from an XmlReader positioned on the color element (fgColor, bgColor, etc.)
        /// </summary>
        private static Color ReadColorFromNode(XmlReader reader)
        {
            string autoAttr = reader.GetAttribute("auto");
            if (!string.IsNullOrEmpty(autoAttr) && ParserUtils.ParseBinaryBool(autoAttr) == 1)
            {
                return Color.CreateAuto();
            }
            string rgbAttr = reader.GetAttribute("rgb");
            if (!string.IsNullOrEmpty(rgbAttr))
            {
                return Color.CreateRgb(rgbAttr);
            }
            string indexedAttr = reader.GetAttribute("indexed");
            if (!string.IsNullOrEmpty(indexedAttr))
            {
                return Color.CreateIndexed(ParserUtils.ParseInt(indexedAttr));
            }
            string themeAttr = reader.GetAttribute("theme");
            if (!string.IsNullOrEmpty(themeAttr))
            {
                int themeIndex = ParserUtils.ParseInt(themeAttr);
                string tintAttr = reader.GetAttribute("tint");
                double? tint = null;
                if (!string.IsNullOrEmpty(tintAttr))
                {
                    tint = ParserUtils.ParseDouble(tintAttr);
                }
                return Color.CreateTheme((Theme.ColorSchemeElement)themeIndex, tint);
            }
            string systemAttr = reader.GetAttribute("system");
            if (!string.IsNullOrEmpty(systemAttr))
            {
                return Color.CreateSystem(new SystemColor(SystemColor.MapStringToValue(systemAttr)));
            }
            return Color.CreateNone();
        }

        /// <summary>
        /// Determines the fonts in an XML node of the style document.
        /// </summary>
        /// <param name="reader">Reader positioned on the fonts start element.</param>
        private void GetFonts(XmlReader reader)
        {
            using (XmlReader subtree = reader.ReadSubtree())
            {
                subtree.Read(); // consume <fonts>
                while (subtree.Read())
                {
                    if (!XmlStreamUtils.IsElement(subtree, "font"))
                    {
                        continue;
                    }
                    Font fontStyle = new Font();
                    ReadFontElement(subtree, fontStyle);
                    fontStyle.InternalID = this.styleReaderContainer.GetNextFontId();
                    this.styleReaderContainer.AddStyleComponent(fontStyle);
                }
            }
        }

        /// <summary>
        /// Reads all child elements of a &lt;font&gt; entry into the given Font object.
        /// </summary>
        private static void ReadFontElement(XmlReader reader, Font fontStyle)
        {
            using (XmlReader fontSubtree = reader.ReadSubtree())
            {
                fontSubtree.Read(); // consume <font>
                while (fontSubtree.Read())
                {
                    if (fontSubtree.NodeType != XmlNodeType.Element)
                    {
                        continue;
                    }
                    string val;
                    switch (fontSubtree.LocalName.ToLowerInvariant())
                    {
                        case "b":
                            fontStyle.Bold = true;
                            break;
                        case "i":
                            fontStyle.Italic = true;
                            break;
                        case "strike":
                            fontStyle.Strike = true;
                            break;
                        case "outline":
                            fontStyle.Outline = true;
                            break;
                        case "shadow":
                            fontStyle.Shadow = true;
                            break;
                        case "condense":
                            fontStyle.Condense = true;
                            break;
                        case "extend":
                            fontStyle.Extend = true;
                            break;
                        case "u":
                            val = fontSubtree.GetAttribute("val");
                            fontStyle.Underline = val == null ? Font.UnderlineValue.Single : Font.GetUnderlineEnum(val);
                            break;
                        case "vertalign":
                            fontStyle.VerticalAlign = Font.GetVerticalTextAlignEnum(fontSubtree.GetAttribute("val"));
                            break;
                        case "sz":
                            fontStyle.Size = ParserUtils.ParseFloat(fontSubtree.GetAttribute("val"));
                            break;
                        case "color":
                            string themeVal = fontSubtree.GetAttribute("theme");
                            if (themeVal != null)
                            {
                                fontStyle.ColorValue = Color.CreateTheme(ParseFontColorSchemeElement(themeVal));
                            }
                            string rgbVal = fontSubtree.GetAttribute("rgb");
                            if (rgbVal != null)
                            {
                                fontStyle.ColorValue = Color.CreateRgb(rgbVal);
                            }
                            break;
                        case "name":
                            fontStyle.Name = fontSubtree.GetAttribute("val");
                            break;
                        case "family":
                            val = fontSubtree.GetAttribute("val");
                            if (val != null)
                            {
                                fontStyle.Family = ParseFontFamilyValue(val);
                            }
                            break;
                        case "scheme":
                            val = fontSubtree.GetAttribute("val");
                            if (val != null)
                            {
                                switch (val)
                                {
                                    case "major":
                                        fontStyle.Scheme = SchemeValue.Major;
                                        break;
                                    case "minor":
                                        fontStyle.Scheme = SchemeValue.Minor;
                                        break;
                                }
                            }
                            break;
                        case "charset":
                            val = fontSubtree.GetAttribute("val");
                            if (val != null)
                            {
                                fontStyle.Charset = ParseFontCharsetValue(val);
                            }
                            break;
                    }
                }
            }
        }

        /// <summary>
        /// Maps a theme index string ("0"-"11") to the corresponding ColorSchemeElement for font color.
        /// </summary>
        private static ColorSchemeElement ParseFontColorSchemeElement(string value)
        {
            switch (value)
            {
                case "1": return ColorSchemeElement.Light1;
                case "2": return ColorSchemeElement.Dark2;
                case "3": return ColorSchemeElement.Light2;
                case "4": return ColorSchemeElement.Accent1;
                case "5": return ColorSchemeElement.Accent2;
                case "6": return ColorSchemeElement.Accent3;
                case "7": return ColorSchemeElement.Accent4;
                case "8": return ColorSchemeElement.Accent5;
                case "9": return ColorSchemeElement.Accent6;
                case "10": return ColorSchemeElement.Hyperlink;
                case "11": return ColorSchemeElement.FollowedHyperlink;
                default: return ColorSchemeElement.Dark1;
            }
        }

        /// <summary>
        /// Maps a font family numeric string to FontFamilyValue.
        /// </summary>
        private static FontFamilyValue ParseFontFamilyValue(string value)
        {
            switch (value)
            {
                case "1": return FontFamilyValue.Roman;
                case "2": return FontFamilyValue.Swiss;
                case "3": return FontFamilyValue.Modern;
                case "4": return FontFamilyValue.Script;
                case "5": return FontFamilyValue.Decorative;
                case "6": return FontFamilyValue.Reserved1;
                case "7": return FontFamilyValue.Reserved2;
                case "8": return FontFamilyValue.Reserved3;
                case "9": return FontFamilyValue.Reserved4;
                case "10": return FontFamilyValue.Reserved5;
                case "11": return FontFamilyValue.Reserved6;
                case "12": return FontFamilyValue.Reserved7;
                case "13": return FontFamilyValue.Reserved8;
                case "14": return FontFamilyValue.Reserved9;
                default: return FontFamilyValue.NotApplicable;
            }
        }

        /// <summary>
        /// Maps a charset numeric string to CharsetValue.
        /// </summary>
        private static CharsetValue ParseFontCharsetValue(string value)
        {
            switch (value)
            {
                case "0": return CharsetValue.ANSI;
                case "1": return CharsetValue.Default;
                case "2": return CharsetValue.Symbols;
                case "77": return CharsetValue.Macintosh;
                case "128": return CharsetValue.JIS;
                case "129": return CharsetValue.Hangul;
                case "130": return CharsetValue.Johab;
                case "134": return CharsetValue.GBK;
                case "136": return CharsetValue.Big5;
                case "161": return CharsetValue.Greek;
                case "162": return CharsetValue.Turkish;
                case "163": return CharsetValue.Vietnamese;
                case "177": return CharsetValue.Hebrew;
                case "178": return CharsetValue.Arabic;
                case "186": return CharsetValue.Baltic;
                case "204": return CharsetValue.Russian;
                case "222": return CharsetValue.Thai;
                case "238": return CharsetValue.EasternEuropean;
                case "255": return CharsetValue.OEM;
                default: return CharsetValue.ApplicationDefined;
            }
        }

        /// <summary>
        /// Determines the cell XF entries in the style document.
        /// cellXfs always follows fonts/fills/borders per OOXML spec, so a single forward pass suffices.
        /// </summary>
        /// <param name="reader">Reader positioned on the cellXfs start element.</param>
        private void GetCellXfs(XmlReader reader)
        {
            using (XmlReader subtree = reader.ReadSubtree())
            {
                subtree.Read(); // consume <cellXfs>
                while (subtree.Read())
                {
                    if (!XmlStreamUtils.IsElement(subtree, "xf"))
                    {
                        continue;
                    }
                    CellXf cellXfStyle = new CellXf();
                    string applyAlignment = subtree.GetAttribute("applyAlignment");
                    if (applyAlignment != null)
                    {
                        cellXfStyle.ForceApplyAlignment = ParserUtils.ParseBinaryBool(applyAlignment) == 1;
                    }
                    string numFmtIdStr = subtree.GetAttribute("numFmtId");
                    string borderIdStr = subtree.GetAttribute("borderId");
                    string fillIdStr = subtree.GetAttribute("fillId");
                    string fontIdStr = subtree.GetAttribute("fontId");
                    if (!subtree.IsEmptyElement)
                    {
                        using (XmlReader xfSubtree = subtree.ReadSubtree())
                        {
                            xfSubtree.Read(); // consume <xf>
                            while (xfSubtree.Read())
                            {
                                if (xfSubtree.NodeType != XmlNodeType.Element)
                                {
                                    continue;
                                }
                                if (xfSubtree.LocalName.Equals("alignment", StringComparison.OrdinalIgnoreCase))
                                {
                                    ReadXfAlignment(xfSubtree, cellXfStyle);
                                }
                                else if (xfSubtree.LocalName.Equals("protection", StringComparison.OrdinalIgnoreCase))
                                {
                                    ReadXfProtection(xfSubtree, cellXfStyle);
                                }
                            }
                        }
                    }
                    cellXfStyle.InternalID = this.styleReaderContainer.GetNextCellXFId();
                    this.styleReaderContainer.AddStyleComponent(cellXfStyle);

                    Style style = new Style();
                    int id;
                    bool hasId;

                    hasId = ParserUtils.TryParseInt(numFmtIdStr, out id);
                    NumberFormat format = this.styleReaderContainer.GetNumberFormat(id);
                    if (!hasId || format == null)
                    {
                        FormatNumber formatNumber;
                        NumberFormat.TryParseFormatNumber(id, out formatNumber);
                        format = new NumberFormat
                        {
                            Number = formatNumber,
                            InternalID = id
                        };
                        this.styleReaderContainer.AddStyleComponent(format);
                    }
                    hasId = ParserUtils.TryParseInt(borderIdStr, out id);
                    Border border = this.styleReaderContainer.GetBorder(id);
                    if (!hasId || border == null)
                    {
                        border = new Border
                        {
                            InternalID = this.styleReaderContainer.GetNextBorderId()
                        };
                    }
                    hasId = ParserUtils.TryParseInt(fillIdStr, out id);
                    Fill fill = this.styleReaderContainer.GetFill(id);
                    if (!hasId || fill == null)
                    {
                        fill = new Fill
                        {
                            InternalID = this.styleReaderContainer.GetNextFillId()
                        };
                    }
                    hasId = ParserUtils.TryParseInt(fontIdStr, out id);
                    Font font = this.styleReaderContainer.GetFont(id);
                    if (!hasId || font == null)
                    {
                        font = new Font
                        {
                            InternalID = this.styleReaderContainer.GetNextFontId()
                        };
                    }
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
        /// Reads alignment attributes from an &lt;alignment&gt; element into the given CellXf.
        /// </summary>
        private static void ReadXfAlignment(XmlReader reader, CellXf cellXfStyle)
        {
            string shrinkToFit = reader.GetAttribute("shrinkToFit");
            if (shrinkToFit != null && ParserUtils.ParseBinaryBool(shrinkToFit) == 1)
            {
                cellXfStyle.Alignment = TextBreakValue.ShrinkToFit;
            }
            string wrapText = reader.GetAttribute("wrapText");
            if (wrapText != null && wrapText == "1")
            {
                cellXfStyle.Alignment = TextBreakValue.WrapText;
            }
            cellXfStyle.HorizontalAlign = CellXf.GetHorizontalAlignEnum(reader.GetAttribute("horizontal") ?? string.Empty);
            cellXfStyle.VerticalAlign = CellXf.GetVerticalAlignEnum(reader.GetAttribute("vertical") ?? string.Empty);
            string indent = reader.GetAttribute("indent");
            if (indent != null)
            {
                cellXfStyle.Indent = ParserUtils.ParseInt(indent);
            }
            string textRotation = reader.GetAttribute("textRotation");
            if (textRotation != null)
            {
                int rotation = ParserUtils.ParseInt(textRotation);
                cellXfStyle.TextRotation = rotation > 90 ? 90 - rotation : rotation;
            }
        }

        /// <summary>
        /// Reads protection attributes from a &lt;protection&gt; element into the given CellXf.
        /// </summary>
        private static void ReadXfProtection(XmlReader reader, CellXf cellXfStyle)
        {
            string hidden = reader.GetAttribute("hidden");
            if (hidden != null && hidden == "1")
            {
                cellXfStyle.Hidden = true;
            }
            string locked = reader.GetAttribute("locked");
            if (locked != null && locked == "0")
            {
                cellXfStyle.Locked = false;
            }
        }

        /// <summary>
        /// Determines the MRU colors in the style document.
        /// </summary>
        /// <param name="reader">Reader positioned on the colors start element.</param>
        private void GetColors(XmlReader reader)
        {
            using (XmlReader subtree = reader.ReadSubtree())
            {
                subtree.Read(); // consume <colors>
                while (subtree.Read())
                {
                    if (subtree.NodeType == XmlNodeType.Element
                        && subtree.LocalName.Equals("mruColors", StringComparison.OrdinalIgnoreCase))
                    {
                        ReadMruColors(subtree);
                    }
                }
            }
        }

        /// <summary>
        /// Reads individual color entries from a &lt;mruColors&gt; element.
        /// </summary>
        private void ReadMruColors(XmlReader reader)
        {
            using (XmlReader subtree = reader.ReadSubtree())
            {
                subtree.Read(); // consume <mruColors>
                while (subtree.Read())
                {
                    if (subtree.NodeType == XmlNodeType.Element
                        && subtree.LocalName.Equals("color", StringComparison.OrdinalIgnoreCase))
                    {
                        string rgb = subtree.GetAttribute("rgb");
                        if (rgb != null)
                        {
                            this.styleReaderContainer.AddMruColor(rgb);
                        }
                    }
                }
            }
        }
        #endregion
    }
}
