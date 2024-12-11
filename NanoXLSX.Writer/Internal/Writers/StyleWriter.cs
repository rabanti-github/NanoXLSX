/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2024
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Collections.Generic;
using System.Text;
using NanoXLSX.Shared.Exceptions;
using NanoXLSX.Shared.Utils;
using NanoXLSX.Styles;

namespace NanoXLSX.Internal.Writers
{
    internal class StyleWriter
    {

        private readonly StyleManager styles;
        private readonly Workbook workbook;

        internal StyleWriter(XlsxWriter writer)
        {
            this.styles = writer.Styles;
            this.workbook = writer.Workbook;
        }

        /// <summary>
        /// Method to create a style sheet as raw XML string
        /// </summary>
        /// <returns>Raw XML string</returns>
        /// <exception cref="StyleException">Throws a StyleException if one of the styles cannot be referenced or is null</exception>
        /// <remarks>The UndefinedStyleException should never happen in this state if the internally managed style collection was not tampered. </remarks>
        internal string CreateStyleSheetDocument()
        {
            string bordersString = CreateStyleBorderString();
            string fillsString = CreateStyleFillString();
            string fontsString = CreateStyleFontString();
            string numberFormatsString = CreateStyleNumberFormatString();
            string xfsStings = CreateStyleXfsString();
            string mruColorString = CreateMruColorsString();
            int fontCount = styles.GetFontStyleNumber();
            int fillCount = styles.GetFillStyleNumber();
            int styleCount = styles.GetStyleNumber();
            int borderCount = styles.GetBorderStyleNumber();
            StringBuilder sb = new StringBuilder();
            sb.Append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">");
            int numFormatCount = styles.GetNumberFormatStyleNumber();
            if (numFormatCount > 0)
            {
                sb.Append("<numFmts count=\"").Append(ParserUtils.ToString(numFormatCount)).Append("\">");
                sb.Append(numberFormatsString + "</numFmts>");
            }
            sb.Append("<fonts x14ac:knownFonts=\"1\" count=\"").Append(ParserUtils.ToString(fontCount)).Append("\">");
            sb.Append(fontsString).Append("</fonts>");
            sb.Append("<fills count=\"").Append(ParserUtils.ToString(fillCount)).Append("\">");
            sb.Append(fillsString).Append("</fills>");
            sb.Append("<borders count=\"").Append(ParserUtils.ToString(borderCount)).Append("\">");
            sb.Append(bordersString).Append("</borders>");
            sb.Append("<cellXfs count=\"").Append(ParserUtils.ToString(styleCount)).Append("\">");
            sb.Append(xfsStings).Append("</cellXfs>");
            if (workbook.WorkbookMetadata != null)
            {
                if (!string.IsNullOrEmpty(mruColorString))
                {
                    sb.Append("<colors>");
                    sb.Append(mruColorString);
                    sb.Append("</colors>");
                }
            }
            sb.Append("</styleSheet>");
            return sb.ToString();
        }

        /// <summary>
        /// Method to create the XML string for the border part of the style sheet document
        /// </summary>
        /// <returns>String with formatted XML data</returns>
        private string CreateStyleBorderString()
        {
            Border[] borderStyles = styles.GetBorders();
            StringBuilder sb = new StringBuilder();
            foreach (Border item in borderStyles)
            {
                if (item.DiagonalDown && !item.DiagonalUp) { sb.Append("<border diagonalDown=\"1\">"); }
                else if (!item.DiagonalDown && item.DiagonalUp) { sb.Append("<border diagonalUp=\"1\">"); }
                else if (item.DiagonalDown && item.DiagonalUp) { sb.Append("<border diagonalDown=\"1\" diagonalUp=\"1\">"); }
                else { sb.Append("<border>"); }

                if (item.LeftStyle != StyleValue.none)
                {
                    sb.Append("<left style=\"" + Border.GetStyleName(item.LeftStyle) + "\">");
                    if (!string.IsNullOrEmpty(item.LeftColor)) { sb.Append("<color rgb=\"").Append(item.LeftColor).Append("\"/>"); }
                    else { sb.Append("<color auto=\"1\"/>"); }
                    sb.Append("</left>");
                }
                else
                {
                    sb.Append("<left/>");
                }
                if (item.RightStyle != StyleValue.none)
                {
                    sb.Append("<right style=\"").Append(Border.GetStyleName(item.RightStyle)).Append("\">");
                    if (!string.IsNullOrEmpty(item.RightColor)) { sb.Append("<color rgb=\"").Append(item.RightColor).Append("\"/>"); }
                    else { sb.Append("<color auto=\"1\"/>"); }
                    sb.Append("</right>");
                }
                else
                {
                    sb.Append("<right/>");
                }
                if (item.TopStyle != StyleValue.none)
                {
                    sb.Append("<top style=\"").Append(Border.GetStyleName(item.TopStyle)).Append("\">");
                    if (!string.IsNullOrEmpty(item.TopColor)) { sb.Append("<color rgb=\"").Append(item.TopColor).Append("\"/>"); }
                    else { sb.Append("<color auto=\"1\"/>"); }
                    sb.Append("</top>");
                }
                else
                {
                    sb.Append("<top/>");
                }
                if (item.BottomStyle != StyleValue.none)
                {
                    sb.Append("<bottom style=\"").Append(Border.GetStyleName(item.BottomStyle)).Append("\">");
                    if (!string.IsNullOrEmpty(item.BottomColor)) { sb.Append("<color rgb=\"").Append(item.BottomColor).Append("\"/>"); }
                    else { sb.Append("<color auto=\"1\"/>"); }
                    sb.Append("</bottom>");
                }
                else
                {
                    sb.Append("<bottom/>");
                }
                if (item.DiagonalStyle != StyleValue.none)
                {
                    sb.Append("<diagonal style=\"").Append(Border.GetStyleName(item.DiagonalStyle)).Append("\">");
                    if (!string.IsNullOrEmpty(item.DiagonalColor)) { sb.Append("<color rgb=\"").Append(item.DiagonalColor).Append("\"/>"); }
                    else { sb.Append("<color auto=\"1\"/>"); }
                    sb.Append("</diagonal>");
                }
                else
                {
                    sb.Append("<diagonal/>");
                }
                sb.Append("</border>");
            }
            return sb.ToString();
        }

        /// <summary>
        /// Method to create the XML string for the font part of the style sheet document
        /// </summary>
        /// <returns>String with formatted XML data</returns>
        private string CreateStyleFontString()
        {
            Font[] fontStyles = styles.GetFonts();
            StringBuilder sb = new StringBuilder();
            foreach (Font item in fontStyles)
            {
                sb.Append("<font>");
                if (item.Bold) { sb.Append("<b/>"); }
                if (item.Italic) { sb.Append("<i/>"); }
                if (item.Strike) { sb.Append("<strike/>"); }
                if (item.Underline != UnderlineValue.none)
                {
                    if (item.Underline == UnderlineValue.u_double) { sb.Append("<u val=\"double\"/>"); }
                    else if (item.Underline == UnderlineValue.singleAccounting) { sb.Append("<u val=\"singleAccounting\"/>"); }
                    else if (item.Underline == UnderlineValue.doubleAccounting) { sb.Append("<u val=\"doubleAccounting\"/>"); }
                    else { sb.Append("<u/>"); }
                }
                if (item.VerticalAlign == VerticalTextAlignValue.subscript) { sb.Append("<vertAlign val=\"subscript\"/>"); }
                else if (item.VerticalAlign == VerticalTextAlignValue.superscript) { sb.Append("<vertAlign val=\"superscript\"/>"); }
                sb.Append("<sz val=\"").Append(ParserUtils.ToString(item.Size)).Append("\"/>");
                if (string.IsNullOrEmpty(item.ColorValue))
                {
                    sb.Append("<color theme=\"").Append(ParserUtils.ToString((int)item.ColorTheme)).Append("\"/>");
                }
                else
                {
                    sb.Append("<color rgb=\"").Append(item.ColorValue).Append("\"/>");
                }
                sb.Append("<name val=\"").Append(item.Name).Append("\"/>");
                sb.Append("<family val=\"").Append(ParserUtils.ToString((int)item.Family)).Append("\"/>");
                if (item.Scheme != SchemeValue.none)
                {
                    if (item.Scheme == SchemeValue.major)
                    { sb.Append("<scheme val=\"major\"/>"); }
                    else if (item.Scheme == SchemeValue.minor)
                    { sb.Append("<scheme val=\"minor\"/>"); }
                }
                if (item.Charset != CharsetValue.Default)
                {
                    sb.Append("<charset val=\"").Append(ParserUtils.ToString((int)item.Charset)).Append("\"/>");
                }
                sb.Append("</font>");
            }
            return sb.ToString();
        }

        /// <summary>
        /// Method to create the XML string for the fill part of the style sheet document
        /// </summary>
        /// <returns>String with formatted XML data</returns>
        private string CreateStyleFillString()
        {
            Fill[] fillStyles = styles.GetFills();
            StringBuilder sb = new StringBuilder();
            foreach (Fill item in fillStyles)
            {
                sb.Append("<fill>");
                sb.Append("<patternFill patternType=\"").Append(Fill.GetPatternName(item.PatternFill)).Append("\"");
                if (item.PatternFill == PatternValue.solid)
                {
                    sb.Append(">");
                    sb.Append("<fgColor rgb=\"").Append(item.ForegroundColor).Append("\"/>");
                    sb.Append("<bgColor indexed=\"").Append(ParserUtils.ToString(item.IndexedColor)).Append("\"/>");
                    sb.Append("</patternFill>");
                }
                else if (item.PatternFill == PatternValue.mediumGray || item.PatternFill == PatternValue.lightGray || item.PatternFill == PatternValue.gray0625 || item.PatternFill == PatternValue.darkGray)
                {
                    sb.Append(">");
                    sb.Append("<fgColor rgb=\"").Append(item.ForegroundColor).Append("\"/>");
                    if (!string.IsNullOrEmpty(item.BackgroundColor))
                    {
                        sb.Append("<bgColor rgb=\"").Append(item.BackgroundColor).Append("\"/>");
                    }
                    sb.Append("</patternFill>");
                }
                else
                {
                    sb.Append("/>");
                }
                sb.Append("</fill>");
            }
            return sb.ToString();
        }

        /// <summary>
        /// Method to create the XML string for the number format part of the style sheet document 
        /// </summary>
        /// <returns>String with formatted XML data</returns>
        private string CreateStyleNumberFormatString()
        {
            NumberFormat[] numberFormatStyles = styles.GetNumberFormats();
            StringBuilder sb = new StringBuilder();
            foreach (NumberFormat item in numberFormatStyles)
            {
                if (item.IsCustomFormat)
                {
                    if (string.IsNullOrEmpty(item.CustomFormatCode))
                    {
                        throw new Shared.Exceptions.FormatException("The number format style component with the ID " + ParserUtils.ToString(item.CustomFormatID) + " cannot be null or empty");
                    }
                    // OOXML: Escaping according to Chp.18.8.31
                    // TODO: v3> Add a custom format builder
                    sb.Append("<numFmt formatCode=\"").Append(XmlUtils.EscapeXmlAttributeChars(item.CustomFormatCode)).Append("\" numFmtId=\"").Append(ParserUtils.ToString(item.CustomFormatID)).Append("\"/>");
                }
            }
            return sb.ToString();
        }

        /// <summary>
        /// Method to create the XML string for the XF part of the style sheet document
        /// </summary>
        /// <returns>String with formatted XML data</returns>
        private string CreateStyleXfsString()
        {
            Style[] styleItems = this.styles.GetStyles();
            StringBuilder sb = new StringBuilder();
            StringBuilder sb2 = new StringBuilder();
            string alignmentString, protectionString;
            int formatNumber, textRotation;
            foreach (Style style in styleItems)
            {
                textRotation = style.CurrentCellXf.CalculateInternalRotation();
                alignmentString = string.Empty;
                protectionString = string.Empty;
                if (style.CurrentCellXf.HorizontalAlign != HorizontalAlignValue.none || style.CurrentCellXf.VerticalAlign != VerticalAlignValue.none || style.CurrentCellXf.Alignment != TextBreakValue.none || textRotation != 0)
                {
                    sb2.Clear();
                    sb2.Append("<alignment");
                    if (style.CurrentCellXf.HorizontalAlign != HorizontalAlignValue.none)
                    {
                        sb2.Append(" horizontal=\"");
                        if (style.CurrentCellXf.HorizontalAlign == HorizontalAlignValue.center) { sb2.Append("center"); }
                        else if (style.CurrentCellXf.HorizontalAlign == HorizontalAlignValue.right) { sb2.Append("right"); }
                        else if (style.CurrentCellXf.HorizontalAlign == HorizontalAlignValue.centerContinuous) { sb2.Append("centerContinuous"); }
                        else if (style.CurrentCellXf.HorizontalAlign == HorizontalAlignValue.distributed) { sb2.Append("distributed"); }
                        else if (style.CurrentCellXf.HorizontalAlign == HorizontalAlignValue.fill) { sb2.Append("fill"); }
                        else if (style.CurrentCellXf.HorizontalAlign == HorizontalAlignValue.general) { sb2.Append("general"); }
                        else if (style.CurrentCellXf.HorizontalAlign == HorizontalAlignValue.justify) { sb2.Append("justify"); }
                        else { sb2.Append("left"); }
                        sb2.Append("\"");
                    }
                    if (style.CurrentCellXf.VerticalAlign != VerticalAlignValue.none)
                    {
                        sb2.Append(" vertical=\"");
                        if (style.CurrentCellXf.VerticalAlign == VerticalAlignValue.center) { sb2.Append("center"); }
                        else if (style.CurrentCellXf.VerticalAlign == VerticalAlignValue.distributed) { sb2.Append("distributed"); }
                        else if (style.CurrentCellXf.VerticalAlign == VerticalAlignValue.justify) { sb2.Append("justify"); }
                        else if (style.CurrentCellXf.VerticalAlign == VerticalAlignValue.top) { sb2.Append("top"); }
                        else { sb2.Append("bottom"); }
                        sb2.Append("\"");
                    }
                    if (style.CurrentCellXf.Indent > 0 &&
                        (style.CurrentCellXf.HorizontalAlign == HorizontalAlignValue.left
                        || style.CurrentCellXf.HorizontalAlign == HorizontalAlignValue.right
                        || style.CurrentCellXf.HorizontalAlign == HorizontalAlignValue.distributed))
                    {
                        sb2.Append(" indent=\"");
                        sb2.Append(ParserUtils.ToString(style.CurrentCellXf.Indent));
                        sb2.Append("\"");
                    }
                    if (style.CurrentCellXf.Alignment != TextBreakValue.none)
                    {
                        if (style.CurrentCellXf.Alignment == TextBreakValue.shrinkToFit) { sb2.Append(" shrinkToFit=\"1"); }
                        else { sb2.Append(" wrapText=\"1"); }
                        sb2.Append("\"");
                    }
                    if (textRotation != 0)
                    {
                        sb2.Append(" textRotation=\"");
                        sb2.Append(ParserUtils.ToString(textRotation));
                        sb2.Append("\"");
                    }
                    sb2.Append("/>"); // </xf>
                    alignmentString = sb2.ToString();
                }

                if (style.CurrentCellXf.Hidden || style.CurrentCellXf.Locked)
                {
                    if (style.CurrentCellXf.Hidden && style.CurrentCellXf.Locked)
                    {
                        protectionString = "<protection locked=\"1\" hidden=\"1\"/>";
                    }
                    else if (style.CurrentCellXf.Hidden && !style.CurrentCellXf.Locked)
                    {
                        protectionString = "<protection hidden=\"1\" locked=\"0\"/>";
                    }
                    else
                    {
                        protectionString = "<protection hidden=\"0\" locked=\"1\"/>";
                    }
                }

                sb.Append("<xf numFmtId=\"");
                if (style.CurrentNumberFormat.IsCustomFormat)
                {
                    sb.Append(ParserUtils.ToString(style.CurrentNumberFormat.CustomFormatID));
                }
                else
                {
                    formatNumber = (int)style.CurrentNumberFormat.Number;
                    sb.Append(ParserUtils.ToString(formatNumber));
                }

                sb.Append("\" borderId=\"").Append(ParserUtils.ToString(style.CurrentBorder.InternalID.Value));
                sb.Append("\" fillId=\"").Append(ParserUtils.ToString(style.CurrentFill.InternalID.Value));
                sb.Append("\" fontId=\"").Append(ParserUtils.ToString(style.CurrentFont.InternalID.Value));
                if (!style.CurrentFont.IsDefaultFont)
                {
                    sb.Append("\" applyFont=\"1");
                }
                if (style.CurrentFill.PatternFill != PatternValue.none)
                {
                    sb.Append("\" applyFill=\"1");
                }
                if (!style.CurrentBorder.IsEmpty())
                {
                    sb.Append("\" applyBorder=\"1");
                }
                if (alignmentString != string.Empty || style.CurrentCellXf.ForceApplyAlignment)
                {
                    sb.Append("\" applyAlignment=\"1");
                }
                if (protectionString != string.Empty)
                {
                    sb.Append("\" applyProtection=\"1");
                }
                if (style.CurrentNumberFormat.Number != FormatNumber.none)
                {
                    sb.Append("\" applyNumberFormat=\"1\"");
                }
                else
                {
                    sb.Append("\"");
                }
                if (alignmentString != string.Empty || protectionString != string.Empty)
                {
                    sb.Append(">");
                    sb.Append(alignmentString);
                    sb.Append(protectionString);
                    sb.Append("</xf>");
                }
                else
                {
                    sb.Append("/>");
                }
            }
            return sb.ToString();
        }

        /// <summary>
        /// Method to create the XML string for the color-MRU part of the style sheet document (recent colors)
        /// </summary>
        /// <returns>String with formatted XML data</returns>
        private string CreateMruColorsString()
        {
            StringBuilder sb = new StringBuilder();
            List<string> tempColors = new List<string>();
            foreach (string item in this.workbook.GetMruColors())
            {
                if (item == Fill.DEFAULT_COLOR)
                {
                    continue;
                }
                if (!tempColors.Contains(item)) { tempColors.Add(item); }
            }
            if (tempColors.Count > 0)
            {
                sb.Append("<mruColors>");
                foreach (string item in tempColors)
                {
                    sb.Append("<color rgb=\"").Append(item).Append("\"/>");
                }
                sb.Append("</mruColors>");
                return sb.ToString();
            }
            return string.Empty;
        }
    }
}
