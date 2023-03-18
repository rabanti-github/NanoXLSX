/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2023
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Shared.Utils;
using NanoXLSX.Shared.Exceptions;
using NanoXLSX.Shared.Utils;
using System.Text.RegularExpressions;
using static NanoXLSX.Shared.Enums.Styles.BorderEnums;
using static NanoXLSX.Shared.Enums.Styles.FillEnums;
using static NanoXLSX.Shared.Enums.Styles.FontEnums;
using static NanoXLSX.Shared.Enums.Styles.NumberFormatEnums;

namespace NanoXLSX.Styles
{
    /// <summary>
    /// Factory class with the most important predefined styles
    /// </summary>
    public static class BasicStyles
    {
        #region enums
        /// <summary>
        /// Enum with style selection
        /// </summary>
        private enum StyleEnum
        {
            /// <summary>Format text bold</summary>
            bold,
            /// <summary>Format text italic</summary>
            italic,
            /// <summary>Format text bold and italic</summary>
            boldItalic,
            /// <summary>Format text with an underline</summary>
            underline,
            /// <summary>Format text with a double underline</summary>
            doubleUnderline,
            /// <summary>Format text with a strike-through</summary>
            strike,
            /// <summary>Format number as date</summary>
            dateFormat,
            /// <summary>Format number as time</summary>
            timeFormat,
            /// <summary>Rounds number as an integer</summary>
            roundFormat,
            /// <summary>Format cell with a thin border</summary>
            borderFrame,
            /// <summary>Format cell with a thin border and a thick bottom line as header cell</summary>
            borderFrameHeader,
            /// <summary>Special pattern fill style for compatibility purpose </summary>
            dottedFill_0_125,
            /// <summary>Style to apply on merged cells </summary>
            mergeCellStyle
        }
        #endregion

        #region staticFields
        private static Style bold;
        private static Style italic;
        private static Style boldItalic;
        private static Style underline;
        private static Style doubleUnderline;
        private static Style strike;
        private static Style dateFormat;
        private static Style timeFormat;
        private static Style roundFormat;
        private static Style borderFrame;
        private static Style borderFrameHeader;
        private static Style dottedFill_0_125;
        private static Style mergeCellStyle;
        #endregion

        #region staticProperties
        /// <summary>Gets the bold style</summary>
        public static Style Bold
        { get { return GetStyle(StyleEnum.bold); } }
        /// <summary>Gets the bold and italic style</summary>
        public static Style BoldItalic
        { get { return GetStyle(StyleEnum.boldItalic); } }
        /// <summary>Gets the border frame style</summary>
        public static Style BorderFrame
        { get { return GetStyle(StyleEnum.borderFrame); } }
        /// <summary>Gets the border style for header cells</summary>
        public static Style BorderFrameHeader
        { get { return GetStyle(StyleEnum.borderFrameHeader); } }
        /// <summary>Gets the date format style</summary>
        public static Style DateFormat
        { get { return GetStyle(StyleEnum.dateFormat); } }
        /// <summary>Gets the time format style</summary>
        public static Style TimeFormat
        { get { return GetStyle(StyleEnum.timeFormat); } }
        /// <summary>Gets the double underline style</summary>
        public static Style DoubleUnderline
        { get { return GetStyle(StyleEnum.doubleUnderline); } }
        /// <summary>Gets the special pattern fill style (for compatibility)</summary>
        public static Style DottedFill_0_125
        { get { return GetStyle(StyleEnum.dottedFill_0_125); } }
        /// <summary>Gets the italic style</summary>
        public static Style Italic
        { get { return GetStyle(StyleEnum.italic); } }
        /// <summary>Gets the style used when merging cells</summary>
        public static Style MergeCellStyle
        { get { return GetStyle(StyleEnum.mergeCellStyle); } }
        /// <summary>Gets the round format style</summary>
        public static Style RoundFormat
        { get { return GetStyle(StyleEnum.roundFormat); } }
        /// <summary>Gets the strike style</summary>
        public static Style Strike
        { get { return GetStyle(StyleEnum.strike); } }
        /// <summary>Gets the underline style</summary>
        public static Style Underline
        { get { return GetStyle(StyleEnum.underline); } }
        #endregion

        #region staticMethods
        /// <summary>
        /// Method to maintain the styles and to create singleton instances
        /// </summary>
        /// <param name="value">Enum value to maintain</param>
        /// <returns>The style according to the passed enum value</returns>
        private static Style GetStyle(StyleEnum value)
        {
            Style s = null;
            switch (value)
            {
                case StyleEnum.bold:
                    if (bold == null)
                    {
                        bold = new Style();
                        bold.CurrentFont.Bold = true;
                    }
                    s = bold;
                    break;
                case StyleEnum.italic:
                    if (italic == null)
                    {
                        italic = new Style();
                        italic.CurrentFont.Italic = true;
                    }
                    s = italic;
                    break;
                case StyleEnum.boldItalic:
                    if (boldItalic == null)
                    {
                        boldItalic = new Style();
                        boldItalic.CurrentFont.Italic = true;
                        boldItalic.CurrentFont.Bold = true;
                    }
                    s = boldItalic;
                    break;
                case StyleEnum.underline:
                    if (underline == null)
                    {
                        underline = new Style();
                        underline.CurrentFont.Underline = UnderlineValue.u_single;
                    }
                    s = underline;
                    break;
                case StyleEnum.doubleUnderline:
                    if (doubleUnderline == null)
                    {
                        doubleUnderline = new Style();
                        doubleUnderline.CurrentFont.Underline = UnderlineValue.u_double;
                    }
                    s = doubleUnderline;
                    break;
                case StyleEnum.strike:
                    if (strike == null)
                    {
                        strike = new Style();
                        strike.CurrentFont.Strike = true;
                    }
                    s = strike;
                    break;
                case StyleEnum.dateFormat:
                    if (dateFormat == null)
                    {
                        dateFormat = new Style();
                        dateFormat.CurrentNumberFormat.Number = FormatNumber.format_14;
                    }
                    s = dateFormat;
                    break;
                case StyleEnum.timeFormat:
                    if (timeFormat == null)
                    {
                        timeFormat = new Style();
                        timeFormat.CurrentNumberFormat.Number = FormatNumber.format_21;
                    }
                    s = timeFormat;
                    break;
                case StyleEnum.roundFormat:
                    if (roundFormat == null)
                    {
                        roundFormat = new Style();
                        roundFormat.CurrentNumberFormat.Number = FormatNumber.format_1;
                    }
                    s = roundFormat;
                    break;
                case StyleEnum.borderFrame:
                    if (borderFrame == null)
                    {
                        borderFrame = new Style();
                        borderFrame.CurrentBorder.TopStyle = StyleValue.thin;
                        borderFrame.CurrentBorder.BottomStyle = StyleValue.thin;
                        borderFrame.CurrentBorder.LeftStyle = StyleValue.thin;
                        borderFrame.CurrentBorder.RightStyle = StyleValue.thin;
                    }
                    s = borderFrame;
                    break;
                case StyleEnum.borderFrameHeader:
                    if (borderFrameHeader == null)
                    {
                        borderFrameHeader = new Style();
                        borderFrameHeader.CurrentBorder.TopStyle = StyleValue.thin;
                        borderFrameHeader.CurrentBorder.BottomStyle = StyleValue.medium;
                        borderFrameHeader.CurrentBorder.LeftStyle = StyleValue.thin;
                        borderFrameHeader.CurrentBorder.RightStyle = StyleValue.thin;
                        borderFrameHeader.CurrentFont.Bold = true;
                    }
                    s = borderFrameHeader;
                    break;
                case StyleEnum.dottedFill_0_125:
                    if (dottedFill_0_125 == null)
                    {
                        dottedFill_0_125 = new Style();
                        dottedFill_0_125.CurrentFill.PatternFill = PatternValue.gray125;
                    }
                    s = dottedFill_0_125;
                    break;
                case StyleEnum.mergeCellStyle:
                    if (mergeCellStyle == null)
                    {
                        mergeCellStyle = new Style();
                        mergeCellStyle.CurrentCellXf.ForceApplyAlignment = true;
                    }
                    s = mergeCellStyle;
                    break;
            }
            return s.CopyStyle(); // Copy makes basic styles immutable
        }

        /// <summary>
        /// Gets a style to colorize the text of a cell
        /// </summary>
        /// <param name="rgb">RGB code in hex format (6 characters, e.g. FF00AC). Alpha will be set to full opacity (FF)</param>
        /// <returns>Style with font color definition</returns>
        /// <exception cref="StyleException">A StyleException is thrown if an invalid hex value is passed</exception>
        public static Style ColorizedText(string rgb)
        {
            Validators.ValidateColor(rgb, false);
            Style s = new Style();
            s.CurrentFont.ColorValue = ParserUtils.ToUpper("FF" + rgb);
            return s;
        }

        /// <summary>
        /// Gets a style to colorize the background of a cell
        /// </summary>
        /// <param name="rgb">RGB code in hex format (6 characters, e.g. FF00AC). Alpha will be set to full opacity (FF)</param>
        /// <returns>Style with background color definition</returns>
        /// <exception cref="StyleException">A StyleException is thrown if an invalid hex value is passed</exception>
        public static Style ColorizedBackground(string rgb)
        {
            Validators.ValidateColor(rgb, false);
            Style s = new Style();
            s.CurrentFill.SetColor(ParserUtils.ToUpper("FF" + rgb), FillType.fillColor);

            return s;
        }

        /// <summary>
        /// Gets a style with a user defined font
        /// </summary>
        /// <param name="fontName">Name of the font</param>
        /// <param name="fontSize">Size of the font in points (optional; default 11)</param>
        /// <param name="isBold">If true, the font will be bold (optional; default false)</param>
        /// <param name="isItalic">If true, the font will be italic (optional; default false)</param>
        /// <returns>Style with font definition</returns>
        /// <remarks>The font name as well as the availability of bold and italic on the font cannot be validated by NanoXLSX. 
        /// The generated file may be corrupt or rendered with a fall-back font in case of an error</remarks>
        public static Style Font(string fontName, float fontSize = 11f, bool isBold = false, bool isItalic = false)
        {
            Style s = new Style();
            s.CurrentFont.Name = fontName;
            s.CurrentFont.Size = fontSize;
            s.CurrentFont.Bold = isBold;
            s.CurrentFont.Italic = isItalic;
            return s;
        }
        #endregion
    }
}
