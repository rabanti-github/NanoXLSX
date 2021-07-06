/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2021
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Globalization;
using System.Runtime.CompilerServices;
using NanoXLSX.Exceptions;


namespace NanoXLSX
{
    /// <summary>
    /// General Utils class with static methods
    /// </summary>
    public static class Utils
    {
        #region constants
        /// <summary>
        /// Minimum valid OAdate value (1900-01-01)
        /// </summary>
        public static readonly double MIN_OADATE_VALUE = 0d;
        /// <summary>
        /// Maximum valid OAdate value (9999-12-31)
        /// </summary>
        public static readonly double MAX_OADATE_VALUE = 2958465.9999d;
        /// <summary>
        /// All dates before this date are shifted in Excel by -1.0, since Excel assumes wrongly that the year 1900 is a leap year.<br/>
        /// See also: <a href="https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year">
        /// https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year</a>
        /// </summary>
        public static readonly DateTime FIRST_VALID_EXCEL_DATE = new DateTime(1900, 3, 1);

        /// <summary>
        /// Constant for number conversion. The invariant culture (represents mostly the US numbering scheme) ensures that no culture-specific 
        /// punctuations are used when converting numbers to strings, This is especially important for OOXML number values.
        /// See also: <a href="https://docs.microsoft.com/en-us/dotnet/api/system.globalization.cultureinfo.invariantculture?view=net-5.0">
        /// https://docs.microsoft.com/en-us/dotnet/api/system.globalization.cultureinfo.invariantculture?view=net-5.0</a>
        /// </summary>
        public static readonly CultureInfo INVARIANT_CULTURE = CultureInfo.InvariantCulture;

        private static readonly float COLUMN_WIDTH_ROUNDING_MODIFIER = 256f;
        private static readonly float SPLIT_WIDTH_MULTIPLIER = 12f;
        private static readonly float SPLIT_WIDTH_OFFSET = 0.5f;
        private static readonly float SPLIT_WIDTH_POINT_MULTIPLIER = 3f / 4f;
        private static readonly float SPLIT_POINT_DIVIDER = 20f;
        private static readonly float SPLIT_WIDTH_POINT_OFFSET = 390f;
        private static readonly float SPLIT_HEIGHT_POINT_OFFSET = 300f;

        #endregion

        /// <summary>
        /// Method to convert a date or date and time into the internal Excel time format (OAdate)
        /// </summary>
        /// <param name="date">Date to process</param>
        /// <returns>Date or date and time as number</returns>
        /// <exception cref="Exceptions.FormatException">Throws a FormatException if the passed date cannot be translated to the OADate format</exception>
        /// <remarks>Excel assumes wrongly that the year 1900 is a leap year. There is a gap of 1.0 between 1900-02-28 and 1900-03-01. This method corrects all dates
        /// from the first valid date (1900-01-01) to 1900-03-01. However, Excel displays the minimum valid date as 1900-01-00, although 0 is not a valid description for a day of month.
        /// In conformance to the OAdate specifications, the maximum valid date is 9999-12-31 23:59:59 (plus 999 milliseconds).<br/>
        ///See also: <a href="https://docs.microsoft.com/en-us/dotnet/api/system.datetime.tooadate?view=netcore-3.1">
        ///https://docs.microsoft.com/en-us/dotnet/api/system.datetime.tooadate?view=netcore-3.1</a><br/>
        ///See also: <a href="https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year">
        ///https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year</a>
        /// </remarks>
        public static string GetOADateTimeString(DateTime date, IFormatProvider culture)
        {
            try
            {
                DateTime dateValue = date;
                if (date < FIRST_VALID_EXCEL_DATE)
                {
                    dateValue = date.AddDays(-1); // Fix of the leap-year-1900-error
                }
                double d = dateValue.ToOADate();
                if (d < MIN_OADATE_VALUE || d > MAX_OADATE_VALUE)
                {
                    throw new Exceptions.FormatException("The date is not in a valid range for Excel. Dates before 1900-01-01 or after 9999-12-31 are not allowed.");
                }
                return d.ToString("G", culture);
            }
            catch (Exception e)
            {
                throw new Exceptions.FormatException("ConversionException", "The date could not be transformed into Excel format (OADate).", e);
            }
        }

        /// <summary>
        /// Method to convert a time into the internal Excel time format (OAdate without days)
        /// </summary>
        /// <param name="time">Time to process. The date component of the timespan is neglected</param>
        /// <param name="culture">CultureInfo for proper formatting of the decimal point</param>
        /// <returns>Time as number</returns>
        /// <exception cref="Exceptions.FormatException">Throws a FormatException if the passed timespan is invalid</exception>
        /// <remarks>The time is represented by a OAdate without the date component. A time range is between &gt;0.0 (00:00:00) and &lt;1.0 (23:59:59)</remarks>
        public static string GetOATimeString(TimeSpan time, IFormatProvider culture)
        {
            try
            {
                int seconds = time.Seconds + time.Minutes * 60 + time.Hours * 3600;
                double d = (double)seconds / 86400d;
                return d.ToString("G", culture);
            }
            catch (Exception e)
            {
                throw new Exceptions.FormatException("ConversionException", "The time could not be transformed into Excel format (OADate).", e);
            }
        }

        /// <summary>
        /// Transforms a string to upper case with null check and invariant culture
        /// </summary>
        /// <param name="input">String to transform</param>
        /// <returns>Upper case string</returns>
        public static string ToUpper(string input)
        {
            if (!string.IsNullOrEmpty(input))
            {
                return input.ToUpper(INVARIANT_CULTURE);
            }
            else
            {
                return input;
            }
            
        }

        /// <summary>
        /// Transforms an integer to an invariant sting
        /// </summary>
        /// <param name="input">Integer to transform</param>
        /// <returns>Integer as string</returns>
        public static string ToString(int input)
        {
            return input.ToString("G", INVARIANT_CULTURE);
        }

        /// <summary>
        /// Calculates the internal width of a column in characters. This width is used only in the XML documents of worksheets and is usually not exposed to the (Excel) end user
        /// </summary>
        /// <remarks>
        /// The internal width deviates slightly from the column width, entered in Excel. Although internal, the default column width of 10 characters is visible in Excel as 10.71.
        /// The deviation depends on the maximum digit width of the default font, as well as its text padding and various constants.<br/>
        /// In case of the width 10.0 and the default digit width 7.0, as well as the padding 5.0 of the default font Calibri (size 11), 
        /// the internal width is approximately 10.7142857 (rounded to 10.71).<br/> Note that the column hight is not affected by this consideration. 
        /// The entered height in Excel is the actual height in the worksheet XML documents.<br/> 
        /// This method is derived from the Perl implementation by John McNamara (<a href="https://stackoverflow.com/a/5010899">https://stackoverflow.com/a/5010899</a>)<br/>
        /// See also: <a href="https://www.ecma-international.org/publications-and-standards/standards/ecma-376/">ECMA-376, Part 1, Chapter 18.3.1.13</a>
        /// </remarks>
        /// <param name="columnWidth">Target column width (displayed in Excel)</param>
        /// <param name="maxDigitWidth">Maximum digit with of the default font (default is 7.0 for Calibri, size 11)</param>
        /// <param name="textPadding">Text padding of the default font (default is 5.0 for Calibri, size 11)</param>
        /// <returns>The internal column width in characters, used in worksheet XML documents</returns>
        public static float GetInternalColumnWidth(float columnWidth, float maxDigitWidth = 7f, float textPadding = 5f)
        {
            if (columnWidth <= 0f || maxDigitWidth <= 0f)
            {
                return 0f;
            }
            else if (columnWidth <= 1f)
            {
                return (float)Math.Floor((columnWidth * (maxDigitWidth + textPadding)) / maxDigitWidth * COLUMN_WIDTH_ROUNDING_MODIFIER) / COLUMN_WIDTH_ROUNDING_MODIFIER;
            }
            else
            {
                return (float)Math.Floor((columnWidth * maxDigitWidth + textPadding) / maxDigitWidth * COLUMN_WIDTH_ROUNDING_MODIFIER) / COLUMN_WIDTH_ROUNDING_MODIFIER;
            }
        }

        /// <summary>
        /// Calculates the internal width of a split pane in a worksheet. This width is used only in the XML documents of worksheets and is not exposed to the (Excel) end user
        /// </summary>
        /// <remarks>
        /// The internal split width is based on the width of one or more columns. 
        /// It also depends on the maximum digit width of the default font, as well as its text padding and various constants.<br/>
        /// See also <see cref="GetInternalColumnWidth(float, float, float)"/> for additional details.<br/>
        /// This method is derived from the Perl implementation by John McNamara (<a href="https://stackoverflow.com/a/5010899">https://stackoverflow.com/a/5010899</a>)<br/>
        /// See also: <a href="https://www.ecma-international.org/publications-and-standards/standards/ecma-376/">ECMA-376, Part 1, Chapter 18.3.1.13</a><br/>
        /// The three optional parameters maxDigitWidth and textPadding probably don't have to be changed ever.
        /// </remarks>
        /// <param name="width">Target column(s) width (one or more columns, displayed in Excel)</param>
        /// <param name="maxDigitWidth">Maximum digit with of the default font (default is 7.0 for Calibri, size 11)</param>
        /// <param name="textPadding">Text padding of the default font (default is 5.0 for Calibri, size 11)</param>
        /// <returns>The internal pane width, used in worksheet XML documents in case of worksheet splitting</returns>
        public static float GetInternalPaneSplitWidth(float width, float maxDigitWidth = 7f, float textPadding = 5f)
        {
            float pixels;
            if (width <= 1f)
            {
                pixels = (float)Math.Floor(width / SPLIT_WIDTH_MULTIPLIER + SPLIT_WIDTH_OFFSET);
            }
            else
            {
                pixels = (float)Math.Floor(width * maxDigitWidth + SPLIT_WIDTH_OFFSET) + textPadding;
            }
            float points = pixels * SPLIT_WIDTH_POINT_MULTIPLIER;
            return points * SPLIT_POINT_DIVIDER + SPLIT_WIDTH_POINT_OFFSET;
        }

        /// <summary>
        /// Calculates the internal height of a split pane in a worksheet. This height is used only in the XML documents of worksheets and is not exposed to the (Excel) user
        /// </summary>
        /// <remarks>
        /// The internal split height is based on the height of one or more rows. It also depends on various constants.<br/>
        /// This method is derived from the Perl implementation by John McNamara (<a href="https://stackoverflow.com/a/5010899">https://stackoverflow.com/a/5010899</a>)
        /// </remarks>
        /// <param name="height">Target row(s) height (one or more rows, displayed in Excel)</param>
        /// <returns>The internal pane height, used in worksheet XML documents in case of worksheet splitting</returns>
        public static float GetInternalPaneSplitHeight(float height)
        {
            return (float)Math.Floor(SPLIT_POINT_DIVIDER * height + SPLIT_HEIGHT_POINT_OFFSET);
        }
    }
}
