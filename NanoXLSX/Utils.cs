/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2022
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Globalization;
using FormatException = NanoXLSX.Shared.Exceptions.FormatException;

namespace NanoXLSX
{
    /// <summary>
    /// General Utils class with static methods
    /// </summary>
    public static class Utils
    {
        #region constants
        /// <summary>
        /// Minimum valid OAdate value (1900-01-01). However, Excel displays this value as 1900-01-00 (day zero)
        /// </summary>
        public static readonly double MIN_OADATE_VALUE = 0d;
        /// <summary>
        /// Maximum valid OAdate value (9999-12-31)
        /// </summary>
        public static readonly double MAX_OADATE_VALUE = 2958465.999988426d;
        /// <summary>
        /// First date that can be displayed by Excel. Real values before this date cannot be processed.
        /// </summary>
        public static readonly DateTime FIRST_ALLOWED_EXCEL_DATE = new DateTime(1900, 1, 1, 0, 0, 0);
        /// <summary>
        /// Last date that can be displayed by Excel. Real values after this date cannot be processed.
        /// </summary>
        public static readonly DateTime LAST_ALLOWED_EXCEL_DATE = new DateTime(9999, 12, 31, 23, 59, 59);

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

        private const float COLUMN_WIDTH_ROUNDING_MODIFIER = 256f;
        private const float SPLIT_WIDTH_MULTIPLIER = 12f;
        private const float SPLIT_WIDTH_OFFSET = 0.5f;
        private const float SPLIT_WIDTH_POINT_MULTIPLIER = 3f / 4f;
        private const float SPLIT_POINT_DIVIDER = 20f;
        private const float SPLIT_WIDTH_POINT_OFFSET = 390f;
        private const float SPLIT_HEIGHT_POINT_OFFSET = 300f;
        private const float ROW_HEIGHT_POINT_MULTIPLIER = 1f / 3f + 1f;
        private static readonly DateTime ROOT_DATE = new DateTime(1899, 12, 30, 0, 0, 0);
        private static readonly double ROOT_MILLIS = (double)new DateTime(1899, 12, 30, 0, 0, 0).Ticks / TimeSpan.TicksPerMillisecond;

        #endregion

        /// <summary>
        /// Method to convert a date or date and time into the internal Excel time format (OAdate)
        /// </summary>
        /// <param name="date">Date to process</param>
        /// <returns>Date or date and time as number string</returns>
        /// <exception cref="NanoXLSX.Shared.Exceptions.FormatException">Throws a FormatException if the passed date cannot be translated to the OADate format</exception>
        /// <remarks>Excel assumes wrongly that the year 1900 is a leap year. There is a gap of 1.0 between 1900-02-28 and 1900-03-01. This method corrects all dates
        /// from the first valid date (1900-01-01) to 1900-03-01. However, Excel displays the minimum valid date as 1900-01-00, although 0 is not a valid description for a day of month.
        /// In conformance to the OAdate specifications, the maximum valid date is 9999-12-31 23:59:59 (plus 999 milliseconds).<br/>
        ///See also: <a href="https://docs.microsoft.com/en-us/dotnet/api/system.datetime.tooadate?view=netcore-3.1">
        ///https://docs.microsoft.com/en-us/dotnet/api/system.datetime.tooadate?view=netcore-3.1</a><br/>
        ///See also: <a href="https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year">
        ///https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year</a>
        /// </remarks>
        public static string GetOADateTimeString(DateTime date)
        {
            double d = GetOADateTime(date);
            return d.ToString("G", INVARIANT_CULTURE);
        }

        /// <summary>
        /// Method to convert a date or date and time into the internal Excel time format (OAdate)
        /// </summary>
        /// <param name="skipCheck">Optional flag to skip the validity check if set to true</param>
        /// <param name="date">Date to process</param>
        /// <returns>Date or date and time as number</returns>
        /// <exception cref="NanoXLSX.Shared.Exceptions.FormatException">Throws a FormatException if the passed date cannot be translated to the OADate format</exception>
        /// <remarks>Excel assumes wrongly that the year 1900 is a leap year. There is a gap of 1.0 between 1900-02-28 and 1900-03-01. This method corrects all dates
        /// from the first valid date (1900-01-01) to 1900-03-01. However, Excel displays the minimum valid date as 1900-01-00, although 0 is not a valid description for a day of month.
        /// In conformance to the OAdate specifications, the maximum valid date is 9999-12-31 23:59:59 (plus 999 milliseconds).<br/>
        ///See also: <a href="https://docs.microsoft.com/en-us/dotnet/api/system.datetime.tooadate?view=netcore-3.1">
        ///https://docs.microsoft.com/en-us/dotnet/api/system.datetime.tooadate?view=netcore-3.1</a><br/>
        ///See also: <a href="https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year">
        ///https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year</a>
        /// </remarks>
        public static double GetOADateTime(DateTime date, bool skipCheck = false)
        {
            if (!skipCheck && (date < FIRST_ALLOWED_EXCEL_DATE || date > LAST_ALLOWED_EXCEL_DATE))
            {
                throw new FormatException("The date is not in a valid range for Excel. Dates before 1900-01-01 or after 9999-12-31 are not allowed.");
            }
            DateTime dateValue = date;
            if (date < FIRST_VALID_EXCEL_DATE)
            {
                dateValue = date.AddDays(-1); // Fix of the leap-year-1900-error
            }
            double currentMillis = (double)dateValue.Ticks / TimeSpan.TicksPerMillisecond;
            return ((dateValue.Second + (dateValue.Minute * 60) + (dateValue.Hour * 3600)) / 86400d) + Math.Floor((currentMillis - ROOT_MILLIS) / 86400000d);
        }

        /// <summary>
        /// Method to convert a time into the internal Excel time format (OAdate without days)
        /// </summary>
        /// <param name="time">Time to process. The date component of the timespan is converted to the total numbers of days</param>
        /// <returns>Time as number string</returns>
        /// <remarks>The time is represented by a OAdate without the date component but a possible number of total days</remarks>
        public static string GetOATimeString(TimeSpan time)
        {
            double d = GetOATime(time);
            return d.ToString("G", INVARIANT_CULTURE);
        }

        /// <summary>
        /// Method to convert a time into the internal Excel time format (OAdate without days)
        /// </summary>
        /// <param name="time">Time to process. The date component of the timespan is converted to the total numbers of days</param>
        /// <returns>Time as number</returns>
        /// <remarks>The time is represented by a OAdate without the date component but a possible number of total days</remarks>
        public static double GetOATime(TimeSpan time)
        {
            int seconds = time.Seconds + time.Minutes * 60 + time.Hours * 3600;
            return time.Days + (double)seconds / 86400d;
        }

        /// <summary>
        /// Method to calculate a common Date from the OA date (OLE automation) format<br/>
        /// OA Date format starts at January 1st 1900 (actually 00.01.1900). Dates beyond this date cannot be handled by Excel under normal circumstances and will throw a FormatException
        /// </summary>
        /// <param name="oaDate">oaDate OA date number</param>
        /// <returns>Converted date</returns>
        /// <remarks>Numbers that represents dates before 1900-03-01 (number of days since 1900-01-01 = 60) are automatically modified.
        /// Until 1900-03-01 is 1.0 added to the number to get the same date, as displayed in Excel.The reason for this is a bug in Excel.
        /// See also: <a href="https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year">
        /// https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year</a></remarks>
        public static DateTime GetDateFromOA(double oaDate)
        {
            if (oaDate < 60)
            {
                oaDate = oaDate + 1;
            }
            return ROOT_DATE.AddSeconds(oaDate * 86400d);
        }

        /// <summary>
        /// Calculates the internal width of a column in characters. This width is used only in the XML documents of worksheets and is usually not exposed to the (Excel) end user
        /// </summary>
        /// <remarks>
        /// The internal width deviates slightly from the column width, entered in Excel. Although internal, the default column width of 10 characters is visible in Excel as 10.71.
        /// The deviation depends on the maximum digit width of the default font, as well as its text padding and various constants.<br/>
        /// In case of the width 10.0 and the default digit width 7.0, as well as the padding 5.0 of the default font Calibri (size 11), 
        /// the internal width is approximately 10.7142857 (rounded to 10.71).<br/> Note that the column height is not affected by this consideration. 
        /// The entered height in Excel is the actual height in the worksheet XML documents.<br/> 
        /// This method is derived from the Perl implementation by John McNamara (<a href="https://stackoverflow.com/a/5010899">https://stackoverflow.com/a/5010899</a>)<br/>
        /// See also: <a href="https://www.ecma-international.org/publications-and-standards/standards/ecma-376/">ECMA-376, Part 1, Chapter 18.3.1.13</a>
        /// </remarks>
        /// <param name="columnWidth">Target column width (displayed in Excel)</param>
        /// <param name="maxDigitWidth">Maximum digit with of the default font (default is 7.0 for Calibri, size 11)</param>
        /// <param name="textPadding">Text padding of the default font (default is 5.0 for Calibri, size 11)</param>
        /// <returns>The internal column width in characters, used in worksheet XML documents</returns>
        /// <exception cref="FormatException">Throws a FormatException if the column width is out of range</exception>
        public static float GetInternalColumnWidth(float columnWidth, float maxDigitWidth = 7f, float textPadding = 5f)
        {
            if (columnWidth < Worksheet.MIN_COLUMN_WIDTH || columnWidth > Worksheet.MAX_COLUMN_WIDTH)
            {
                throw new FormatException("The column width " + columnWidth + " is not valid. The valid range is between " + Worksheet.MIN_COLUMN_WIDTH + " and " + Worksheet.MAX_COLUMN_WIDTH);
            }
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
        /// Calculates the internal height of a row. This height is used only in the XML documents of worksheets and is usually not exposed to the (Excel) end user
        /// </summary>
        /// <remarks>The height is based on the calculated amount of pixels. One point are ~1.333 (1+1/3) pixels. 
        /// After the conversion, the number of pixels is rounded to the nearest integer and calculated back to points.<br/>
        /// Therefore, the originally defined row height will slightly deviate, based on this pixel snap</remarks>
        /// <param name="rowHeight">Target row height (displayed in Excel)</param>
        /// <returns>The internal row height which snaps to the nearest pixel</returns>
        /// <exception cref="FormatException">Throws a FormatException if the row height is out of range</exception>
        public static float GetInternalRowHeight(float rowHeight)
        {
            if (rowHeight < Worksheet.MIN_ROW_HEIGHT || rowHeight > Worksheet.MAX_ROW_HEIGHT)
            {
                throw new FormatException("The row height " + rowHeight + " is not valid. The valid range is between " + Worksheet.MIN_ROW_HEIGHT + " and " + Worksheet.MAX_ROW_HEIGHT);
            }
            if (rowHeight == 0f)
            {
                return 0f;
            }
            double heightInPixel = Math.Round(rowHeight * ROW_HEIGHT_POINT_MULTIPLIER);
            return (float)heightInPixel / ROW_HEIGHT_POINT_MULTIPLIER;
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
        /// The two optional parameters maxDigitWidth and textPadding probably don't have to be changed ever. Negative column widths are automatically transformed to 0.
        /// </remarks>
        /// <param name="width">Target column(s) width (one or more columns, displayed in Excel)</param>
        /// <param name="maxDigitWidth">Maximum digit with of the default font (default is 7.0 for Calibri, size 11)</param>
        /// <param name="textPadding">Text padding of the default font (default is 5.0 for Calibri, size 11)</param>
        /// <returns>The internal pane width, used in worksheet XML documents in case of worksheet splitting</returns>
        public static float GetInternalPaneSplitWidth(float width, float maxDigitWidth = 7f, float textPadding = 5f)
        {
            float pixels;
            // TODO: Check the <1 part again. Leads always to 390
            if (width <= 1f)
            {
                width = 0;
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
        /// This method is derived from the Perl implementation by John McNamara (<a href="https://stackoverflow.com/a/5010899">https://stackoverflow.com/a/5010899</a>).<br/>
        /// Negative row heights are automatically transformed to 0.
        /// </remarks>
        /// <param name="height">Target row(s) height (one or more rows, displayed in Excel)</param>
        /// <returns>The internal pane height, used in worksheet XML documents in case of worksheet splitting</returns>
        public static float GetInternalPaneSplitHeight(float height)
        {
            if (height < 0)
            {
                height = 0f;
            }
            return (float)Math.Floor(SPLIT_POINT_DIVIDER * height + SPLIT_HEIGHT_POINT_OFFSET);
        }

        /// <summary>
        /// Calculates the height of a split pane in a worksheet, based on the internal value (calculated by <see cref="GetInternalPaneSplitHeight(float)"/>)
        /// </summary>
        /// <param name="internalHeight">Internal pane height stored in a worksheet. The minimal value is defined by <see cref="SPLIT_HEIGHT_POINT_OFFSET"/></param>
        /// <returns>Actual pane height</returns>
        /// <remarks>Depending on the initial height, the result value of <see cref="GetInternalPaneSplitHeight(float)"/> may not lead back to the initial value, 
        /// since rounding is applied when calculating the internal height</remarks>
        public static float GetPaneSplitHeight(float internalHeight)
        {
            if (internalHeight < 300f)
            {
                return 0;
            }
            else
            {
                return (internalHeight - SPLIT_HEIGHT_POINT_OFFSET) / SPLIT_POINT_DIVIDER;
            }
        }

        /// <summary>
        /// Calculates the width of a split pane in a worksheet, based on the internal value (calculated by <see cref="GetInternalPaneSplitWidth(float, float, float)"/>)
        /// </summary>
        /// <param name="internalWidth">Internal pane width stored in a worksheet. The minimal value is defined by <see cref="SPLIT_WIDTH_POINT_OFFSET"/></param>
        /// <param name="maxDigitWidth">Maximum digit with of the default font (default is 7.0 for Calibri, size 11)</param>
        /// <param name="textPadding">Text padding of the default font (default is 5.0 for Calibri, size 11)</param>
        /// <returns>Actual pane width</returns>
        /// <remarks>Depending on the initial width, the result value of <see cref="GetInternalPaneSplitWidth(float,float,float)"/> may not lead back to the initial value, 
        /// since rounding is applied when calculating the internal width</remarks>
        public static float GetPaneSplitWidth(float internalWidth, float maxDigitWidth = 7f, float textPadding = 5f)
        {
            float points = (internalWidth - SPLIT_WIDTH_POINT_OFFSET) / SPLIT_POINT_DIVIDER;
            if (points < 0.001f)
            {
                return 0;
            }
            else
            {
                float width = points / SPLIT_WIDTH_POINT_MULTIPLIER;
                return (width - textPadding - SPLIT_WIDTH_OFFSET) / maxDigitWidth;
            }
        }

        /// <summary>
        /// Method to generate an Excel internal password hash to protect workbooks or worksheets<br></br>This method is derived from the c++ implementation by Kohei Yoshida (<a href="http://kohei.us/2008/01/18/excel-sheet-protection-password-hash/">http://kohei.us/2008/01/18/excel-sheet-protection-password-hash/</a>)
        /// </summary>
        /// <remarks>WARNING! Do not use this method to encrypt 'real' passwords or data outside from NanoXLSX. This is only a minor security feature. Use a proper cryptography method instead.</remarks>
        /// <param name="password">Password string in UTF-8 to encrypt</param>
        /// <returns>16 bit hash as hex string</returns>
        public static string GeneratePasswordHash(string password)
        {
            if (string.IsNullOrEmpty(password)) { return string.Empty; }
            int passwordLength = password.Length;
            int passwordHash = 0;
            char character;
            for (int i = passwordLength; i > 0; i--)
            {
                character = password[i - 1];
                passwordHash = ((passwordHash >> 14) & 0x01) | ((passwordHash << 1) & 0x7fff);
                passwordHash ^= character;
            }
            passwordHash = ((passwordHash >> 14) & 0x01) | ((passwordHash << 1) & 0x7fff);
            passwordHash ^= (0x8000 | ('N' << 8) | 'K');
            passwordHash ^= passwordLength;
            return passwordHash.ToString("X");
        }
    }
}
