/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2021
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Globalization;
using NanoXLSX.Exceptions;


namespace NanoXLSX
{
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

        public static readonly CultureInfo INVARIANT_CULTURE = CultureInfo.InvariantCulture;

        #endregion

        /// <summary>
        /// Method to convert a date or date and time into the internal Excel time format (OAdate)
        /// </summary>
        /// <param name="date">Date to process</param>
        /// <param name="culture">CultureInfo for proper formatting of the decimal point</param>
        /// <returns>Date or date and time as number</returns>
        /// <exception cref="Exceptions.FormatException">Throws a FormatException if the passed date cannot be translated to the OADate format</exception>
        /// <remarks>OAdate format starts at January 1st 1900 (actually 00.01.1900) and ends at December 31 9999. 
        /// Values beyond these dates cannot be handled by Excel under normal circumstances and will throw a FormatException</remarks>
        public static string GetOADateTimeString(DateTime date, IFormatProvider culture)
        {
            try
            {
                double d = date.ToOADate();
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

    }
}
