/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using FormatException = NanoXLSX.Exceptions.FormatException;

namespace NanoXLSX.Utils
{
    /// <summary>
    /// General data utils class with static methods
    /// </summary>
    public static class DataUtils

    {
        #region constants
        /// <summary>
        /// Minimum valid OAdate value (1900-01-01). However, Excel displays this value as 1900-01-00 (day zero)
        /// </summary>
#pragma warning disable CA1805 // Suppress: Do not initialize unnecessarily (to make clear that this is the minimum value)
        public static readonly double MinOADateValue = 0d;
#pragma warning restore CA1805
        /// <summary>
        /// Maximum valid OAdate value (9999-12-31)
        /// </summary>
        public static readonly double MaxOADateValue = 2958465.999988426d;
        /// <summary>
        /// First date that can be displayed by Excel. Real values before this date cannot be processed.
        /// </summary>
        public static readonly DateTime FirstAllowedExcelDate = new DateTime(1900, 1, 1, 0, 0, 0, DateTimeKind.Unspecified);
        /// <summary>
        /// Last date that can be displayed by Excel. Real values after this date cannot be processed.
        /// </summary>
        public static readonly DateTime LastAllowedExcelDate = new DateTime(9999, 12, 31, 23, 59, 59, DateTimeKind.Unspecified);

        /// <summary>
        /// All dates before this date are shifted in Excel by -1.0, since Excel assumes wrongly that the year 1900 is a leap year.<br />
        /// See also: <a href="https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year">
        /// https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year</a>
        /// </summary>
        public static readonly DateTime FirstValidExcelDate = new DateTime(1900, 3, 1, 0, 0, 0, DateTimeKind.Unspecified);

        /// <summary>
        /// Constant for number conversions. The invariant culture (represents mostly the US numbering scheme) ensures that no culture-specific 
        /// punctuations are used when converting numbers to strings, This is especially important for OOXML number values.
        /// See also: <a href="https://docs.microsoft.com/en-us/dotnet/api/system.globalization.cultureinfo.invariantculture?view=net-5.0">
        /// https://docs.microsoft.com/en-us/dotnet/api/system.globalization.cultureinfo.invariantculture?view=net-5.0</a>
        /// </summary>
        public static readonly CultureInfo InvariantCulture = CultureInfo.InvariantCulture;

        private const float COLUMN_WIDTH_ROUNDING_MODIFIER = 256f;
        private const float SPLIT_WIDTH_MULTIPLIER = 12f;
        private const float SPLIT_WIDTH_OFFSET = 0.5f;
        private const float SPLIT_WIDTH_POINT_MULTIPLIER = 3f / 4f;
        private const float SPLIT_POINT_DIVIDER = 20f;
        private const float SPLIT_WIDTH_POINT_OFFSET = 390f;
        private const float SPLIT_HEIGHT_POINT_OFFSET = 300f;
        private const float ROW_HEIGHT_POINT_MULTIPLIER = 1f / 3f + 1f;
        private static readonly DateTime ROOT_DATE = new DateTime(1899, 12, 30, 0, 0, 0, DateTimeKind.Unspecified);
        private static readonly double ROOT_MILLIS = (double)new DateTime(1899, 12, 30, 0, 0, 0, DateTimeKind.Unspecified).Ticks / TimeSpan.TicksPerMillisecond;

        #endregion

        #region enums
        /// <summary>
        /// Strategy how ranges should be merged
        /// </summary>
        public enum RangeMergeStrategy
        {
            /// <summary>
            /// No merge should be performed
            /// </summary>
            NoMerge,
            /// <summary>
            /// Ranges of the same columns should be merged
            /// </summary>
            MergeColumns,
            /// <summary>
            /// Ranges of the same row should be merged
            /// </summary>
            MergeRows
        }
        #endregion

        /// <summary>
        /// Method to convert a date or date and time into the internal Excel time format (OAdate)
        /// </summary>
        /// <param name="date">Date to process</param>
        /// <returns>Date or date and time as number string</returns>
        /// <exception cref="NanoXLSX.Exceptions.FormatException">Throws a FormatException if the passed date cannot be translated to the OADate format</exception>
        /// \remark <remarks>Excel assumes wrongly that the year 1900 is a leap year. There is a gap of 1.0 between 1900-02-28 and 1900-03-01. This method corrects all dates
        /// from the first valid date (1900-01-01) to 1900-03-01. However, Excel displays the minimum valid date as 1900-01-00, although 0 is not a valid description for a day of month.
        /// In conformance to the OAdate specifications, the maximum valid date is 9999-12-31 23:59:59 (plus 999 milliseconds).<br />
        ///See also: <a href="https://docs.microsoft.com/en-us/dotnet/api/system.datetime.tooadate?view=netcore-3.1">
        ///https://docs.microsoft.com/en-us/dotnet/api/system.datetime.tooadate?view=netcore-3.1</a><br />
        ///See also: <a href="https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year">
        ///https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year</a>
        /// </remarks>
        public static string GetOADateTimeString(DateTime date)
        {
            double d = GetOADateTime(date);
            return d.ToString("G", InvariantCulture);
        }

        /// <summary>
        /// Method to convert a date or date and time into the internal Excel time format (OAdate)
        /// </summary>
        /// <param name="skipCheck">Optional flag to skip the validity check if set to true</param>
        /// <param name="date">Date to process</param>
        /// <returns>Date or date and time as number</returns>
        /// <exception cref="NanoXLSX.Exceptions.FormatException">Throws a FormatException if the passed date cannot be translated to the OADate format</exception>
        /// \remark <remarks>Excel assumes wrongly that the year 1900 is a leap year. There is a gap of 1.0 between 1900-02-28 and 1900-03-01. This method corrects all dates
        /// from the first valid date (1900-01-01) to 1900-03-01. However, Excel displays the minimum valid date as 1900-01-00, although 0 is not a valid description for a day of month.
        /// In conformance to the OAdate specifications, the maximum valid date is 9999-12-31 23:59:59 (plus 999 milliseconds).<br />
        ///See also: <a href="https://docs.microsoft.com/en-us/dotnet/api/system.datetime.tooadate?view=netcore-3.1">
        ///https://docs.microsoft.com/en-us/dotnet/api/system.datetime.tooadate?view=netcore-3.1</a><br />
        ///See also: <a href="https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year">
        ///https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year</a>
        /// </remarks>
        public static double GetOADateTime(DateTime date, bool skipCheck = false)
        {
            if (!skipCheck && (date < FirstAllowedExcelDate || date > LastAllowedExcelDate))
            {
                throw new FormatException("The date is not in a valid range for Excel. Dates before 1900-01-01 or after 9999-12-31 are not allowed.");
            }
            DateTime dateValue = date;
            if (date < FirstValidExcelDate)
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
        /// \remark <remarks>The time is represented by a OAdate without the date component but a possible number of total days</remarks>
        public static string GetOATimeString(TimeSpan time)
        {
            double d = GetOATime(time);
            return d.ToString("G", InvariantCulture);
        }

        /// <summary>
        /// Method to convert a time into the internal Excel time format (OAdate without days)
        /// </summary>
        /// <param name="time">Time to process. The date component of the timespan is converted to the total numbers of days</param>
        /// <returns>Time as number</returns>
        /// \remark <remarks>The time is represented by a OAdate without the date component but a possible number of total days</remarks>
        public static double GetOATime(TimeSpan time)
        {
            int seconds = time.Seconds + time.Minutes * 60 + time.Hours * 3600;
            return time.Days + (double)seconds / 86400d;
        }

        /// <summary>
        /// Method to calculate a common Date from the OA date (OLE automation) format<br />
        /// OA Date format starts at January 1st 1900 (actually 00.01.1900). Dates beyond this date cannot be handled by Excel under normal circumstances and will throw a FormatException
        /// </summary>
        /// <param name="oaDate">oaDate OA date number</param>
        /// <returns>Converted date</returns>
        /// \remark <remarks>Numbers that represents dates before 1900-03-01 (number of days since 1900-01-01 = 60) are automatically modified.
        /// Until 1900-03-01 is 1.0 added to the number to get the same date, as displayed in Excel.The reason for this is a bug in Excel.
        /// See also: <a href="https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year">
        /// https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year</a></remarks>
        public static DateTime GetDateFromOA(double oaDate)
        {
            if (oaDate < 60)
            {
                oaDate++;
            }
            return ROOT_DATE.AddSeconds(oaDate * 86400d);
        }

        /// <summary>
        /// Calculates the internal width of a column in characters. This width is used only in the XML documents of worksheets and is usually not exposed to the (Excel) end user
        /// </summary>
        /// \remark <remarks>
        /// The internal width deviates slightly from the column width, entered in Excel. Although internal, the default column width of 10 characters is visible in Excel as 10.71.
        /// The deviation depends on the maximum digit width of the default font, as well as its text padding and various constants.<br />
        /// In case of the width 10.0 and the default digit width 7.0, as well as the padding 5.0 of the default font Calibri (size 11), 
        /// the internal width is approximately 10.7142857 (rounded to 10.71).<br /> Note that the column height is not affected by this consideration. 
        /// The entered height in Excel is the actual height in the worksheet XML documents.<br /> 
        /// This method is derived from the Perl implementation by John McNamara (<a href="https://stackoverflow.com/a/5010899">https://stackoverflow.com/a/5010899</a>)<br />
        /// See also: <a href="https://www.ecma-international.org/publications-and-standards/standards/ecma-376/">ECMA-376, Part 1, Chapter 18.3.1.13</a>
        /// </remarks>
        /// <param name="columnWidth">Target column width (displayed in Excel)</param>
        /// <param name="maxDigitWidth">Maximum digit with of the default font (default is 7.0 for Calibri, size 11)</param>
        /// <param name="textPadding">Text padding of the default font (default is 5.0 for Calibri, size 11)</param>
        /// <returns>The internal column width in characters, used in worksheet XML documents</returns>
        /// <exception cref="FormatException">Throws a FormatException if the column width is out of range</exception>
        public static float GetInternalColumnWidth(float columnWidth, float maxDigitWidth = 7f, float textPadding = 5f)
        {
            if (columnWidth < Worksheet.MinColumnWidth || columnWidth > Worksheet.MaxColumnWidth)
            {
                throw new FormatException("The column width " + columnWidth + " is not valid. The valid range is between " + Worksheet.MinColumnWidth + " and " + Worksheet.MaxColumnWidth);
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
        /// \remark <remarks>The height is based on the calculated amount of pixels. One point are ~1.333 (1+1/3) pixels. 
        /// After the conversion, the number of pixels is rounded to the nearest integer and calculated back to points.<br />
        /// Therefore, the originally defined row height will slightly deviate, based on this pixel snap</remarks>
        /// <param name="rowHeight">Target row height (displayed in Excel)</param>
        /// <returns>The internal row height which snaps to the nearest pixel</returns>
        /// <exception cref="FormatException">Throws a FormatException if the row height is out of range</exception>
        public static float GetInternalRowHeight(float rowHeight)
        {
            if (rowHeight < Worksheet.MinRowHeight || rowHeight > Worksheet.MaxRowHeight)
            {
                throw new FormatException("The row height " + rowHeight + " is not valid. The valid range is between " + Worksheet.MinRowHeight + " and " + Worksheet.MaxRowHeight);
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
        /// \remark <remarks>
        /// The internal split width is based on the width of one or more columns. 
        /// It also depends on the maximum digit width of the default font, as well as its text padding and various constants.<br />
        /// See also <see cref="GetInternalColumnWidth(float, float, float)"/> for additional details.<br />
        /// This method is derived from the Perl implementation by John McNamara (<a href="https://stackoverflow.com/a/5010899">https://stackoverflow.com/a/5010899</a>)<br />
        /// See also: <a href="https://www.ecma-international.org/publications-and-standards/standards/ecma-376/">ECMA-376, Part 1, Chapter 18.3.1.13</a><br />
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
        /// \remark <remarks>
        /// The internal split height is based on the height of one or more rows. It also depends on various constants.<br />
        /// This method is derived from the Perl implementation by John McNamara (<a href="https://stackoverflow.com/a/5010899">https://stackoverflow.com/a/5010899</a>).<br />
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
        /// \remark <remarks>Depending on the initial height, the result value of <see cref="GetInternalPaneSplitHeight(float)"/> may not lead back to the initial value, 
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
        /// \remark <remarks>Depending on the initial width, the result value of <see cref="GetInternalPaneSplitWidth(float,float,float)"/> may not lead back to the initial value, 
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
        /// Merges a range with a list of given ranges. If there is no intersection between the list and the new range, the range is just added to the given list. If there is an 
        /// intersection, the range will be merged and the new list of ranges will be returned
        /// </summary>
        /// <param name="givenRanges">List of given ranges</param>
        /// <param name="newRange">The range to be merged</param>
        /// <param name="strategy">Strategy for the range recalculation. Depending on the value, the resulting ranges are either merged along rows, along columns (default), or not merged at all</param>
        /// <returns>List of resulting ranges after merging.</returns>
        public static IReadOnlyList<Range> MergeRange(List<Range> givenRanges, Range newRange, RangeMergeStrategy strategy = RangeMergeStrategy.MergeColumns)
        {
            List<Range> result = new List<Range>();
            List<Range> mergedCandidates = new List<Range> { newRange };
            // Step 1: Find intersecting ranges and remove them from existingRanges
            foreach (Range range in givenRanges)
            {
                if (IsMergeCandidate(newRange, range, strategy))
                {
                    mergedCandidates.Add(range);
                }
                else
                {
                    result.Add(range);
                }
            }
            // Step 2: Slice intersecting/adjacent ranges into uniform rectangular pieces.
            List<Range> slicedRanges = SliceRanges(mergedCandidates);
            // Step 3: Merge adjacent rectangles where possible.
            if (strategy == RangeMergeStrategy.MergeColumns)
            {
                result.AddRange(MergeAdjacentRanges(slicedRanges, RangeMergeStrategy.MergeColumns));
                result = MergeAdjacentRanges(result, RangeMergeStrategy.MergeRows);
            }
            else if (strategy == RangeMergeStrategy.MergeRows)
            {
                result.AddRange(MergeAdjacentRanges(slicedRanges, RangeMergeStrategy.MergeRows));
                result = MergeAdjacentRanges(result, RangeMergeStrategy.MergeColumns);
            }
            else
            {
                result.AddRange(slicedRanges);
            }
            return result;
        }

        /// <summary>
        /// Returns true if the two ranges are either overlapping or adjacent in the appropriate direction
        /// for the chosen merge strategy.
        /// For vertical merging (MergeColumns): the ranges must share the same column boundaries and be either overlapping 
        /// or immediately adjacent vertically.
        /// For horizontal merging (MergeRows): the ranges must share the same row boundaries and be either overlapping 
        /// or immediately adjacent horizontally.
        /// </summary>
        private static bool IsMergeCandidate(Range a, Range b, RangeMergeStrategy strategy)
        {
            // First, if they overlap, they are candidates.
            if (a.Overlaps(b))
            {
                return true;
            }

            // Otherwise, check for adjacency according to the strategy.
            if (strategy == RangeMergeStrategy.MergeColumns)
            {
                // Vertical merging: require same columns.
                if (a.StartAddress.Column == b.StartAddress.Column &&
                    a.EndAddress.Column == b.EndAddress.Column &&
                    (a.EndAddress.Row + 1 == b.StartAddress.Row || b.EndAddress.Row + 1 == a.StartAddress.Row))
                {
                    return true;
                }
            }
            else if (strategy == RangeMergeStrategy.MergeRows)
            {
                // Horizontal merging: require same rows.
                if (a.StartAddress.Row == b.StartAddress.Row &&
                    a.EndAddress.Row == b.EndAddress.Row &&
                    (a.EndAddress.Column + 1 == b.StartAddress.Column || b.EndAddress.Column + 1 == a.StartAddress.Column))
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Subtracts a range form a list of given ranges. If the range to be removed does not intersect any of the given ranges, nothing happens. If the range intersects 
        /// at least one of the given ranges, the intersection will be removed and the new ranges well be returned.
        /// </summary>
        /// <param name="givenRanges">List of given ranges</param>
        /// <param name="rangeToRemove">The range to be removed</param>
        /// <param name="strategy">Strategy for the range recalculation. Depending on the value, the resulting ranges are either merged along rows, along columns (default), or not merged at all</param>
        /// <returns>List of resulting ranges after subtraction and recalculation</returns>
        public static IReadOnlyList<Range> SubtractRange(List<Range> givenRanges, Range rangeToRemove, RangeMergeStrategy strategy = RangeMergeStrategy.MergeColumns)
        {
            List<Range> result = new List<Range>();
            // Process each existing range.
            foreach (Range range in givenRanges)
            {
                if (!range.Overlaps(rangeToRemove))
                {
                    // No overlap: keep the range unchanged.
                    result.Add(range);
                }
                else
                {
                    // Overlapping range: subtract the removal area.
                    List<Range> subtractedPieces = SubtractRect(range, rangeToRemove);
                    result.AddRange(subtractedPieces);
                }
            }
            // Slice all ranges before merge
            List<Range> slicedRanges = SliceRanges(result);
            // Merge adjacent pieces if requested.
            if (strategy == RangeMergeStrategy.MergeColumns)
            {
                result = MergeAdjacentRanges(slicedRanges, RangeMergeStrategy.MergeColumns);
                result = MergeAdjacentRanges(result, RangeMergeStrategy.MergeRows);
            }
            else if (strategy == RangeMergeStrategy.MergeRows)
            {
                result = MergeAdjacentRanges(slicedRanges, RangeMergeStrategy.MergeRows);
                result = MergeAdjacentRanges(result, RangeMergeStrategy.MergeColumns);
            }
            else
            {
                result = slicedRanges;
            }
            return result;
        }

        /// <summary>
        /// Method to slice possibly overlapping ranges into contiguous ranges without intersections
        /// </summary>
        /// <param name="ranges">Ranges to slice</param>
        /// <returns>List of sliced, contiguous ranges</returns>
        private static List<Range> SliceRanges(List<Range> ranges)
        {
            HashSet<int> uniqueCols = new HashSet<int>();
            HashSet<int> uniqueRows = new HashSet<int>();

            // Collect all column and row boundaries
            foreach (Range range in ranges)
            {
                uniqueCols.Add(range.StartAddress.Column);
                uniqueCols.Add(range.EndAddress.Column + 1); // To handle gaps properly
                uniqueRows.Add(range.StartAddress.Row);
                uniqueRows.Add(range.EndAddress.Row + 1);
            }

            // Convert to sorted lists for iteration
            List<int> sortedCols = uniqueCols.OrderBy(c => c).ToList();
            List<int> sortedRows = uniqueRows.OrderBy(r => r).ToList();

            List<Range> slicedRanges = new List<Range>();

            // Step through the row and column boundaries to create the smallest sub-rectangles
            for (int r = 0; r < sortedRows.Count - 1; r++)
            {
                for (int c = 0; c < sortedCols.Count - 1; c++)
                {
                    Range subRange = new Range(
                        sortedCols[c],
                        sortedRows[r],
                        sortedCols[c + 1] - 1,
                        sortedRows[r + 1] - 1
                    );

                    // Only keep the sub-range if it was originally covered
                    if (ranges.Exists(range => range.Contains(subRange)))
                    {
                        slicedRanges.Add(subRange);
                    }
                }
            }
            return slicedRanges;
        }

        /// <summary>
        /// Subtracts the removal range from an original range.
        /// Returns up to 4 rectangular pieces that cover (original minus the intersecting part).
        /// If there is no intersection, returns the original range.
        /// </summary>
        private static List<Range> SubtractRect(Range original, Range toRemove)
        {
            List<Range> pieces = new List<Range>();
            // Original boundaries:
            int orig_left = original.StartAddress.Column;
            int orig_top = original.StartAddress.Row;
            int orig_right = original.EndAddress.Column;
            int orig_bottom = original.EndAddress.Row;
            // Removal boundaries:
            int rem_left = toRemove.StartAddress.Column;
            int rem_top = toRemove.StartAddress.Row;
            int rem_right = toRemove.EndAddress.Column;
            int rem_bottom = toRemove.EndAddress.Row;
            // Compute intersection boundaries.
            int isct_left = Math.Max(orig_left, rem_left);
            int isct_top = Math.Max(orig_top, rem_top);
            int isct_right = Math.Min(orig_right, rem_right);
            int isct_bottom = Math.Min(orig_bottom, rem_bottom);

            // Slice the original rectangle into up to four pieces.
            // Top piece: if any rows exist above the intersection.
            if (orig_top < isct_top)
            {
                pieces.Add(new Range(orig_left, orig_top, orig_right, isct_top - 1));
            }
            // Bottom piece: if any rows exist below the intersection.
            if (isct_bottom < orig_bottom)
            {
                pieces.Add(new Range(orig_left, isct_bottom + 1, orig_right, orig_bottom));
            }
            // Left piece: if any columns exist to the left of the intersection within the vertical boundaries of the intersection.
            if (orig_left < isct_left)
            {
                pieces.Add(new Range(orig_left, isct_top, isct_left - 1, isct_bottom));
            }
            // Right piece: if any columns exist to the right of the intersection within the vertical boundaries of the intersection.
            if (isct_right < orig_right)
            {
                pieces.Add(new Range(isct_right + 1, isct_top, orig_right, isct_bottom));
            }
            return pieces;
        }

        /// <summary>
        /// Method to merge ranges by rows or columns (according to the strategy) that can be merged into a new ranges, 
        /// so that all addresses of the original ranges are still covered and no additional addresses are used.
        /// </summary>
        /// <param name="ranges">Original (sliced) ranges</param>
        /// <param name="strategy">Merge strategy. If the strategy is NoMerge, the original list will be returned</param>
        /// <returns>List of merged ranges</returns>
        private static List<Range> MergeAdjacentRanges(List<Range> ranges, RangeMergeStrategy strategy)
        {
            if (ranges.Count == 0)
            {
                return new List<Range>();
            }
            List<Range> mergedRanges = new List<Range>();

            if (strategy == RangeMergeStrategy.MergeColumns)
            {
                // Vertical merging: Ranges must have identical column boundaries.
                // Group by StartAddress.Column and EndAddress.Column.
                var groups = ranges.GroupBy(r => new
                {
                    StartCol = r.StartAddress.Column,
                    EndCol = r.EndAddress.Column
                });
                foreach (var group in groups)
                {
                    // Order by row (ascending)
                    List<Range> sorted = group.OrderBy(r => r.StartAddress.Row).ToList();
                    Range current = sorted[0];
                    for (int i = 1; i < sorted.Count; i++)
                    {
                        Range next = sorted[i];
                        // Check if the current range is contiguous with or overlapping the next.
                        // (That is, if current.EndAddress.Row + 1 is >= next.StartAddress.Row.)
                        if (current.EndAddress.Row + 1 >= next.StartAddress.Row)
                        {
                            // They share the same columns.
                            // Create a new range from current.StartAddress.Row to the maximum of the two EndAddress.Row values.
                            int newStartRow = current.StartAddress.Row;
                            int newEndRow = Math.Max(current.EndAddress.Row, next.EndAddress.Row);
                            current = new Range(
                                current.StartAddress.Column, newStartRow,
                                current.EndAddress.Column, newEndRow);
                        }
                        else
                        {
                            mergedRanges.Add(current);
                            current = next;
                        }
                    }
                    mergedRanges.Add(current);
                }
            }
            else if (strategy == RangeMergeStrategy.MergeRows)
            {
                // Horizontal merging: Ranges must have identical row boundaries.
                // Group by StartAddress.Row and EndAddress.Row.
                var groups = ranges.GroupBy(r => new
                {
                    StartRow = r.StartAddress.Row,
                    EndRow = r.EndAddress.Row
                });
                foreach (var group in groups)
                {
                    var sorted = group.OrderBy(r => r.StartAddress.Column).ToList();
                    Range current = sorted[0];
                    for (int i = 1; i < sorted.Count; i++)
                    {
                        Range next = sorted[i];
                        // Check if current.EndAddress.Column + 1 is >= next.StartAddress.Column.
                        if (current.EndAddress.Column + 1 >= next.StartAddress.Column)
                        {
                            int newStartCol = current.StartAddress.Column;
                            int newEndCol = Math.Max(current.EndAddress.Column, next.EndAddress.Column);
                            current = new Range(
                                newStartCol, current.StartAddress.Row,
                                newEndCol, current.EndAddress.Row);
                        }
                        else
                        {
                            mergedRanges.Add(current);
                            current = next;
                        }
                    }
                    mergedRanges.Add(current);
                }
            }
            return mergedRanges;
        }

    }
}
