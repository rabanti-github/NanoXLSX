using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using NanoXLSX.Utils;
using Xunit;
using FormatException = NanoXLSX.Exceptions.FormatException;

namespace NanoXLSX.Test.Core.MiscTest
{
    public class UtilsTest
    {
        [Theory(DisplayName = "Test of the GetOADateTimestring function")]
        [InlineData("01.01.1900 00:00:00", "1")]
        [InlineData("02.01.1900 12:35:20", "2.5245370370370401")]
        [InlineData("27.02.1900 00:00:00", "58")]
        [InlineData("28.02.1900 00:00:00", "59")]
        [InlineData("28.02.1900 12:30:32", "59.521203703703705")]
        [InlineData("01.03.1900 00:00:00", "61")]
        [InlineData("01.03.1900 08:08:11", "61.339016203703707")]
        [InlineData("20.05.1960 22:11:05", "22056.924363425926")]
        [InlineData("01.01.2021 00:00:00", "44197")]
        [InlineData("12.12.5870 11:30:12", "1450360.47930556")]
        public void GetOADateTimeStringTest(string dateString, string expectedOaDate)
        {
            CultureInfo provider = CultureInfo.InvariantCulture;
            string format = "dd.MM.yyyy HH:mm:ss";
            DateTime date = DateTime.ParseExact(dateString, format, provider);
            string oaDate = DataUtils.GetOADateTimeString(date);
            float expected = float.Parse(expectedOaDate);
            float given = float.Parse(oaDate);
            float threshold = 0.000000001f; // Ignore everything below a millisecond
            Assert.True(Math.Abs(expected - given) < threshold);
        }

        [Theory(DisplayName = "Test of the GetOADateTime function")]
        [InlineData("01.01.1900 00:00:00", 1d)]
        [InlineData("02.01.1900 12:35:20", 2.5245370370370401d)]
        [InlineData("27.02.1900 00:00:00", 58d)]
        [InlineData("28.02.1900 00:00:00", 59d)]
        [InlineData("28.02.1900 12:30:32", 59.521203703703705d)]
        [InlineData("01.03.1900 00:00:00", 61d)]
        [InlineData("01.03.1900 08:08:11", 61.339016203703707d)]
        [InlineData("20.05.1960 22:11:05", 22056.924363425926d)]
        [InlineData("01.01.2021 00:00:00", 44197d)]
        [InlineData("12.12.5870 11:30:12", 1450360.47930556d)]
        public void GetOADateTimeString(string dateString, double expectedOaDate)
        {
            CultureInfo provider = CultureInfo.InvariantCulture;
            string format = "dd.MM.yyyy HH:mm:ss";
            DateTime date = DateTime.ParseExact(dateString, format, provider);
            double oaDate = DataUtils.GetOADateTime(date);
            float threshold = 0.00000001f; // Ignore everything below a millisecond (double precision may vary)
            Assert.True(Math.Abs(expectedOaDate - oaDate) < threshold);
        }

        [Theory(DisplayName = "Test of the successful GetOADateTime function on invalid dates when checks are disabled")]
        [InlineData("02.01.0001 00:00:00")] // The DateTime does not allow negative values. The leap year fix would lead to this on 1.1.0001
        [InlineData("18.05.0712 11:15:02")]
        [InlineData("31.12.1899 23:59:59")]
        public void GetOADateTimeTest2(string dateString)
        {
            CultureInfo provider = CultureInfo.InvariantCulture;
            string format = "dd.MM.yyyy HH:mm:ss";
            DateTime date = DateTime.ParseExact(dateString, format, provider);
            double given = DataUtils.GetOADateTime(date, true);
            Assert.NotEqual(0d, given);
        }

        [Theory(DisplayName = "Test of the failing GetOADateTimeString function on invalid dates")]
        [InlineData("01.01.0001 00:00:00")]
        [InlineData("18.05.0712 11:15:02")]
        [InlineData("31.12.1899 23:59:59")]
        public void GetOADateTimeStringFailTest(string dateString)
        {
            // Note: Dates beyond the year 10000 cannot be tested though
            CultureInfo provider = CultureInfo.InvariantCulture;
            string format = "dd.MM.yyyy HH:mm:ss";
            DateTime date = DateTime.ParseExact(dateString, format, provider);
            Assert.Throws<FormatException>(() => DataUtils.GetOADateTimeString(date));
        }

        [Theory(DisplayName = "Test of the failing GetOADateTime function on invalid dates")]
        [InlineData("01.01.0001 00:00:00")]
        [InlineData("18.05.0712 11:15:02")]
        [InlineData("31.12.1899 23:59:59")]
        public void GetOADateTimeFailTest(string dateString)
        {
            // Note: Dates beyond the year 10000 cannot be tested though
            CultureInfo provider = CultureInfo.InvariantCulture;
            string format = "dd.MM.yyyy HH:mm:ss";
            DateTime date = DateTime.ParseExact(dateString, format, provider);
            Assert.Throws<FormatException>(() => DataUtils.GetOADateTime(date));
        }

        [Theory(DisplayName = "Test of the GetOATimeString function")]
        [InlineData("00:00:00", "0.0")]
        [InlineData("12:00:00", "0.5")]
        [InlineData("23:59:59", "0.999988425925926")]
        [InlineData("13:11:10", "0.549421296296296")]
        [InlineData("18:00:00", "0.75")]
        public void GetOATimeStringTest(string timeString, string expectedOaTime)
        {
            CultureInfo provider = CultureInfo.InvariantCulture;
            string format = "hh\\:mm\\:ss";
            TimeSpan time = TimeSpan.ParseExact(timeString, format, provider);
            string oaDate = DataUtils.GetOATimeString(time);
            float expected = float.Parse(expectedOaTime);
            float given = float.Parse(oaDate);
            float threshold = 0.000000001f; // Ignore everything below a millisecond
            Assert.True(Math.Abs(expected - given) < threshold);
        }

        [Theory(DisplayName = "Test of the GetOATime function")]
        [InlineData("00:00:00", 0.0d)]
        [InlineData("12:00:00", 0.5d)]
        [InlineData("23:59:59", 0.999988425925926d)]
        [InlineData("13:11:10", 0.549421296296296d)]
        [InlineData("18:00:00", 0.75d)]
        public void GetOATimeTest(string timeString, double expectedOaTime)
        {
            CultureInfo provider = CultureInfo.InvariantCulture;
            string format = "hh\\:mm\\:ss";
            TimeSpan time = TimeSpan.ParseExact(timeString, format, provider);
            double oaTime = DataUtils.GetOATime(time);
            float threshold = 0.000000001f; // Ignore everything below a millisecond
            Assert.True(Math.Abs(expectedOaTime - oaTime) < threshold);
        }

        [Theory(DisplayName = "Test of the GetInternalColumnWidth function")]
        [InlineData(0.5, 0.85546875)]
        [InlineData(1, 1.7109375)]
        [InlineData(10, 10.7109375)]
        [InlineData(15, 15.7109375)]
        [InlineData(60, 60.7109375)]
        [InlineData(254, 254.7109375)]
        [InlineData(255, 255.7109375)]
        [InlineData(0, 0f)]
        public void GetInternalColumnWidthTest(float width, float expectedInternalWidth)
        {
            float internalWidth = DataUtils.GetInternalColumnWidth(width);
            Assert.Equal(expectedInternalWidth, internalWidth);
        }

        [Theory(DisplayName = "Test of the failing GetInternalColumnWidth function on invalid column widths")]
        [InlineData(-0.1)]
        [InlineData(-10)]
        [InlineData(255.01)]
        [InlineData(10000)]
        public void GetInternalColumnWidthFailTest(float width)
        {
            Assert.Throws<FormatException>(() => DataUtils.GetInternalColumnWidth(width));
        }

        [Theory(DisplayName = "Test of the GetInternalRowHeight function")]
        [InlineData(0.1, 0f)]
        [InlineData(0.5, 0.75)]
        [InlineData(1, 0.75)]
        [InlineData(10, 9.75)]
        [InlineData(15, 15)]
        [InlineData(409, 408.75)]
        [InlineData(409.5, 409.5)]
        [InlineData(0, 0f)]
        public void GetInternalRowHeightTest(float height, float expectedInternalHeight)
        {
            float internalHeight = DataUtils.GetInternalRowHeight(height);
            Assert.Equal(expectedInternalHeight, internalHeight);
        }

        [Theory(DisplayName = "Test of the failing GetInternalRowHeight function on invalid row heights")]
        [InlineData(-0.1)]
        [InlineData(-10)]
        [InlineData(409.6)]
        [InlineData(10000)]
        public void GetInternalRowHeightFailTest(float height)
        {
            Assert.Throws<FormatException>(() => DataUtils.GetInternalRowHeight(height));
        }

        [Theory(DisplayName = "Test of the GetInternalPaneSplitWidth function")]
        [InlineData(0.1f, 390f)]
        [InlineData(1f, 390f)]
        [InlineData(18.5, 2415f)]
        [InlineData(32f, 3825f)]
        [InlineData(255, 27240f)]
        [InlineData(256, 27345f)]
        [InlineData(1000, 105465f)]
        [InlineData(0, 390f)]
        [InlineData(-1, 390f)]
        [InlineData(-10, 390f)]
        public void GetInternalPaneSplitWidthTest(float width, float expectedSplitWidth)
        {
            float splitWidth = DataUtils.GetInternalPaneSplitWidth(width);
            Assert.Equal(expectedSplitWidth, splitWidth);
        }

        [Theory(DisplayName = "Test of the GetInternalPaneSplitHeight function")]
        [InlineData(0.1f, 302f)]
        [InlineData(0.5f, 310f)]
        [InlineData(1f, 320f)]
        [InlineData(15f, 600f)]
        [InlineData(409.5, 8490f)]
        [InlineData(500, 10300f)]
        [InlineData(0, 300f)]
        [InlineData(-1, 300f)]
        [InlineData(-10, 300f)]
        public void GetInternalPaneSplitHeightTest(float height, float expectedSplitHeight)
        {
            float splitHeight = DataUtils.GetInternalPaneSplitHeight(height);
            Assert.Equal(expectedSplitHeight, splitHeight);
        }

        [Theory(DisplayName = "Test of the GetPaneSplitHeight function")]
        [InlineData(301f, 0.05f)]
        [InlineData(320f, 1f)]
        [InlineData(600f, 15f)]
        [InlineData(310f, 0.5f)]
        [InlineData(8490f, 409.5)]
        [InlineData(10300f, 500)]
        [InlineData(300f, 0)]
        [InlineData(299.9, 0)]
        [InlineData(-10, 0)]
        public void GetPaneSplitHeightTest(float height, float expectedSplitHeight)
        {
            float splitHeight = DataUtils.GetPaneSplitHeight(height);
            Assert.Equal(expectedSplitHeight, splitHeight);
        }

        [Theory(DisplayName = "Test of the GetPaneSplitWidth function")]
        [InlineData(390f, 0f)]
        [InlineData(2415f, 18.5f)]
        [InlineData(1680f, 11.5f)]
        [InlineData(3825f, 31.9286f)]
        [InlineData(27240f, 254.9286f)]
        [InlineData(27345f, 255.9286f)]
        [InlineData(105465f, 999.9286f)]
        public void GetPaneSplitWidthTest(float width, float expectedSplitWidth)
        {
            float splitWidth = DataUtils.GetPaneSplitWidth(width);
            float delta = Math.Abs(splitWidth - expectedSplitWidth);
            Assert.True(delta < 0.001);
        }

        [Theory(DisplayName = "Test of the GetDateFromOA function")]
        [InlineData(1, "01.01.1900 00:00:00")]
        [InlineData(2.5245370370370401, "02.01.1900 12:35:20")]
        [InlineData(58, "27.02.1900 00:00:00")]
        [InlineData(59, "28.02.1900 00:00:00")]
        [InlineData(59.521203703703705, "28.02.1900 12:30:32")]
        [InlineData(61, "01.03.1900 00:00:00")]
        [InlineData(61.339016203703707, "01.03.1900 08:08:11")]
        [InlineData(22056.924363425926, "20.05.1960 22:11:05")]
        [InlineData(44197, "01.01.2021 00:00:00")]
        [InlineData(1450360.47930556, "12.12.5870 11:30:12")]
        public void GetDateFromOATest(double givenValue, string expectedDateString)
        {
            DateTime expectedDate = DateTime.ParseExact(expectedDateString, "dd.MM.yyyy HH:mm:ss", CultureInfo.InvariantCulture);
            DateTime date = DataUtils.GetDateFromOA(givenValue);
            Assert.Equal(expectedDate, date);
        }

        [Fact]
        public void ClusterTestTemp()
        {
            List<Range> ranges = new List<Range>();
            Range r1 = new Range("B4:C7");
            Range r1b = new Range("E5:F10");
            ranges.Add(r1);
            ranges.Add(r1b);
            Range r2 = new Range("C6:E8");
            var result = DataUtils.SubtractRange(ranges, r2);
            int i = 0;
        }

        [Theory(DisplayName = "Test of the MergeRange function")]
        [InlineData("A2:A2", "A2:A2", DataUtils.RangeMergeStrategy.MergeColumns, "A2:A2")]
        [InlineData("A2:A2", "A5:A6", DataUtils.RangeMergeStrategy.MergeColumns, "A2:A2,A5:A6")]
        [InlineData("B2:B3", "B5:B6", DataUtils.RangeMergeStrategy.MergeColumns, "B2:B3,B5:B6")]
        [InlineData("B2:B4", "B5:B6", DataUtils.RangeMergeStrategy.MergeColumns, "B2:B6")]
        [InlineData("B2:C2", "D2:E2", DataUtils.RangeMergeStrategy.MergeColumns, "B2:E2")] //  Strategy leads to same result, since both ranges are contiguous and rectangular
        [InlineData("B5:B6", "B2:B4", DataUtils.RangeMergeStrategy.MergeColumns, "B2:B6")]
        [InlineData("D2:E2", "B2:C2", DataUtils.RangeMergeStrategy.MergeColumns, "B2:E2")] //  Strategy leads to same result, since both ranges are contiguous and rectangular
        [InlineData("B2:C5", "C4:D6", DataUtils.RangeMergeStrategy.MergeColumns, "B2:B5,C2:C6,D4:D6")]
        [InlineData("B2:C5,E2:F2", "C4:D6", DataUtils.RangeMergeStrategy.MergeColumns, "B2:B5,C2:C6,D4:D6,E2:F2")]
        [InlineData("B2:C5,E3:F4", "C4:E6", DataUtils.RangeMergeStrategy.MergeColumns, "B2:B5,C2:C6,D4:D6,E3:E6,F3:F4")]
        [InlineData("A2:A2", "A2:A2", DataUtils.RangeMergeStrategy.MergeRows, "A2:A2")]
        [InlineData("A2:A2", "A5:A6", DataUtils.RangeMergeStrategy.MergeRows, "A2:A2,A5:A6")]
        [InlineData("B2:B3", "B5:B6", DataUtils.RangeMergeStrategy.MergeRows, "B2:B3,B5:B6")]
        [InlineData("B2:B4", "B5:B6", DataUtils.RangeMergeStrategy.MergeRows, "B2:B6")] //  Strategy leads to same result, since both ranges are contiguous and rectangular   
        [InlineData("B2:C2", "D2:E2", DataUtils.RangeMergeStrategy.MergeRows, "B2:E2")]
        [InlineData("B5:B6", "B2:B4", DataUtils.RangeMergeStrategy.MergeRows, "B2:B6")] //  Strategy leads to same result, since both ranges are contiguous and rectangular   
        [InlineData("D2:E2", "B2:C2", DataUtils.RangeMergeStrategy.MergeRows, "B2:E2")]
        [InlineData("B2:C5", "C4:D6", DataUtils.RangeMergeStrategy.MergeRows, "B2:C3,B4:D5,C6:D6")]
        [InlineData("B2:C5,E2:F2", "C4:D6", DataUtils.RangeMergeStrategy.MergeRows, "B2:C3,B4:D5,C6:D6,E2:F2")]
        [InlineData("B2:C5,E3:F4", "C4:E6", DataUtils.RangeMergeStrategy.MergeRows, "B2:C3,B4:F4,B5:E5,C6:E6,E3:F3")]
        [InlineData("A2:A2", "A2:A2", DataUtils.RangeMergeStrategy.NoMerge, "A2:A2")]
        [InlineData("A2:A2", "A5:A6", DataUtils.RangeMergeStrategy.NoMerge, "A2:A2,A5:A6")]
        [InlineData("B2:B3", "B5:B6", DataUtils.RangeMergeStrategy.NoMerge, "B2:B3,B5:B6")]
        [InlineData("B2:B4", "B5:B6", DataUtils.RangeMergeStrategy.NoMerge, "B2:B4,B5:B6")]
        [InlineData("B2:C2", "D2:E2", DataUtils.RangeMergeStrategy.NoMerge, "B2:C2,D2:E2")]
        [InlineData("B5:B6", "B2:B4", DataUtils.RangeMergeStrategy.NoMerge, "B2:B4,B5:B6")]
        [InlineData("D2:E2", "B2:C2", DataUtils.RangeMergeStrategy.NoMerge, "B2:C2,D2:E2")]
        [InlineData("B2:C5", "C4:D6", DataUtils.RangeMergeStrategy.NoMerge, "B2:B3,C2:C3,B4:B5,C4:C5,D4:D5,C6:C6,D6:D6")]
        public void MergeRangeTest(string givenRangesString, string rangeToAddString, DataUtils.RangeMergeStrategy mergeStrategy, string expectedRangesString)
        {
            List<Range> givenRanges = givenRangesString
                    .Split(',', StringSplitOptions.RemoveEmptyEntries)
                    .Select(r => new Range(r))
                    .ToList();

            Range rangeToAdd = new Range(rangeToAddString);
            
            List<Range> expectedRanges = expectedRangesString
                .Split(',', StringSplitOptions.RemoveEmptyEntries)
                .Select(range => new Range(range))
                .ToList();

            IReadOnlyList<Range> resultRanges = DataUtils.MergeRange(givenRanges, rangeToAdd, mergeStrategy);

            Assert.Equal(resultRanges.Count, expectedRanges.Count);
            foreach (Range range in expectedRanges)
            {
                Assert.Contains(resultRanges, r => r.ToString() == range.ToString());
            }
        }


        [Theory(DisplayName = "Test of the SubtractRange function")]
        [InlineData("B5:C6", "A2:B3", DataUtils.RangeMergeStrategy.MergeColumns, "B5:C6")]
        [InlineData("A2:A2", "A2:A2", DataUtils.RangeMergeStrategy.MergeColumns, "")]
        [InlineData("B3:D5", "A2:E6", DataUtils.RangeMergeStrategy.MergeColumns, "")]
        [InlineData("A2:A2", "A5:A6", DataUtils.RangeMergeStrategy.MergeColumns, "A2:A2")]
        [InlineData("B2:B5", "B4:B6", DataUtils.RangeMergeStrategy.MergeColumns, "B2:B3")]
        [InlineData("B4:B7", "B2:B5", DataUtils.RangeMergeStrategy.MergeColumns, "B6:B7")]
        [InlineData("B2:B7", "A3:C4", DataUtils.RangeMergeStrategy.MergeColumns, "B2:B2,B5:B7")] 
        [InlineData("B3:D5", "A4:E4", DataUtils.RangeMergeStrategy.MergeColumns, "B3:D3,B5:D5")]
        [InlineData("B3:D5", "C2:C6", DataUtils.RangeMergeStrategy.MergeColumns, "B3:B5,D3:D5")]
        [InlineData("B3:D5", "A1:B3", DataUtils.RangeMergeStrategy.MergeColumns, "B4:B5,C3:D5")]
        [InlineData("B3:D5", "A5:B6", DataUtils.RangeMergeStrategy.MergeColumns, "B3:B4,C3:D5")]
        [InlineData("B5:C6,E2:F4", "E4:F4", DataUtils.RangeMergeStrategy.MergeColumns, "B5:C6,E2:F3")]
        [InlineData("B3:C8,D3:E5", "C5:D6", DataUtils.RangeMergeStrategy.MergeColumns, "B3:B8,C3:D4,E3:E5,C7:C8")]
        [InlineData("B3:C8,D3:F5,E7:F7", "C5:E7", DataUtils.RangeMergeStrategy.MergeColumns, "B3:B8,C3:E4,C8:C8,F3:F5,F7:F7")]
        [InlineData("B5:C6", "A2:B3", DataUtils.RangeMergeStrategy.MergeRows, "B5:C6")]
        [InlineData("A2:A2", "A2:A2", DataUtils.RangeMergeStrategy.MergeRows, "")]
        [InlineData("B3:D5", "A2:E6", DataUtils.RangeMergeStrategy.MergeRows, "")]
        [InlineData("A2:A2", "A5:A6", DataUtils.RangeMergeStrategy.MergeRows, "A2:A2")]
        [InlineData("B2:B5", "B4:B6", DataUtils.RangeMergeStrategy.MergeRows, "B2:B3")]
        [InlineData("B4:B7", "B2:B5", DataUtils.RangeMergeStrategy.MergeRows, "B6:B7")]
        [InlineData("B2:B7", "A3:C4", DataUtils.RangeMergeStrategy.MergeRows, "B2:B2,B5:B7")]
        [InlineData("B3:D5", "A4:E4", DataUtils.RangeMergeStrategy.MergeRows, "B3:D3,B5:D5")]
        [InlineData("B3:D5", "C2:C6", DataUtils.RangeMergeStrategy.MergeRows, "B3:B5,D3:D5")]
        [InlineData("B3:D5", "A1:B3", DataUtils.RangeMergeStrategy.MergeRows, "C3:D3,B4:D5")]
        [InlineData("B3:D5", "A5:B6", DataUtils.RangeMergeStrategy.MergeRows, "B3:D4,C5:D5")]
        [InlineData("B5:C6,E2:F4", "E4:F4", DataUtils.RangeMergeStrategy.MergeRows, "B5:C6,E2:F3")]
        [InlineData("B3:C8,D3:E5", "C5:D6", DataUtils.RangeMergeStrategy.MergeRows, "B3:E4,E5:E5,B5:B6,B7:C8")]
        [InlineData("B3:C8,D3:F5,E7:F7", "C5:E7", DataUtils.RangeMergeStrategy.MergeRows, "B3:F4,F5:F5,B5:B7,B8:C8,F7:F7")]
        [InlineData("B5:C6", "A2:B3", DataUtils.RangeMergeStrategy.NoMerge, "B5:C6")]
        [InlineData("A2:A2", "A2:A2", DataUtils.RangeMergeStrategy.NoMerge, "")]
        [InlineData("B3:D5", "A2:E6", DataUtils.RangeMergeStrategy.NoMerge, "")]
        [InlineData("A2:A2", "A5:A6", DataUtils.RangeMergeStrategy.NoMerge, "A2:A2")]
        [InlineData("B2:B5", "B4:B6", DataUtils.RangeMergeStrategy.NoMerge, "B2:B3")]
        [InlineData("B4:B7", "B2:B5", DataUtils.RangeMergeStrategy.NoMerge, "B6:B7")]
        [InlineData("B2:B7", "A3:C4", DataUtils.RangeMergeStrategy.NoMerge, "B2:B2,B5:B7")]
        [InlineData("B3:D5", "A4:E4", DataUtils.RangeMergeStrategy.NoMerge, "B3:D3,B5:D5")]
        [InlineData("B3:D5", "C2:C6", DataUtils.RangeMergeStrategy.NoMerge, "B3:B5,D3:D5")]
        [InlineData("B3:D5", "A1:B3", DataUtils.RangeMergeStrategy.NoMerge, "C3:D3,B4:B5,C4:D5")]
        [InlineData("B3:D5", "A5:B6", DataUtils.RangeMergeStrategy.NoMerge, "B3:B4,C3:D4,C5:D5")]
        [InlineData("B5:C6,E2:F4", "E4:F4", DataUtils.RangeMergeStrategy.NoMerge, "B5:C6,E2:F3")]
        [InlineData("B3:C8,D3:E5", "C5:D6", DataUtils.RangeMergeStrategy.NoMerge, "B3:B4,C3:C4,D3:D4,E3:E4,B5:B5,B6:B6,E5:E5,B7:B8,C7:C8")]
        [InlineData("B3:C8,D3:F5,E7:F7", "C5:E7", DataUtils.RangeMergeStrategy.NoMerge, "B3:B4,C3:C4,D3:E4,F3:F4,B5:B5,F5:F5,B6:B6,B7:B7,F7:F7,B8:B8,C8:C8")]
        public void SubtractRangeTest(string givenRangesString, string rangeToRemoveString, DataUtils.RangeMergeStrategy mergeStrategy, string expectedRangesString)
        {
            List<Range> givenRanges = givenRangesString
                    .Split(',', StringSplitOptions.RemoveEmptyEntries)
                    .Select(r => new Range(r))
                    .ToList();

            Range rangeToRemove = new Range(rangeToRemoveString);

            List<Range> expectedRanges = expectedRangesString
                .Split(',', StringSplitOptions.RemoveEmptyEntries)
                .Select(range => new Range(range))
                .ToList();

            IReadOnlyList<Range> resultRanges = DataUtils.SubtractRange(givenRanges, rangeToRemove, mergeStrategy);

            Assert.Equal(resultRanges.Count, expectedRanges.Count);
            foreach (Range range in expectedRanges)
            {
                Assert.Contains(resultRanges, r => r.ToString() == range.ToString());
            }
        }


    }
}
