using NanoXLSX;
using NanoXLSX.Exceptions;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;
using FormatException = NanoXLSX.Exceptions.FormatException;

namespace NanoXLSX_Test.Misc
{
    public class UtilsTest
    {
        [Theory(DisplayName = "Test of the GetOADateTimeString function")]
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
            String oaDate = Utils.GetOADateTimeString(date);
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
            double oaDate = Utils.GetOADateTime(date);
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
            double given = Utils.GetOADateTime(date, true);
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
            Assert.Throws<FormatException>(() => Utils.GetOADateTimeString(date));
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
            Assert.Throws<FormatException>(() => Utils.GetOADateTime(date));
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
            string oaDate = Utils.GetOATimeString(time);
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
            double oaTime = Utils.GetOATime(time);
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
            float internalWidth = Utils.GetInternalColumnWidth(width);
            Assert.Equal(expectedInternalWidth, internalWidth);
        }

        [Theory(DisplayName = "Test of the failing GetInternalColumnWidth function on invalid column widths")]
        [InlineData(-0.1)] 
        [InlineData(-10)] 
        [InlineData(255.01)] 
        [InlineData(10000)] 
        public void GetInternalColumnWidthFailTest(float width)
        {
            Assert.Throws<FormatException>(() => Utils.GetInternalColumnWidth(width));
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
            float internalHeight = Utils.GetInternalRowHeight(height);
            Assert.Equal(expectedInternalHeight, internalHeight);
        }

        [Theory(DisplayName = "Test of the failing GetInternalRowHeight function on invalid row heights")]
        [InlineData(-0.1)]
        [InlineData(-10)]
        [InlineData(409.6)]
        [InlineData(10000)]
        public void GetInternalRowHeightFailTest(float height)
        {
            Assert.Throws<FormatException>(() => Utils.GetInternalRowHeight(height));
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
            float splitWidth = Utils.GetInternalPaneSplitWidth(width);
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
            float splitHeight = Utils.GetInternalPaneSplitHeight(height);
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
            float splitHeight = Utils.GetPaneSplitHeight(height);
            Assert.Equal(expectedSplitHeight, splitHeight);
        }


        [Theory(DisplayName = "Test of the Utils ToUpper function")]
        [InlineData("", "")]
        [InlineData(null, null)]
        [InlineData("123", "123")]
        [InlineData("abc", "ABC")]
        [InlineData("ABC", "ABC")]
        public void ToUpperTest(string givenValue, string expectedValue)
        {
            string value = Utils.ToUpper(givenValue);
            Assert.Equal(expectedValue, value);
        }

        [Theory(DisplayName = "Test of the Utils ToString function")]
        [InlineData(-10, "-10")]
        [InlineData(0, "0")]
        [InlineData(1, "1")]
        [InlineData(100, "100")]
        public void ToStringTest(int givenValue, string expectedValue)
        {
            string value = Utils.ToString(givenValue);
            Assert.Equal(expectedValue, value);
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
            DateTime date = Utils.GetDateFromOA(givenValue);
            Assert.Equal(expectedDate, date);
        }

        [Theory(DisplayName = "Test of the GeneratePasswordHash function")]
        [InlineData("x", "CEBA")]
        [InlineData("Test@1-2,3!", "F767")]
        [InlineData(" ", "CE0A")]
        [InlineData("", "")]
        [InlineData(null, "")]
        public void GeneratePasswordHashTest(string givenVPassword, string expectedHash)
        {
            string hash = Utils.GeneratePasswordHash(givenVPassword);
            Assert.Equal(expectedHash, hash);
        }

    }
}
