using NanoXLSX.Exceptions;
using Xunit;

namespace NanoXLSX.Test.Cells
{
    public class ColumnTest
    {

        [Theory(DisplayName = "Test of the ColumnAddress property, as well as proper modification")]
        [InlineData("A", "A", "B", "B")]
        [InlineData("a", "A", "b", "B")]
        [InlineData("AAB", "AAB", "A", "A")]
        [InlineData("a", "A", "XFD", "XFD")]
        public void ColumnAddressTest(string initialValue, string expectedValue, string changedValue, string expectedChangedValue)
        {
            Column column = new Column(initialValue);
            Assert.Equal(expectedValue, column.ColumnAddress);
            column.ColumnAddress = changedValue;
            Assert.Equal(expectedChangedValue, column.ColumnAddress);
        }

        [Theory(DisplayName = "Test of the failing ColumnAddress property")]
        [InlineData("")]
        [InlineData(null)]
        [InlineData("4")]
        [InlineData("-")]
        [InlineData(".")]
        [InlineData("$")]
        [InlineData("XFE")]
        public void ColumnAddressTest2(string value)
        {
            Column column = new Column("A");
            RangeException ex = Assert.Throws<RangeException>(() => column.ColumnAddress = value);
        }

        [Fact(DisplayName = "Test of the HasAutoFilter property, as well as the constructor and proper modification")]
        public void HasAutoFilterTest()
        {
            Column column = new Column("A");
            Assert.False(column.HasAutoFilter);
            column.HasAutoFilter = true;
            Assert.True(column.HasAutoFilter);
        }

        [Fact(DisplayName = "Test of the IsHidden property, as well as proper modification")]
        public void IsHiddenTest()
        {
            Column column = new Column("A");
            Assert.False(column.IsHidden);
            column.IsHidden = true;
            Assert.True(column.IsHidden);
        }

        [Theory(DisplayName = "Test of the Number property, as well as the constructor and proper modification")]
        [InlineData(0, 0, 1, 1)]
        [InlineData(999, 999, 5, 5)]
        [InlineData(0, 0, 16383, 16383)]
        public void NumberTest(int initialValue, int expectedValue, int changedValue, int expectedChangedValue)
        {
            Column column = new Column(initialValue);
            Assert.Equal(expectedValue, column.Number);
            column.Number = changedValue;
            Assert.Equal(expectedChangedValue, column.Number);
        }

        [Theory(DisplayName = "Test of the failing Number property")]
        [InlineData(-1)]
        [InlineData(16384)]
        public void NumberTest2(int value)
        {
            Column column = new Column(2);
            RangeException ex = Assert.Throws<RangeException>(() => column.Number = value);
        }

        [Theory(DisplayName = "Test of the Width property, as well as proper modification")]
        [InlineData(15f, 15f)]
        [InlineData(11.1f, 11.1f)]
        [InlineData(0f, 0f)]
        [InlineData(255f, 255f)]
        public void WidthTest(float initialValue, float expectedValue)
        {
            Column column = new Column(0);
            Assert.Equal(Worksheet.DEFAULT_COLUMN_WIDTH, column.Width);
            column.Width = initialValue;
            Assert.Equal(expectedValue, column.Width);
        }

        [Theory(DisplayName = "Test of the failing Width property")]
        [InlineData(-1f)]
        [InlineData(255.1f)]
        public void WidthTest2(float value)
        {
            Column column = new Column(0);
            RangeException ex = Assert.Throws<RangeException>(() => column.Width = value);
        }

    }
}
