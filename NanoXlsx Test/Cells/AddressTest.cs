using NanoXLSX;
using NanoXLSX.Shared.Exceptions;
using NanoXLSX.Shared.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;
using static NanoXLSX.Cell;

namespace NanoXLSX_Test.Cells
{
    // Ensure that these tests are executed sequentially, since static repository methods may be called 
    [Collection(nameof(SequentialCollection))]
    public class AddressTest
    {

        #region passingTests

        [Theory(DisplayName = "Constructor call with string as parameter")]
        [InlineData("A1", 0, 0, AddressType.Default)]
        [InlineData("b10", 1, 9, AddressType.Default)]
        [InlineData("$A1", 0, 0, AddressType.FixedColumn)]
        [InlineData("A$1048576", 0, 1048575, AddressType.FixedRow)]
        [InlineData("$xFd$1", 16383, 0, AddressType.FixedRowAndColumn)]
        public void AddressConstructorTest(string address, int expectedColumn, int expectedRow, AddressType expectedType)
        {
            Address actualAddress = new Address(address);
            Assert.Equal(expectedRow, actualAddress.Row);
            Assert.Equal(expectedColumn, actualAddress.Column);
            Assert.Equal(expectedType, actualAddress.Type);
        }

        [Theory(DisplayName = "Constructor call with row and column as parameters")]
        [InlineData(0, 0, "A1")]
        [InlineData(4, 9, "E10")]
        [InlineData(16383, 1048575, "XFD1048576")]
        [InlineData(2, 99, "C100")]
        public void AddressConstructorTest2(int column, int row, string expectedAddress)
        {
            Address actualAddress = new Address(column, row);
            Assert.Equal(expectedAddress, actualAddress.ToString());
            Assert.Equal(Cell.AddressType.Default, actualAddress.Type);
        }

        [Theory(DisplayName = "Constructor call with all parameters")]
        [InlineData(0, 0, AddressType.Default, "A1")]
        [InlineData(4, 9, AddressType.FixedColumn, "$E10")]
        [InlineData(16383, 1048575, AddressType.FixedRow, "XFD$1048576")]
        [InlineData(2, 99, AddressType.FixedRowAndColumn, "$C$100")]
        public void AddressConstructorTest3(int column, int row, AddressType type, string expectedAddress)
        {
            Address actualAddress = new Address(column, row, type);
            Assert.Equal(expectedAddress, actualAddress.ToString());
        }

        [Theory(DisplayName = "Constructor call with string and type as parameters")]
        [InlineData("A1", AddressType.Default, "A1")]
        [InlineData("A1", AddressType.FixedColumn, "$A1")]
        [InlineData("A1", AddressType.FixedRow, "A$1")]
        [InlineData("A1", AddressType.FixedRowAndColumn, "$A$1")]
        [InlineData("$A1", AddressType.Default, "A1")]
        [InlineData("A$1", AddressType.Default, "A1")]
        [InlineData("$A$1", AddressType.Default, "A1")]
        public void AddressConstructorTest4(string address, AddressType type, string expectedAddress)
        {
            Address actualAddress = new Address(address, type);
            Assert.Equal(expectedAddress, actualAddress.ToString());
        }

        [Theory(DisplayName = "Test of Equals() implementation")]
        [InlineData("A1", "A1", true)]
        [InlineData("A1", "A2", false)]
        [InlineData("A1", "B1", false)]
        [InlineData("$A1", "$A1", true)]
        [InlineData("$A1", "A1", false)]
        [InlineData("$A1", "A$1", false)]
        [InlineData("$A1", "$A2", false)]
        [InlineData("$A1", "$B1", false)]
        [InlineData("$A$1", "$A$1", true)]
        [InlineData("$A$1", "A1", false)]
        [InlineData("$A$1", "$A1", false)]
        [InlineData("$A$1", "$A$2", false)]
        [InlineData("$A$1", "$B$1", false)]
        [InlineData("A$1", "A$1", true)]
        [InlineData("A$1", "A1", false)]
        [InlineData("A$1", "$A1", false)]
        [InlineData("A$1", "$A$1", false)]
        [InlineData("A$1", "A$2", false)]
        [InlineData("A$1", "B$1", false)]
        public void AddressEqualsTest(string address1, string address2, bool expectedEquality)
        {
            Address currentAddress = new Address(address1);
            Address otherAddress = new Address(address2);
            bool actualEquality = currentAddress.Equals(otherAddress);
            // Enforcing overload usage
            bool actualEquality2 = currentAddress.Equals((object)otherAddress);
            Assert.Equal(expectedEquality, actualEquality);
            Assert.Equal(expectedEquality, actualEquality2);
            if (expectedEquality)
            {
                Assert.True(currentAddress == otherAddress);
                Assert.False(currentAddress != otherAddress);
            }
            else
            {
                Assert.True(currentAddress != otherAddress);
                Assert.False(currentAddress == otherAddress);
            }
        }

        [Fact(DisplayName = "Test of Equals() implementation returning false on different types")]
        public void AddressEqualsTest2()
        {
            Address currentAddress = new Address("A1");
            string other = "test";
            Assert.False(currentAddress.Equals(other));
        }

            [Theory(DisplayName = "Test of the GetAddress method (string output)")]
        [InlineData(0, 0, AddressType.Default, "A1")]
        [InlineData(4, 9, AddressType.FixedColumn, "$E10")]
        [InlineData(16383, 1048575, AddressType.FixedRow, "XFD$1048576")]
        [InlineData(2, 99, AddressType.FixedRowAndColumn, "$C$100")]
        public void GetAddressTest(int column, int row, AddressType type, string expectedAddress)
        {
            Address actualAddress = new Address(column, row, type);
            Assert.Equal(expectedAddress, actualAddress.GetAddress());
        }

        [Theory(DisplayName = "Test of the GetColumn function")]
        [InlineData(0,0, AddressType.Default, "A")]
        [InlineData(5, 100, AddressType.FixedColumn, "F")]
        [InlineData(26, 100, AddressType.FixedRow, "AA")]
        [InlineData(1, 5, AddressType.FixedRowAndColumn, "B")]
        public void GetColumnTest(int columnNumber, int rowNumber, AddressType type, string expectedColumn)
        {
            Address address = new Address(columnNumber, rowNumber, type);
            Assert.Equal(expectedColumn, address.GetColumn());
        }

        [Theory(DisplayName = "Test of GetHashCode() implementation")]
        [InlineData("A1", "A1", true)]
        [InlineData("A1", "A2", false)]
        [InlineData("A1", "B1", false)]
        [InlineData("$A1", "$A1", true)]
        [InlineData("$A1", "A1", false)]
        [InlineData("$A1", "A$1", false)]
        [InlineData("$A1", "$A2", false)]
        [InlineData("$A1", "$B1", false)]
        [InlineData("$A$1", "$A$1", true)]
        [InlineData("$A$1", "A1", false)]
        [InlineData("$A$1", "$A1", false)]
        [InlineData("$A$1", "$A$2", false)]
        [InlineData("$A$1", "$B$1", false)]
        [InlineData("A$1", "A$1", true)]
        [InlineData("A$1", "A1", false)]
        [InlineData("A$1", "$A1", false)]
        [InlineData("A$1", "$A$1", false)]
        [InlineData("A$1", "A$2", false)]
        [InlineData("A$1", "B$1", false)]
        public void AddressGetHashCodeTest(string address1, string address2, bool expectedEquality)
        {
            Address currentAddress = new Address(address1);
            Address otherAddress = new Address(address2);
            if (expectedEquality)
            {
                Assert.Equal(currentAddress.GetHashCode(), otherAddress.GetHashCode());
            }
            else
            {
                Assert.NotEqual(currentAddress.GetHashCode(), otherAddress.GetHashCode());
            }
        }

        #endregion

        #region failingTest

        // Tests which expects an exception

        [Theory(DisplayName = "Fail on invalid constructor calls with an address string")]
        [InlineData(null, typeof(NanoXLSX.Shared.Exceptions.FormatException))]
        [InlineData("", typeof(NanoXLSX.Shared.Exceptions.FormatException))]
        [InlineData("$", typeof(NanoXLSX.Shared.Exceptions.FormatException))]
        [InlineData("2", typeof(NanoXLSX.Shared.Exceptions.FormatException))]
        [InlineData("$D", typeof(NanoXLSX.Shared.Exceptions.FormatException))]
        [InlineData("$2", typeof(NanoXLSX.Shared.Exceptions.FormatException))]
        [InlineData("Z", typeof(NanoXLSX.Shared.Exceptions.FormatException))]
        [InlineData("A1048577", typeof(RangeException))]
        [InlineData("XFE1", typeof(RangeException))]
        public void AddressConstructorFailTest(string address, Type expectedExceptionType)
        {
            Exception ex;
            if (expectedExceptionType == typeof(NanoXLSX.Shared.Exceptions.FormatException))
            {
                // Common malformed addresses
                ex = Assert.Throws<NanoXLSX.Shared.Exceptions.FormatException>(() => new Address(address));
            }
            else
            {
                // Out of range addresses
                ex = Assert.Throws<RangeException>(() => new Address(address));
            }
            Assert.Equal(expectedExceptionType, ex.GetType());
        }

        [Theory(DisplayName = "Fail on invalid constructor calls with column and row numbers")]
        [InlineData(0, -100)]
        [InlineData(-100, 0)]
        [InlineData(-1, -1)]
        [InlineData(16384, 0)]
        [InlineData(0, 1048576)]
        public void AddressConstructorFailTest2(int column, int row)
        {
            Assert.Throws<RangeException>(() => new Address(column, row, AddressType.Default));
        }

        [Theory(DisplayName = "Test of the CompareTo function")]
        [InlineData("A1", "A1", 0)]
        [InlineData("A10", "A2", 1)]
        [InlineData("B2", "D4", -1)]
        [InlineData("$X$99", "X99", 0)] // $ Should have no influence
        [InlineData("A100", "A$20", 1)] // $ Should have no influence
        [InlineData("$C$2", "$D$4", -1)] // $ Should have no influence
        public void CompareToTest(string address1, string address2, int expectedResult)
        {
            Address address = new Address(address1);
            Address otherAddress = new Address(address2);
            int result = address.CompareTo(otherAddress);
            Assert.Equal(expectedResult, result);
        }

        private static object SequentialCollection()
        {
            throw new NotImplementedException();
        }
        #endregion

    }
}
