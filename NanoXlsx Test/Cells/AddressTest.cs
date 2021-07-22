using NanoXLSX;
using NanoXLSX.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;
using static NanoXLSX.Cell;

namespace NanoXLSX_Test.Cells
{
    // Ensure that these tests are executed sequentially, since static repository methods are called 
    [Collection(nameof(SequentialCollection))]
    public class AddressTest
    {

        #region passingTests

        [Theory(DisplayName = "Constructor call with string as parameters")]
        [InlineData("A1", 0, 0, AddressType.Default)]
        [InlineData("b10", 1, 9, AddressType.Default)]
        [InlineData("$A1", 0, 0, AddressType.FixedColumn)]
        [InlineData("A$1048576", 0, 1048575, AddressType.FixedRow)]
        [InlineData("$xFd$1", 16383, 0, AddressType.FixedRowAndColumn)]
        public void AddressConstructorTest(string address, int expectedColumn, int expectedRow, AddressType expectedType)
        {
            Address actaulAddress = new Address(address);
            Assert.Equal(expectedRow, actaulAddress.Row);
            Assert.Equal(expectedColumn, actaulAddress.Column);
            Assert.Equal(expectedType, actaulAddress.Type);
        }

        [Theory(DisplayName = "Constructor call with row and column as parameters")]
        [InlineData(0, 0, "A1")]
        [InlineData(4, 9, "E10")]
        [InlineData(16383, 1048575, "XFD1048576")]
        [InlineData(2, 99, "C100")]
        public void AddressConstructorTest2(int column, int row, string expectedAddress)
        {
            Address actaulAddress = new Address(column, row);
            Assert.Equal(expectedAddress, actaulAddress.ToString());
            Assert.Equal(Cell.AddressType.Default, actaulAddress.Type);
        }

        [Theory(DisplayName = "Constructor call with all parameters")]
        [InlineData(0, 0, AddressType.Default, "A1")]
        [InlineData(4, 9, AddressType.FixedColumn, "$E10")]
        [InlineData(16383, 1048575, AddressType.FixedRow, "XFD$1048576")]
        [InlineData(2, 99, AddressType.FixedRowAndColumn, "$C$100")]
        public void AddressConstructorTest3(int column, int row, AddressType type, string expectedAddress)
        {
            Address actaulAddress = new Address(column, row, type);
            Assert.Equal(expectedAddress, actaulAddress.ToString());
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
            Address actaulAddress = new Address(address, type);
            Assert.Equal(expectedAddress, actaulAddress.ToString());
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
            Assert.Equal(expectedEquality, actualEquality);
        }


        [Theory(DisplayName = "Test of the GetAddress method (string output)")]
        [InlineData(0, 0, AddressType.Default, "A1")]
        [InlineData(4, 9, AddressType.FixedColumn, "$E10")]
        [InlineData(16383, 1048575, AddressType.FixedRow, "XFD$1048576")]
        [InlineData(2, 99, AddressType.FixedRowAndColumn, "$C$100")]
        public void GetAddressTest(int column, int row, AddressType type, string expectedAddress)
        {
            Address actaulAddress = new Address(column, row, type);
            Assert.Equal(expectedAddress, actaulAddress.GetAddress());
        }

        [Theory(DisplayName = "Test of the Equals method (override)")]
        [InlineData("A1", "A1", true)]
        [InlineData("$E10", "$E10", true)]
        [InlineData("XFD$1048576", "XFD$1048576", true)]
        [InlineData("$C$100", "$C$100", true)]
        [InlineData("A1", "$A1", false)]
        [InlineData("A1", "A$1", false)]
        [InlineData("A1", "$A$1", false)]
        [InlineData("$A1", "A$1", false)]
        [InlineData("$A$1", "$A1", false)]
        [InlineData("$A$1", "A$1", false)]
        public void EqualsTest(string addressString1, string addressString2, bool expectedEqual)
        {
            Address address1 = new Address(addressString1);
            Address address2 = new Address(addressString2);
            Assert.Equal(address1.Equals(address2), expectedEqual);
        }

        #endregion

        #region failingTest

        // Tests which expects an exception

        [Theory(DisplayName = "Fail on invalid constructor calls with an address string")]
        [InlineData(null, typeof(NanoXLSX.Exceptions.FormatException))]
        [InlineData("", typeof(NanoXLSX.Exceptions.FormatException))]
        [InlineData("$", typeof(NanoXLSX.Exceptions.FormatException))]
        [InlineData("2", typeof(NanoXLSX.Exceptions.FormatException))]
        [InlineData("$D", typeof(NanoXLSX.Exceptions.FormatException))]
        [InlineData("$2", typeof(NanoXLSX.Exceptions.FormatException))]
        [InlineData("Z", typeof(NanoXLSX.Exceptions.FormatException))]
        [InlineData("A1048577", typeof(RangeException))]
        [InlineData("XFE1", typeof(RangeException))]
        public void AddressConstructorFailTest(string address, Type expectedExceptionType)
        {
            Exception ex;
            if (expectedExceptionType == typeof(NanoXLSX.Exceptions.FormatException))
            {
                // Common malformed addresses
                ex = Assert.Throws<NanoXLSX.Exceptions.FormatException>(() => new Address(address));
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

        private static object SequentialCollection()
        {
            throw new NotImplementedException();
        }



        #endregion

    }
}
