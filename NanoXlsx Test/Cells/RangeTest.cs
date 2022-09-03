using NanoXLSX;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace NanoXLSX_Test.Cells
{
    // Ensure that these tests are executed sequentially, since static repository methods may be called 
    [Collection(nameof(SequentialCollection))]
    public class RangeTest
    {

        [Theory(DisplayName = "Test of the Range constructor with start and end address")]
        [InlineData("A1", "A1", "A1:A1")]
        [InlineData("A1", "C4", "A1:C4")]
        [InlineData("C3", "A1", "A1:C3")]
        [InlineData("$A1", "$A$2", "$A1:$A$2")]
        [InlineData("A$1", "C$4", "A$1:C$4")]
        [InlineData("$C$3", "$A1", "$A1:$C$3")]
        public void ConstructorTest(string startAddress, string endAddress, string expectedRange)
        {
            Address start = new Address(startAddress);
            Address end = new Address(endAddress);
            NanoXLSX.Range range = new NanoXLSX.Range(start, end);
            Assert.Equal(expectedRange, range.ToString());
        }

        [Theory(DisplayName = "Test of the Range constructor with range expression string")]
        [InlineData("A1:A1", "A1:A1")]
        [InlineData("c2:C3", "C2:C3")]
        [InlineData("$A1:$F10", "$A1:$F10")]
        [InlineData("$r$1:$b$2", "$B$2:$R$1")]
        public void ConstructorTest2(string rangeExpression, string expectedRange)
        {
            NanoXLSX.Range range = new NanoXLSX.Range(rangeExpression);
            Assert.Equal(expectedRange, range.ToString());
        }

        [Theory(DisplayName = "Test of the ResolveEnclosedAddressesTest method")]
        [InlineData("A1:A1", "A1")]
        [InlineData("A1:A4", "A1,A2,A3,A4")]
        [InlineData("A1:B3", "A1,A2,A3,B1,B2,B3")]
        [InlineData("B3:A2", "A2,A3,B2,B3")]
        public void ResolveEnclosedAddressesTest(string rangeExpression, string expectedAddresses)
        {
            NanoXLSX.Range range = new NanoXLSX.Range(rangeExpression);
            IReadOnlyList<Address> addresses = range.ResolveEnclosedAddresses();
            TestUtils.AssertCellRange(expectedAddresses, addresses.ToList());
        }

        [Theory(DisplayName = "Test of the Equals method")]
        [InlineData("A1:A1", "A1:A1", true)]
        [InlineData("A1:A4", "A$1:A$4", false)]
        [InlineData("A1:B3", "A1:B4", false)]
        [InlineData("B3:A2", "A2:B3", true)]
        [InlineData("B$3:A2", "A2:B$3", true)]
        public void EqualsTest(string rangeExpression1, string rangeExpression2, bool expectedEquality)
        {
            NanoXLSX.Range range1 = new NanoXLSX.Range(rangeExpression1);
            NanoXLSX.Range range2 = new NanoXLSX.Range(rangeExpression2);
            bool result = range1.Equals(range2);
            Assert.Equal(expectedEquality, result);
        }

        [Fact(DisplayName = "Test of the Equals method returning false on invalid values")]
        public void EqualsTest2()
        {
            NanoXLSX.Range range1 = new NanoXLSX.Range("A1:A7");
            bool result = range1.Equals(null);
            Assert.False(result);
            result = range1.Equals("Wrong type");
            Assert.False(result);
        }
    }
}
