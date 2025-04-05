using System.Collections.Generic;
using System.Linq;
using NanoXLSX.Test.Core.Utils;
using Xunit;

namespace NanoXLSX.Test.Core.CellTest
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
            Range range = new Range(start, end);
            Assert.Equal(expectedRange, range.ToString());
        }

        [Theory(DisplayName = "Test of the Range constructor with range expression string")]
        [InlineData("A1:A1", "A1:A1")]
        [InlineData("c2:C3", "C2:C3")]
        [InlineData("$A1:$F10", "$A1:$F10")]
        [InlineData("$r$1:$b$2", "$B$2:$R$1")]
        public void ConstructorTest2(string rangeExpression, string expectedRange)
        {
            Range range = new Range(rangeExpression);
            Assert.Equal(expectedRange, range.ToString());
        }

        [Theory(DisplayName = "Test of the Range constructor with column and row numbers")]
        [InlineData(0, 0, 0, 0, "A1:A1")]
        [InlineData(0, 0, 1, 1, "A1:B2")]
        [InlineData(1, 1, 0, 0, "A1:B2")]
        public void ConstructorTest3(int startColumn, int startRow, int endColumn, int endRow, string expectedRange)
        {
            Range range = new Range(startColumn, startRow, endColumn, endRow);
            Assert.Equal(expectedRange, range.ToString());
        }

        [Theory(DisplayName = "Test of the ResolveEnclosedAddressesTest method")]
        [InlineData("A1:A1", "A1")]
        [InlineData("A1:A4", "A1,A2,A3,A4")]
        [InlineData("A1:B3", "A1,A2,A3,B1,B2,B3")]
        [InlineData("B3:A2", "A2,A3,B2,B3")]
        public void ResolveEnclosedAddressesTest(string rangeExpression, string expectedAddresses)
        {
            Range range = new Range(rangeExpression);
            IReadOnlyList<Address> addresses = range.ResolveEnclosedAddresses();
            TestUtils.AssertCellRange(expectedAddresses, addresses.ToList());
        }

        [Theory(DisplayName = "Test of the Contains method on addresses")]
        [InlineData("A1:A1", "A1", true)]
        [InlineData("B2:F5", "C3", true)]
        [InlineData("B2:F5", "F5", true)]
        [InlineData("B2:F5", "B2", true)]
        [InlineData("B2:F5", "B5", true)]
        [InlineData("B2:F5", "F2", true)]
        [InlineData("B2:B2", "B1", false)]
        [InlineData("B2:F5", "F6", false)]
        public void ContainsTest(string rangeExpression, string givenAddress, bool expectedResult)
        {
            Range range = new Range(rangeExpression);
            Address address = new Address(givenAddress);
            bool contains = range.Contains(address);
            Assert.Equal(contains, expectedResult);
        }

        [Theory(DisplayName = "Test of the Contains method on ranges")]
        [InlineData("A1:A1", "A1:A1", true)]
        [InlineData("B2:F5", "C3:C3", true)]
        [InlineData("B2:F5", "B2:F5", true)]
        [InlineData("B2:F5", "B2:C3", true)]
        [InlineData("B2:F5", "E4:F5", true)]
        [InlineData("B2:F5", "E2:F3", true)]
        [InlineData("B2:F5", "B4:C5", true)]
        [InlineData("B2:F5", "B1:C3", false)]
        [InlineData("B2:F5", "E2:G3", false)]
        [InlineData("B2:F5", "B5:B6", false)]
        [InlineData("B2:F5", "E4:G6", false)]
        [InlineData("B2:F5", "A1:A2", false)]
        [InlineData("B2:F5", "G1:H2", false)]
        [InlineData("B2:F5", "A6:B8", false)]
        [InlineData("B2:F5", "E6:G7", false)]
        [InlineData("B2:B2", "B1:B1", false)]
        [InlineData("B2:F5", "H3:F6", false)]
        [InlineData("B2:F5", "A1:G8", false)]
        public void ContainsTest2(string rangeExpression, string givenRange, bool expectedResult)
        {
            Range range = new Range(rangeExpression);
            Range range2 = new Range(givenRange);
            bool contains = range.Contains(range2);
            Assert.Equal(contains, expectedResult);
        }

        [Theory(DisplayName = "Test of the Overlaps method")]
        [InlineData("A1:A1", "A1:A1", true)]
        [InlineData("B2:F5", "C3:C3", true)]
        [InlineData("B2:F5", "B2:F5", true)]
        [InlineData("B2:F5", "B2:C3", true)]
        [InlineData("B2:F5", "E4:F5", true)]
        [InlineData("B2:F5", "E2:F3", true)]
        [InlineData("B2:F5", "B4:C5", true)]
        [InlineData("B2:F5", "A1:G8", true)]
        [InlineData("B2:F5", "B1:C3", true)]
        [InlineData("B2:F5", "E2:G3", true)]
        [InlineData("B2:F5", "B5:B6", true)]
        [InlineData("B2:F5", "E4:G6", true)]
        [InlineData("B2:F5", "A1:A2", false)]
        [InlineData("B2:F5", "G1:H2", false)]
        [InlineData("B2:F5", "A6:B8", false)]
        [InlineData("B2:F5", "E6:G7", false)]
        [InlineData("B2:B2", "B1:B1", false)]
        [InlineData("B2:F5", "H3:F6", false)]

        public void OverlapsTest(string rangeExpression, string givenRange, bool expectedResult)
        {
            Range range = new Range(rangeExpression);
            Range range2 = new Range(givenRange);
            bool contains = range.Overlaps(range2);
            Assert.Equal(contains, expectedResult);
        }

        [Theory(DisplayName = "Test of the Equals method")]
        [InlineData("A1:A1", "A1:A1", true)]
        [InlineData("A1:A4", "A$1:A$4", false)]
        [InlineData("A1:B3", "A1:B4", false)]
        [InlineData("B3:A2", "A2:B3", true)]
        [InlineData("B$3:A2", "A2:B$3", true)]
        public void EqualsTest(string rangeExpression1, string rangeExpression2, bool expectedEquality)
        {
            Range range1 = new Range(rangeExpression1);
            Range range2 = new Range(rangeExpression2);
            bool result = range1.Equals(range2);
            Assert.Equal(expectedEquality, result);
            if (expectedEquality)
            {
                Assert.True(range1 == range2);
                Assert.False(range1 != range2);
            }
            else
            {
                Assert.True(range1 != range2);
                Assert.False(range1 == range2);
            }
        }

        [Fact(DisplayName = "Test of the Equals method returning false on invalid values")]
        public void EqualsTest2()
        {
            Range range1 = new Range("A1:A7");
            bool result = range1.Equals(null);
            Assert.False(result);
            result = range1.Equals("Wrong type");
            Assert.False(result);
        }

        [Theory(DisplayName = "Test of the GetHashCode method (equality of two identical objects)")]
        [InlineData("A1:A1", "A1:A1", true)]
        [InlineData("A1:A4", "A$1:A$4", false)]
        [InlineData("A1:B3", "A1:B4", false)]
        [InlineData("B3:A2", "A2:B3", true)]
        [InlineData("B$3:A2", "A2:B$3", true)]
        public void GetHashCodeTest(string rangeExpression1, string rangeExpression2, bool expectedEquality)
        {
            Range range1 = new Range(rangeExpression1);
            Range range2 = new Range(rangeExpression2);
            if (expectedEquality)
            {
                Assert.Equal(range1.GetHashCode(), range2.GetHashCode());
            }
            else
            {
                Assert.NotEqual(range1.GetHashCode(), range2.GetHashCode());
            }
        }
    }
}
