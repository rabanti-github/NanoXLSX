using NanoXLSX;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;
using static NanoXLSX.Cell;

namespace NanoXLSX_Test.Cells
{
    public class TestUtils
    {
        public static void AssertEquals<T>(T value1, T value2, T inequalValue, Address cellAddress)
        {
            Cell cell1 = new Cell(value1, CellType.DEFAULT, cellAddress);
            Cell cell2 = new Cell(value2, CellType.DEFAULT, cellAddress);
            Cell cell3 = new Cell(inequalValue, CellType.DEFAULT, cellAddress);
            Assert.True(cell1.Equals(cell2));
            Assert.False(cell1.Equals(cell3));
        }
        public static void AssertCellRange(string expectedAddresses, List<Address> addresses)
        {
            string[] addressStrings = SplitValues(expectedAddresses);
            List<Address> expected = new List<Address>();
            foreach (string address in addressStrings)
            {
                expected.Add(new Address(address));
            }
            Assert.Equal(expected.Count, addresses.Count);
            for (int i = 0; i < expected.Count; i++)
            {
                Assert.Equal(expected[i], addresses[i]);
            }
        }

        public static List<string> SplitValuesAsList(string valueString)
        {
            return new List<string>(SplitValues(valueString));
        }

            public static string[] SplitValues(string valueString)
        {
            if (valueString == null || valueString == "")
            {
                return new string[0];
            }
            return valueString.Split(new char[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries);
        }
    }
}
