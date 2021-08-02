using NanoXLSX;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;
using static NanoXLSX.Cell;

namespace NanoXLSX_Test.Cells.Types
{
    // Ensure that these tests are executed sequentially, since static repository methods may be called 
    [Collection(nameof(SequentialCollection))]
    public class OtherTypesCellTest
    {
        CellTypeUtils utils;

        public OtherTypesCellTest()
        {
            utils = new CellTypeUtils();
        }


        [Fact(DisplayName = "Unknown value cell test: Test of the cell values, as well as proper modification")]
        public void UnknownClassesCellTest()
        {
            DummyClass obj1 = new DummyClass(1);
            Cell actualCell = new Cell(obj1, Cell.CellType.DEFAULT, utils.CellAddress);
            Assert.Equal(DummyClass.PREFIX + "1", actualCell.Value.ToString());
            Assert.Equal(typeof(DummyClass), actualCell.Value.GetType());
            Assert.Equal(CellType.STRING, actualCell.DataType);
            actualCell.Value = new DummyClass2(2);
            Assert.Equal(DummyClass2.PREFIX + "2", actualCell.Value.ToString());
            Assert.Equal(typeof(DummyClass2), actualCell.Value.GetType()); // should return the new class type
        }

    }

    public class DummyClass
    {
        public int Number { get; set; }
        public DummyClass(int number)
        {
            Number = number;
        }
        public const string PREFIX = "DummyValue = ";
        public override string ToString()
        {
            return PREFIX + Number.ToString();
        }
    }

    public class DummyClass2
    {
        public int Number { get; set; }
        public DummyClass2(int number)
        {
            Number = number;
        }
        public const string PREFIX = "DummyValue2 = ";
        public override string ToString()
        {
            return PREFIX + Number.ToString();
        }
    }
}
