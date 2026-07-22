using NanoXLSX.Exceptions;
using Xunit;
using FormatException = NanoXLSX.Exceptions.FormatException;

namespace NanoXLSX.Test.Core.CellTest
{
    public class CellReferenceTest
    {
        [Fact(DisplayName = "Test of SetReference(DefinedName) sets DataType=Reference and Value=name")]
        public void SetReference_DefinedName()
        {
            DefinedName dn = new DefinedName("MyName", "Sheet1!$A$1");
            Cell cell = new Cell();
            cell.SetReference(dn);
            Assert.Equal(Cell.CellType.Reference, cell.DataType);
            Assert.Equal("MyName", cell.Value);
        }

        [Fact(DisplayName = "Test of SetReference(string) sets DataType=Reference and Value=name")]
        public void SetReference_String()
        {
            Cell cell = new Cell();
            cell.SetReference("MyName");
            Assert.Equal(Cell.CellType.Reference, cell.DataType);
            Assert.Equal("MyName", cell.Value);
        }

        [Fact(DisplayName = "Test that SetReference(null DefinedName) throws WorksheetException")]
        public void SetReference_NullDefinedNameThrows()
        {
            Cell cell = new Cell();
            Assert.Throws<WorksheetException>(() => cell.SetReference((DefinedName)null));
        }

        [Theory(DisplayName = "Test that SetReference(string) rejects null or empty")]
        [InlineData(null)]
        [InlineData("")]
        public void SetReference_InvalidStringThrows(string name)
        {
            Cell cell = new Cell();
            Assert.Throws<FormatException>(() => cell.SetReference(name));
        }

        [Fact(DisplayName = "Test that setting Value after SetReference re-resolves DataType (last-write-wins)")]
        public void Value_OverwritesReference()
        {
            Cell cell = new Cell();
            cell.SetReference("MyName");
            Assert.Equal(Cell.CellType.Reference, cell.DataType);
            cell.Value = 42;
            Assert.Equal(Cell.CellType.Number, cell.DataType);
            Assert.Equal(42, cell.Value);
        }

        [Fact(DisplayName = "Test that SetReference after a Value overwrites to Reference (last-write-wins)")]
        public void Reference_OverwritesValue()
        {
            Cell cell = new Cell { Value = 42 };
            Assert.Equal(Cell.CellType.Number, cell.DataType);
            cell.SetReference("MyName");
            Assert.Equal(Cell.CellType.Reference, cell.DataType);
            Assert.Equal("MyName", cell.Value);
        }
    }
}
