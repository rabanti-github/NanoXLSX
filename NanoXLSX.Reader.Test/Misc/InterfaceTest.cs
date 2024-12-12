using System.Text;
using NanoXLSX;
using NanoXLSX.Internal.Structures;
using Xunit;

namespace NanoXLSX_Test.Reader.Misc
{
    public class InterfaceTest
    {

        [Theory(DisplayName = "test of the GetFormattedValue implementation of the PlainText class")]
        [InlineData("", "<t></t>")]
        [InlineData(null, "<t></t>")]
        [InlineData(" ", "<t xml:space=\"preserve\"> </t>")]
        [InlineData("test", "<t>test</t>")]
        [InlineData(" test", "<t xml:space=\"preserve\"> test</t>")]
        [InlineData("tEst ", "<t xml:space=\"preserve\">tEst </t>")]
        [InlineData(" Test ", "<t xml:space=\"preserve\"> Test </t>")]
        public void PlainTextAddFormattedValueTest(string givenPlainValue, string expectedFormattedValue)
        {
            PlainText plainText = new PlainText(givenPlainValue);
            StringBuilder sb = new StringBuilder();
            plainText.AddFormattedValue(sb);
            Assert.Equal(expectedFormattedValue, sb.ToString());
        }

        [Theory(DisplayName = "test of the GetFormattedValue implementation of the PlainText class")]
        [InlineData("test", "test", true)]
        [InlineData("", "", true)]
        [InlineData(null, null, true)]
        [InlineData("test", "test2", false)]
        [InlineData(null, "test", false)]
        [InlineData(null, "", false)]
        [InlineData("", null, false)]
        [InlineData("test", null, false)]
        public void PlainTextEqualsTest(string thisValue, string otherValue, bool expectedEquality)
        {
            PlainText plainText1 = new PlainText(thisValue);
            PlainText plainText2 = new PlainText(otherValue);
            Assert.Equal(expectedEquality, plainText1.Equals(plainText2));
        }

        [Theory(DisplayName = "Test of the HashCode implementation of the PlainText class")]
        [InlineData("", false)]
        [InlineData(" ", false)]
        [InlineData("Test", false)]
        [InlineData(null, true)]
        public void PlainTextHashCodeTest(string value, bool expectedZero)
        {
            PlainText plainText1 = new PlainText(value);
            if (expectedZero)
            {
                Assert.Equal(0, plainText1.GetHashCode());
            }
            else
            {
                Assert.NotEqual(0, plainText1.GetHashCode());
            }
        }

        [Fact(DisplayName = "Test of the accurate handling of strings if a PlainText was passed as cell value")]
        public void InvokePlainTextValuetest()
        {
            Workbook workbook = new Workbook("worksheet");
            PlainText plainText = new PlainText("test1");
            workbook.CurrentWorksheet.AddCell(plainText, "A1");
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.Equal(Cell.CellType.STRING, givenWorkbook.CurrentWorksheet.Cells["A1"].DataType);
            Assert.Equal("test1", givenWorkbook.CurrentWorksheet.Cells["A1"].Value.ToString());
        }


    }
}
