using NanoXLSX;
using Xunit;

namespace NanoXLSX_Test.Misc
{
    public class XlsxWriterTest
    {
        [Theory(DisplayName = "Test of the 'EscapeXmlChars' method on characters that has to be replaced, when writing a workbook")]
        [InlineData("test", 0x41, "testAtest")] // Not printable
        [InlineData("test", 0x8, "test test")]    // "
        [InlineData("test", 0xC, "test test")]    // "
        [InlineData("test", 0x1F, "test test")]   // "
        [InlineData("test", 0xD800, "test test")] // Above valid UTF range
        [InlineData("test", 0x3C, "test<test")]   // internally saved as &lt;
        [InlineData("test", 0x3E, "test>test")]   // internally saved as &gt;
        [InlineData("test", 0x26, "test&test")]   // internally saved as &amp;
        public void EscapeXmlCharsTest(string givenPrePostFix, int charToEscape, string expectedText)
        {
            string givenText = givenPrePostFix + (char)charToEscape + givenPrePostFix;
            Workbook workbook = new Workbook("worksheet1");
            workbook.CurrentWorksheet.AddCell(givenText, "A1");
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.Equal(expectedText, givenWorkbook.CurrentWorksheet.Cells["A1"].Value);
        }

        [Theory(DisplayName = "Test of the 'EscapeXmlAttributeChars' method on characters that has to be replaced, when writing a workbook")]
        [InlineData("ws", 0x41, "wsAws")] // Not printable
        [InlineData("ws", 0x8, "ws ws")]    // "
        [InlineData("ws", 0xC, "ws ws")]    // "
        [InlineData("ws", 0x1F, "ws ws")]   // "
        [InlineData("ws", 0xD800, "ws ws")] // Above valid UTF range
        [InlineData("ws", 0x22, "ws\"ws")]  // internally saved as &quot;
        [InlineData("ws", 0x3C, "ws<ws")]   // internally saved as &lt;
        [InlineData("ws", 0x3E, "ws>ws")]   // internally saved as &gt;
        [InlineData("ws", 0x26, "ws&ws")]   // internally saved as &amp;
        public void EscapeXmlAttributeCharsTest(string givenPrePostFix, int charToEscape, string expectedText)
        {
            // To test the function, the worksheet name is used, since defined as workbook attribute
            string givenName = givenPrePostFix + (char)charToEscape + givenPrePostFix;
            Workbook workbook = new Workbook(givenName);
            workbook.CurrentWorksheet.AddCell(42, "A1");
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.Equal(expectedText, givenWorkbook.CurrentWorksheet.SheetName);
        }

    }
}
