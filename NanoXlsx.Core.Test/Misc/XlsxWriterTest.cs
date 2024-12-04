using NanoXLSX;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace NanoXLSX_Test.Misc
{
    public class XlsxWriterTest
    {

        [Theory(DisplayName = "Test of common string cell values when writing a workbook")]
        [InlineData(null, null)]
        [InlineData("", "")]
        [InlineData(" ", " ")]
        [InlineData("A", "A")]
        [InlineData("0", "0")]
        [InlineData("test", "test")]
        [InlineData("ÄÖÜËäöüëőÀÉÈàéèç$£°%°@¢~", "ÄÖÜËäöüëőÀÉÈàéèç$£°%°@¢~")]
        [InlineData("月曜日", "月曜日")]
        [InlineData("星期一", "星期一")]
        [InlineData("월요일", "월요일")]
        [InlineData("понеділок", "понеділок")]
        [InlineData("Понедельник", "Понедельник")]
        [InlineData("Δευτέρα", "Δευτέρα")]
        [InlineData("الاثنين", "الاثنين")]
        [InlineData("יוֹם שֵׁנִי", "יוֹם שֵׁנִי")]
        [InlineData("सोमवार", "सोमवार")]
        [InlineData("Thứ hai", "Thứ hai")]
        [InlineData("ሶኑይ", "ሶኑይ")]
        [InlineData("সোমবাৰ", "সোমবাৰ")]
        [InlineData("วันจันทร์", "วันจันทร์")]
        [InlineData("ორშაბათი", "ორშაბათი")]
        [InlineData("پیر", "پیر")]
        [InlineData("सोमबार", "सोमबार")]
        [InlineData("திங்கட்கிழமை", "திங்கட்கிழமை")]
        [InlineData("دوو شەمە", "دوو شەمە")]
        [InlineData("ថ្ងៃច័ន្ទ", "ថ្ងៃច័ន្ទ")]
        [InlineData("دۈشەنبە", "دۈشەنبە")]
        [InlineData("సోమవారం", "సోమవారం")]
        [InlineData("തിങ്കളാഴ്ച", "തിങ്കളാഴ്ച")]
        [InlineData("တနင်္လာနေ့", "တနင်္လာနေ့")]
        [InlineData("සඳුදා", "සඳුදා")]
        [InlineData("ހޯމަ ދުވަސް", "ހޯމަ ދުވަސް")]
        [InlineData("x\uD835\uDC00x", "x𝐀x")] // Surrogates test
        [InlineData("x\x0x", "x x")] // Replacement test
        public void StringValueWriteReadTest(string givenString, string expectedString)
        {
            Workbook workbook = new Workbook("worksheet1");
            workbook.CurrentWorksheet.AddCell(givenString, "A1");
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.Equal(expectedString, givenWorkbook.CurrentWorksheet.Cells["A1"].Value);
        }

        [Theory(DisplayName = "Test of the 'EscapeXmlChars' method on characters that has to be replaced, when writing a workbook")]
        [InlineData("test", 0x41, "testAtest")  ] // Not printable
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
