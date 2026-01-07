using NanoXLSX.Colors;
using NanoXLSX.Styles;
using NanoXLSX.Test.Writer_Reader.Utils;
using NanoXLSX.Themes;
using Xunit;

namespace NanoXLSX.Test.Writer_Reader.MiscTest
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

        [Theory(DisplayName = "Test of the 'SanitizeXmlValue' method on characters that has to be replaced, when writing a workbook")]
        [InlineData("test", 0x41, "testAtest")] // Not printable
        [InlineData("test", 0x8, "test test")]    // "
        [InlineData("test", 0xC, "test test")]    // "
        [InlineData("test", 0x1F, "test test")]   // "
        [InlineData("test", 0xD800, "test test")] // Above valid UTF range
        [InlineData("test", 0x3C, "test<test")]   // internally saved as &lt;
        [InlineData("test", 0x3E, "test>test")]   // internally saved as &gt;
        [InlineData("test", 0x26, "test&test")]   // internally saved as &amp;
        public void SanitizeXmlValueTest(string givenPrePostFix, int charToEscape, string expectedText)
        {
            string givenText = givenPrePostFix + (char)charToEscape + givenPrePostFix;
            Workbook workbook = new Workbook("worksheet1");
            workbook.CurrentWorksheet.AddCell(givenText, "A1");
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.Equal(expectedText, givenWorkbook.CurrentWorksheet.Cells["A1"].Value);
        }

        [Theory(DisplayName = "Test of the 'SanitizeXmlValue' method on characters that has to be replaced in attributes, when writing a workbook")]
        [InlineData("ws", 0x41, "wsAws")] // Not printable
        [InlineData("ws", 0x8, "ws ws")]    // "
        [InlineData("ws", 0xC, "ws ws")]    // "
        [InlineData("ws", 0x1F, "ws ws")]   // "
        [InlineData("ws", 0xD800, "ws ws")] // Above valid UTF range
        [InlineData("ws", 0x22, "ws\"ws")]  // internally saved as &quot;
        [InlineData("ws", 0x3C, "ws<ws")]   // internally saved as &lt;
        [InlineData("ws", 0x3E, "ws>ws")]   // internally saved as &gt;
        [InlineData("ws", 0x26, "ws&ws")]   // internally saved as &amp;
        public void SanitizeXmlValueAttributeTest(string givenPrePostFix, int charToEscape, string expectedText)
        {
            // To test the function, the worksheet name is used, since defined as workbook attribute
            string givenName = givenPrePostFix + (char)charToEscape + givenPrePostFix;
            Workbook workbook = new Workbook(givenName);
            workbook.CurrentWorksheet.AddCell(42, "A1");
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.Equal(expectedText, givenWorkbook.CurrentWorksheet.SheetName);
        }

        [Fact]
        public void styleTest()
        {
            Workbook wb = new Workbook(@"C:\purge-temp\test1\files\styleTest1.xlsx", "worksheet1");
            Style s1 = new Style();
            SrgbColor c1 = new SrgbColor("FFCC34AF");
            s1.CurrentFill.SetColor(c1, Fill.FillType.FillColor);
            wb.CurrentWorksheet.AddCell("SRGB(FFCC34AF) - Fill", "A1", s1);

            Style s2 = new Style();
            SrgbColor c2 = new SrgbColor("FFAADD00");
            s2.CurrentFill.SetColor(c2, Fill.FillType.PatternColor);
            s2.CurrentFill.PatternFill = Fill.PatternValue.MediumGray;
            wb.CurrentWorksheet.AddCell("SRGB(FFAADD00) - Pattern", "A2", s2);

            Style s3 = new Style();
            IndexedColor c3 = new IndexedColor(IndexedColor.Value.DarkTeal);
            s3.CurrentFill.SetColor(c3, Fill.FillType.FillColor);
            wb.CurrentWorksheet.AddCell("Indexed(DarkTeal) - Fill", "A3", s3);

            Style s4 = new Style();
            IndexedColor c4 = new IndexedColor(IndexedColor.Value.Rose);
            s4.CurrentFill.SetColor(c4, Fill.FillType.PatternColor);
            s4.CurrentFill.PatternFill = Fill.PatternValue.LightGray;
            wb.CurrentWorksheet.AddCell("Indexed(Rose) - Pattern", "A4", s4);

            Style s5 = new Style();
            ThemeColor c5 = new ThemeColor(Theme.ColorSchemeElement.Accent1);
            s5.CurrentFill.SetColor(c5, Fill.FillType.FillColor);
            wb.CurrentWorksheet.AddCell("Theme(Accent1) - Fill", "A5", s5);

            Style s6 = new Style();
            ThemeColor c6 = new ThemeColor(Theme.ColorSchemeElement.Light1);
            s6.CurrentFill.SetColor(c6, Fill.FillType.PatternColor);
            s6.CurrentFill.PatternFill = Fill.PatternValue.LightGray;
            wb.CurrentWorksheet.AddCell("Theme(Light1) - Pattern", "A6", s6);

            Style s7 = new Style();
            AutoColor c7 = new AutoColor();
            s7.CurrentFill.SetColor(c7, Fill.FillType.FillColor);
            wb.CurrentWorksheet.AddCell("Auto - Fill", "A7", s7);

            Style s8 = new Style();
            AutoColor c8 = new AutoColor();
            s8.CurrentFill.SetColor(c8, Fill.FillType.PatternColor);
            s8.CurrentFill.PatternFill = Fill.PatternValue.LightGray;
            wb.CurrentWorksheet.AddCell("Auto(LightGray Pattern) - Pattern", "A8", s8);

            Style s9 = new Style();
            SrgbColor c9 = new SrgbColor("FFCC34AF");
            s9.CurrentFont.ColorValue = Color.CreateRgb(c9);
            wb.CurrentWorksheet.AddCell("SRGB(FFCC34AF) - Font", "B1", s9);

            Style s10 = new Style();
            IndexedColor c10 = new IndexedColor(IndexedColor.Value.Salmon);
            s10.CurrentFont.ColorValue = Color.CreateIndexed(c10);
            wb.CurrentWorksheet.AddCell("Indexed(Salmon) - Font", "B2", s10);

            Style s11 = new Style();
            ThemeColor c11= new ThemeColor(Theme.ColorSchemeElement.Hyperlink);
            s11.CurrentFont.ColorValue = Color.CreateTheme(c11);
            wb.CurrentWorksheet.AddCell("Theme(Hyperlink) - Font", "B3", s11);

            Style s12 = new Style();
            SystemColor c12 = new SystemColor(SystemColor.Value.ScrollBar);
            s12.CurrentFont.ColorValue = Color.CreateSystem(c12);
            wb.CurrentWorksheet.AddCell("System(ScrollBar) - Font", "B4", s12);

            Style s13 = new Style();
            s13.CurrentFont.ColorValue = 63;// Color.CreateAuto();
            wb.CurrentWorksheet.AddCell("Auto - Font", "B5", s13);
            
            // ---

            Style s100 = new Style();
            SrgbColor c100 = new SrgbColor("FFCC34AF");
            s100.CurrentFill.SetColor(c100, Fill.FillType.FillColor);
            s100.CurrentFill.ForegroundColor.Tint = 0.75;
            wb.CurrentWorksheet.AddCell("SRGB(FFCC34AF) - Fill", "C1", s100);

            Style s200 = new Style();
            SrgbColor c200 = new SrgbColor("FFAADD00");
            s200.CurrentFill.SetColor(c200, Fill.FillType.PatternColor);
            s200.CurrentFill.PatternFill = Fill.PatternValue.MediumGray;
            s200.CurrentFill.BackgroundColor.Tint = 0.75;
            wb.CurrentWorksheet.AddCell("SRGB(FFAADD00) - Pattern", "C2", s200);

            Style s300 = new Style();
            IndexedColor c300 = new IndexedColor(IndexedColor.Value.DarkTeal);
            s300.CurrentFill.SetColor(c300, Fill.FillType.FillColor);
            s300.CurrentFill.ForegroundColor.Tint = -0.75;
            wb.CurrentWorksheet.AddCell("Indexed(DarkTeal) - Fill", "C3", s300);

            Style s400 = new Style();
            IndexedColor c400 = new IndexedColor(IndexedColor.Value.Rose);
            s400.CurrentFill.SetColor(c400, Fill.FillType.PatternColor);
            s400.CurrentFill.PatternFill = Fill.PatternValue.LightGray;
            s400.CurrentFill.BackgroundColor.Tint = -0.75;
            wb.CurrentWorksheet.AddCell("Indexed(Rose) - Pattern", "C4", s400);

            Style s500 = new Style();
            ThemeColor c500 = new ThemeColor(Theme.ColorSchemeElement.Accent1);
            s500.CurrentFill.SetColor(c5, Fill.FillType.FillColor);
            s500.CurrentFill.ForegroundColor.Tint = 0.333;
            wb.CurrentWorksheet.AddCell("Theme(Accent1) - Fill", "C5", s500);

            Style s600 = new Style();
            ThemeColor c600 = new ThemeColor(Theme.ColorSchemeElement.Light1);
            s600.CurrentFill.SetColor(c600, Fill.FillType.PatternColor);
            s600.CurrentFill.PatternFill = Fill.PatternValue.LightGray;
            s600.CurrentFill.BackgroundColor.Tint = -0.333;
            wb.CurrentWorksheet.AddCell("Theme(Light1) - Pattern", "C6", s600);

            Style s700 = new Style();
            AutoColor c700 = new AutoColor();
            s700.CurrentFill.SetColor(c700, Fill.FillType.FillColor);
            s700.CurrentFill.ForegroundColor.Tint = 0.5;
            wb.CurrentWorksheet.AddCell("Auto - Fill", "C7", s700);

            Style s800 = new Style();
            AutoColor c800 = new AutoColor();
            s800.CurrentFill.SetColor(c800, Fill.FillType.PatternColor);
            s800.CurrentFill.PatternFill = Fill.PatternValue.LightGray;
            s800.CurrentFill.BackgroundColor.Tint = -0.5;
            wb.CurrentWorksheet.AddCell("Auto(LightGray Pattern) - Pattern", "C8", s800);

            Style s900 = new Style();
            SrgbColor c900 = new SrgbColor("FFCC34AF");
            s900.CurrentFont.ColorValue = Color.CreateRgb(c900);
            s900.CurrentFont.ColorValue.Tint = 0.25;
            wb.CurrentWorksheet.AddCell("SRGB(FFCC34AF) - Font", "D1", s900);

            Style s1000 = new Style();
            IndexedColor c1000 = new IndexedColor(IndexedColor.Value.Salmon);
            s1000.CurrentFont.ColorValue = Color.CreateIndexed(c1000);
            s1000.CurrentFont.ColorValue.Tint = -0.25;
            wb.CurrentWorksheet.AddCell("Indexed(Salmon) - Font", "D2", s1000);

            Style s1100 = new Style();
            ThemeColor c1100 = new ThemeColor(Theme.ColorSchemeElement.Hyperlink);
            s1100.CurrentFont.ColorValue = Color.CreateTheme(c1100);
            s1100.CurrentFont.ColorValue.Tint = 0.95;
            wb.CurrentWorksheet.AddCell("Theme(Hyperlink) - Font", "D3", s1100);

            Style s1200 = new Style();
            SystemColor c1200 = new SystemColor(SystemColor.Value.ScrollBar);
            s1200.CurrentFont.ColorValue = Color.CreateSystem(c1200);
            s1200.CurrentFont.ColorValue.Tint = -0.95;
            wb.CurrentWorksheet.AddCell("System(ScrollBar) - Font", "D4", s1200);

            Style s1300 = new Style();
            s1300.CurrentFont.ColorValue = 63;// Color.CreateAuto();
            s1300.CurrentFont.ColorValue.Tint = 0.86;
            wb.CurrentWorksheet.AddCell("Auto - Font", "D5", s1300);

            wb.Save();
        }

    }
}
