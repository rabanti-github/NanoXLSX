using System.IO;
using System.Text;
using NanoXLSX.Internal.Readers;
using NanoXLSX.Styles;
using Xunit;

namespace NanoXLSX_Test.Reader
{
    [Collection(nameof(SequentialCollection))]
    public class StyleReaderTest
    {
        private readonly string xml;

        public StyleReaderTest()
        {
            xml = "<styleSheet>" +
                  " <numFmts count=\"1\">" +
                  "   <numFmt numFmtId=\"169\" formatCode=\"Does not matter\"/>" +
                  " </numFmts>" +
                  " <fonts count=\"1\">" +
                  "   <font>" +
                  "     <sz val=\"9\"/>" +
                  "     <color rgb=\"FF000000\"/>" +
                  "     <name val=\"Arial\"/>" +
                  "     <family val=\"2\"/>" +
                  "     <charset val=\"238\"/>" +
                  "   </font>" +
                  " </fonts>" +
                  " <fills count=\"1\">" +
                  "   <fill>" +
                  "     <patternFill patternType=\"none\"/>" +
                  "   </fill>" +
                  " </fills>" +
                  " <borders count=\"1\">" +
                  "   <border>" +
                  "     <left/>" +
                  "     <right/>" +
                  "     <top/>" +
                  "     <bottom/>" +
                  "     <diagonal/>" +
                  "   </border>" +
                  " </borders>" +
                  " <cellXfs count=\"15\">" +
                  "   <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>" +
                  "   <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>" +
                  "   <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>" +
                  "   <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>" +
                  "   <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>" +
                  "   <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>" +
                  "   <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>" +
                  "   <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>" +
                  "   <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>" +
                  "   <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>" +
                  "   <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>" +
                  "   <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>" +
                  "   <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>" +
                  "   <xf numFmtId=\"20\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>" +
                  "   <xf numFmtId=\"14\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>" +
                  " </cellXfs>" +
                  "</styleSheet>";
        }

        [Theory(DisplayName = "Test of dynamically created number formats from styles containing numFmtId")]
        [InlineData(0)]
        [InlineData(14)]
        [InlineData(20)]
        public void CreatedImplicitNumberFormatExistsWithCorrectId(int formatId)
        {
            using (MemoryStream memStream = new MemoryStream(Encoding.UTF8.GetBytes(xml)))
            {
                StyleReader styleReader = new StyleReader();
                styleReader.Read(memStream);

                NumberFormat numberFormat = styleReader.StyleReaderContainer.GetNumberFormat(formatId);
                Assert.NotSame(null, numberFormat);
            }
        }

        [Theory(DisplayName = "Test of dynamically created number formats from styles containing numFmtId")]
        [InlineData(1)]
        [InlineData(2)]
        [InlineData(3)]
        [InlineData(4)]
        [InlineData(5)]
        [InlineData(6)]
        [InlineData(7)]
        [InlineData(8)]
        [InlineData(9)]
        [InlineData(10)]
        [InlineData(11)]
        [InlineData(12)]
        [InlineData(13)]
        public void NumberFormatNotInSourceAreNotPresent(int formatId)
        {
            using (MemoryStream memStream = new MemoryStream(Encoding.UTF8.GetBytes(xml)))
            {
                StyleReader styleReader = new StyleReader();
                styleReader.Read(memStream);

                NumberFormat numberFormat = styleReader.StyleReaderContainer.GetNumberFormat(formatId);
                Assert.Same(null, numberFormat);
            }
        }

        [Fact(DisplayName = "Test of reusing dynamically created number formats from styles containing numFmtId")]
        public void ImplicitNumberFormatBeingReUsed()
        {
            using (MemoryStream memStream = new MemoryStream(Encoding.UTF8.GetBytes(xml)))
            {
                StyleReader styleReader = new StyleReader();
                styleReader.Read(memStream);

                Style zeroStyle = styleReader.StyleReaderContainer.GetStyle(0, out _, out _);
                Style firstStyle = styleReader.StyleReaderContainer.GetStyle(1, out _, out _);

                Assert.Same(zeroStyle.CurrentNumberFormat, firstStyle.CurrentNumberFormat);
            }
        }


        [Fact(DisplayName = "Test of dynamically created number formats from styles containing numFmtId")]
        public void DateTimeImplicitNumberFormatAfter14ZeroNumberFormats()
        {
            using (MemoryStream memStream = new MemoryStream(Encoding.UTF8.GetBytes(xml)))
            {
                StyleReader styleReader = new StyleReader();
                styleReader.Read(memStream);
                Assert.Equal(15, styleReader.StyleReaderContainer.StyleCount);

				NanoXLSX.Shared.Enums.Styles.NumberFormatEnums.FormatNumber formatNumber = styleReader.StyleReaderContainer.GetStyle(14, out var isDateStyle, out _).CurrentNumberFormat.Number;

                Assert.Equal(true, isDateStyle);
                Assert.Equal(NanoXLSX.Shared.Enums.Styles.NumberFormatEnums.FormatNumber.format_14, formatNumber);
            }
        }
    }
}