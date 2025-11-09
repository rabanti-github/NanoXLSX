using System.IO;
using System.Text;
using NanoXLSX.Internal;
using NanoXLSX.Internal.Readers;
using NanoXLSX.Registry;
using NanoXLSX.Styles;
using NanoXLSX.Test.Writer_Reader.Utils;
using Xunit;
using static NanoXLSX.Styles.NumberFormat;

namespace NanoXLSX.Test.Writer_Reader.ReaderTest
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
                Workbook workbook = new Workbook("test");
                StyleReader styleReader = new StyleReader();
                styleReader.Init(memStream, workbook, new ReaderOptions());
                styleReader.Execute();
                StyleReaderContainer styleReaderContainer = workbook.AuxiliaryData.GetData<StyleReaderContainer>(PlugInUUID.StyleReader, PlugInUUID.StyleEntity);
                NumberFormat numberFormat = styleReaderContainer.GetNumberFormat(formatId);
                Assert.NotNull(numberFormat);
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
                Workbook workbook = new Workbook("test");
                StyleReader styleReader = new StyleReader();
                styleReader.Init(memStream, workbook, new ReaderOptions());
                styleReader.Execute();
                StyleReaderContainer styleReaderContainer = workbook.AuxiliaryData.GetData<StyleReaderContainer>(PlugInUUID.StyleReader, PlugInUUID.StyleEntity);
                NumberFormat numberFormat = styleReaderContainer.GetNumberFormat(formatId);
                Assert.Null(numberFormat);
            }
        }

        [Fact(DisplayName = "Test of reusing dynamically created number formats from styles")]
        public void ImplicitNumberFormatBeingReUsed()
        {
            using (MemoryStream memStream = new MemoryStream(Encoding.UTF8.GetBytes(xml)))
            {
                Workbook workbook = new Workbook("test");
                StyleReader styleReader = new StyleReader();
                styleReader.Init(memStream, workbook, new ReaderOptions());
                styleReader.Execute();
                StyleReaderContainer styleReaderContainer = workbook.AuxiliaryData.GetData<StyleReaderContainer>(PlugInUUID.StyleReader, PlugInUUID.StyleEntity);

                Style zeroStyle = styleReaderContainer.GetStyle(0);
                Style firstStyle = styleReaderContainer.GetStyle(1);

                Assert.Same(zeroStyle.CurrentNumberFormat, firstStyle.CurrentNumberFormat);
            }
        }

        [Fact(DisplayName = "Test of reusing dynamically created number formats from styles containing numFmtId")]
        public void ImplicitNumberFormatBeingReUsed2()
        {
            using (MemoryStream memStream = new MemoryStream(Encoding.UTF8.GetBytes(xml)))
            {
                Workbook workbook = new Workbook("test");
                StyleReader styleReader = new StyleReader();
                styleReader.Init(memStream, workbook, new ReaderOptions());
                styleReader.Execute();
                StyleReaderContainer styleReaderContainer = workbook.AuxiliaryData.GetData<StyleReaderContainer>(PlugInUUID.StyleReader, PlugInUUID.StyleEntity);

                Style zeroStyle = styleReaderContainer.GetStyle(0, out _, out _);
                Style firstStyle = styleReaderContainer.GetStyle(1, out _, out _);

                Assert.Same(zeroStyle.CurrentNumberFormat, firstStyle.CurrentNumberFormat);
            }
        }


        [Fact(DisplayName = "Test of dynamically created number formats from styles containing numFmtId")]
        public void DateTimeImplicitNumberFormatAfter14ZeroNumberFormats()
        {
            using (MemoryStream memStream = new MemoryStream(Encoding.UTF8.GetBytes(xml)))
            {
                Workbook workbook = new Workbook("test");
                StyleReader styleReader = new StyleReader();
                styleReader.Init(memStream, workbook, new ReaderOptions());
                styleReader.Execute();
                StyleReaderContainer styleReaderContainer = workbook.AuxiliaryData.GetData<StyleReaderContainer>(PlugInUUID.StyleReader, PlugInUUID.StyleEntity);

                Assert.Equal(15, styleReaderContainer.StyleCount);

                FormatNumber formatNumber = styleReaderContainer.GetStyle(14, out var isDateStyle, out _).CurrentNumberFormat.Number;

                Assert.True(isDateStyle);
                Assert.Equal(FormatNumber.format_14, formatNumber);
            }
        }
    }
}
