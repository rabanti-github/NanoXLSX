using System;
using NanoXLSX.Internal.Readers;
using NanoXLSX.Internal.Structures;
using NanoXLSX.Internal.Writers;
using NanoXLSX.Test.Writer_Reader.Utils;
using Xunit;
using static NanoXLSX.Internal.Enums.Password;

namespace NanoXLSX.Test.Writer_Reader.MiscTest
{
    public class InterfaceTest
    {
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
        public void InvokePlainTextValueTest()
        {
            Workbook workbook = new Workbook("worksheet");
            PlainText plainText = new PlainText("test1");
            workbook.CurrentWorksheet.AddCell(plainText, "A1");
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.Equal(Cell.CellType.STRING, givenWorkbook.CurrentWorksheet.Cells["A1"].DataType);
            Assert.Equal("test1", givenWorkbook.CurrentWorksheet.Cells["A1"].Value.ToString());
        }

        [Theory(DisplayName = "Test of the failing attempt of accessing CopyFrom of LegacyPasswordReader (for code coverage)")]
        [InlineData(LegacyPasswordReader.PasswordType.WORKBOOK_PROTECTION)]
        [InlineData(LegacyPasswordReader.PasswordType.WORKSHEET_PROTECTION)]
        public void FailingLegacyPasswordReaderCopyFromTest(LegacyPasswordReader.PasswordType type)
        {
            LegacyPasswordReader reader = new LegacyPasswordReader(type);
            LegacyPasswordReader reader2 = new LegacyPasswordReader(type);
            Assert.Throws<NotImplementedException>(() => reader.CopyFrom(reader2));
        }

        [Theory(DisplayName = "Test of the failing attempt of accessing SetPassword of LegacyPasswordReader (for code coverage)")]
        [InlineData(LegacyPasswordReader.PasswordType.WORKBOOK_PROTECTION)]
        [InlineData(LegacyPasswordReader.PasswordType.WORKSHEET_PROTECTION)]
        public void FailingLegacyPasswordReaderSetPasswordTest(LegacyPasswordReader.PasswordType type)
        {
            LegacyPasswordReader reader = new LegacyPasswordReader(type);
            Assert.Throws<NotImplementedException>(() => reader.SetPassword("test"));
        }

        [Theory(DisplayName = "Test of the failing attempt of accessing UnsetPassword of LegacyPasswordReader (for code coverage)")]
        [InlineData(LegacyPasswordReader.PasswordType.WORKBOOK_PROTECTION)]
        [InlineData(LegacyPasswordReader.PasswordType.WORKSHEET_PROTECTION)]
        public void FailingLegacyPasswordReaderUnsetPasswordTest(LegacyPasswordReader.PasswordType type)
        {
            LegacyPasswordReader reader = new LegacyPasswordReader(type);
            Assert.Throws<NotImplementedException>(() => reader.UnsetPassword());
        }

        [Theory(DisplayName = "Test of the failing attempt of accessing CopyFrom of LegacyPasswordWriter (for code coverage)")]
        [InlineData(PasswordType.WORKBOOK_PROTECTION)]
        [InlineData(PasswordType.WORKSHEET_PROTECTION)]
        public void FailingLegacyPasswordWriterCopyFromTest(PasswordType type)
        {
            LegacyPasswordWriter writer = new LegacyPasswordWriter(type, "1337");
            LegacyPasswordWriter writer2 = new LegacyPasswordWriter(type, "1337");
            Assert.Throws<NotImplementedException>(() => writer.CopyFrom(writer2));
        }

        [Theory(DisplayName = "Test of the failing attempt of accessing SetPassword of LegacyPasswordWriter (for code coverage)")]
        [InlineData(PasswordType.WORKBOOK_PROTECTION)]
        [InlineData(PasswordType.WORKSHEET_PROTECTION)]
        public void FailingLegacyPasswordWriterSetPasswordTest(PasswordType type)
        {
            LegacyPasswordWriter writer = new LegacyPasswordWriter(type, "1337");
            Assert.Throws<NotImplementedException>(() => writer.SetPassword("test"));
        }

        [Theory(DisplayName = "Test of the failing attempt of accessing UnsetPassword of LegacyPasswordWriter (for code coverage)")]
        [InlineData(PasswordType.WORKBOOK_PROTECTION)]
        [InlineData(PasswordType.WORKSHEET_PROTECTION)]
        public void FailingLegacyPasswordWriterUnsetPasswordTest(PasswordType type)
        {
            LegacyPasswordWriter writer = new LegacyPasswordWriter(type, "1337");
            Assert.Throws<NotImplementedException>(() => writer.UnsetPassword());
        }

        [Theory(DisplayName = "Test of the failing attempt of accessing GetPassword of LegacyPasswordWriter (for code coverage)")]
        [InlineData(PasswordType.WORKBOOK_PROTECTION)]
        [InlineData(PasswordType.WORKSHEET_PROTECTION)]
        public void FailingLegacyPasswordWriterGetPasswordTest(PasswordType type)
        {
            LegacyPasswordWriter writer = new LegacyPasswordWriter(type, "1337");
            Assert.Throws<NotImplementedException>(() => writer.GetPassword());
        }
    }
}
