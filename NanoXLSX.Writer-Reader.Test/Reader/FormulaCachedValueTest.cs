using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using NanoXLSX.Enums;
using NanoXLSX.Extensions;
using Xunit;

namespace NanoXLSX.Test.Writer_Reader.Reader
{
    public class FormulaCachedValueTest
    {
        private static readonly XNamespace SpreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        [Fact(DisplayName = "Test of writing and reading a formula without a cached value")]
        public void FormulaWithoutCachedValueWriteReadTest()
        {
            Workbook workbook = CreateWorkbook();
            workbook.CurrentWorksheet.AddCellFormula("1+1", "A1");

            using (MemoryStream stream = SaveWorkbook(workbook))
            {
                XElement cellElement = ReadCellElement(stream, "A1");
                Assert.Null(cellElement.Attribute("t"));
                Assert.NotNull(cellElement.Element(SpreadsheetNamespace + "f"));
                Assert.Null(cellElement.Element(SpreadsheetNamespace + "v"));

                stream.Position = 0;
                Cell cell = WorkbookReader.Load(stream).CurrentWorksheet.Cells["A1"];
                Assert.Equal(Cell.CellType.Formula, cell.DataType);
                Assert.Equal("1+1", cell.Value);
                Assert.Null(cell.Formula.CachedValue);
                Assert.Equal(Cell.CellType.Default, cell.Formula.CachedValueType);
            }
        }

        [Theory(DisplayName = "Test of writing API-created formula cached value types")]
        [InlineData("0", Cell.CellType.String, "str", "0")]
        [InlineData(2, Cell.CellType.Number, null, "2")]
        [InlineData(true, Cell.CellType.Bool, "b", "1")]
        public void FormulaCachedValueWriteTest(object cachedValue, Cell.CellType expectedCachedValueType, string expectedXmlType, string expectedValue)
        {
            Workbook workbook = CreateWorkbookWithCachedFormula(cachedValue);

            using (MemoryStream stream = SaveWorkbook(workbook))
            {
                XElement cellElement = ReadCellElement(stream, "A1");
                Assert.Equal(expectedXmlType, cellElement.Attribute("t")?.Value);
                Assert.Equal(expectedValue, cellElement.Element(SpreadsheetNamespace + "v")?.Value);
                Assert.Equal(expectedCachedValueType, workbook.CurrentWorksheet.Cells["A1"].Formula.CachedValueType);
            }
        }

        [Fact(DisplayName = "Test of writing special API-created formula cached value types")]
        public void FormulaCachedValueWriteTest2()
        {
            Workbook workbook = CreateWorkbook();
            AddFormulaWithCache(workbook, "A1", new DateTime(2026, 7, 30));
            AddFormulaWithCache(workbook, "B1", TimeSpan.FromHours(12));
            AddFormulaWithCache(workbook, "C1", Errors.FormulaError.DivisionByZero);

            using (MemoryStream stream = SaveWorkbook(workbook))
            {
                AssertCellXml(stream, "A1", null, "46233");
                AssertCellXml(stream, "B1", null, "0.5");
                AssertCellXml(stream, "C1", "e", "#DIV/0!");
            }
        }

        [Theory(DisplayName = "Test of distinguishing numeric and ISO date formula caches")]
        [InlineData("46233.5", null)]
        [InlineData("1,234", "d")]
        [InlineData("123-", "d")]
        [InlineData("2026-07-30T00:00:00Z", "d")]
        public void FormulaCachedDateValueWriteTest(string cachedValue, string expectedXmlType)
        {
            Workbook workbook = CreateWorkbookWithCachedFormula(cachedValue);
            workbook.CurrentWorksheet.Cells["A1"].Formula.CachedValueType = Cell.CellType.Date;

            using (MemoryStream stream = SaveWorkbook(workbook))
            {
                XElement cellElement = ReadCellElement(stream, "A1");
                Assert.Equal(expectedXmlType, cellElement.Attribute("t")?.Value);
                Assert.Equal(cachedValue, cellElement.Element(SpreadsheetNamespace + "v")?.Value);
            }
        }

        [Theory(DisplayName = "Test of reading formula cached value types")]
        [InlineData("A1", Cell.CellType.Number, "2")]
        [InlineData("B1", Cell.CellType.Number, "3")]
        [InlineData("C1", Cell.CellType.String, "0")]
        [InlineData("D1", Cell.CellType.Bool, "1")]
        [InlineData("E1", Cell.CellType.Error, "#DIV/0!")]
        [InlineData("F1", Cell.CellType.Date, "2026-07-30T00:00:00Z")]
        [InlineData("G1", Cell.CellType.String, "shared")]
        [InlineData("H1", Cell.CellType.String, "inline")]
        [InlineData("I1", Cell.CellType.Default, "unsupported")]
        [InlineData("J1", Cell.CellType.Default, null)]
        public void FormulaCachedValueReadTest(string address, Cell.CellType expectedCachedValueType, string expectedCachedValue)
        {
            using (MemoryStream stream = CreateStandardsFormulaWorkbook())
            {
                Cell cell = WorkbookReader.Load(stream).CurrentWorksheet.Cells[address];

                Assert.Equal(Cell.CellType.Formula, cell.DataType);
                Assert.Equal("1+1", cell.Value);
                Assert.Equal(expectedCachedValue, cell.Formula.CachedValue);
                Assert.Equal(expectedCachedValueType, cell.Formula.CachedValueType);
            }
        }

        [Fact(DisplayName = "Test of normalizing formula cached value types after reading and writing")]
        public void FormulaCachedValueWriteReadTest()
        {
            using (MemoryStream source = CreateStandardsFormulaWorkbook())
            {
                Workbook workbook = WorkbookReader.Load(source);
                using (MemoryStream result = SaveWorkbook(workbook))
                {
                    AssertCellXml(result, "A1", null, "2");
                    AssertCellXml(result, "B1", null, "3");
                    AssertCellXml(result, "C1", "str", "0");
                    AssertCellXml(result, "D1", "b", "1");
                    AssertCellXml(result, "E1", "e", "#DIV/0!");
                    AssertCellXml(result, "F1", "d", "2026-07-30T00:00:00Z");
                    AssertCellXml(result, "G1", "str", "shared");
                    AssertCellXml(result, "H1", "str", "inline");
                    AssertCellXml(result, "I1", "str", "unsupported");
                    XElement noCache = ReadCellElement(result, "J1");
                    Assert.Null(noCache.Attribute("t"));
                    Assert.Null(noCache.Element(SpreadsheetNamespace + "v"));
                }
            }
        }

        [Theory(DisplayName = "Test of writing a numeric zero cache for empty defined-name references")]
        [InlineData(null)]
        [InlineData("")]
        public void EmptyDefinedNameReferenceWriteTest(object cachedValue)
        {
            Workbook workbook = CreateWorkbook();
            DefinedName definedName = workbook.AddDefinedNameFormula("emptyRef", "B1");
            workbook.CurrentWorksheet.AddCellReference(definedName, "A1", cachedValue);

            using (MemoryStream stream = SaveWorkbook(workbook))
            {
                AssertCellXml(stream, "A1", null, "0");
                Assert.Equal(Cell.CellType.Number, workbook.CurrentWorksheet.Cells["A1"].Formula.CachedValueType);
            }
        }

        [Fact(DisplayName = "Test of writing a boolean constant defined-name cache")]
        public void BooleanDefinedNameReferenceWriteTest()
        {
            Workbook workbook = CreateWorkbook();
            DefinedName definedName = workbook.AddDefinedNameConstant("boolRef", true);
            workbook.CurrentWorksheet.AddCellReference(definedName, "A1");

            using (MemoryStream stream = SaveWorkbook(workbook))
            {
                AssertCellXml(stream, "A1", "b", "1");
                Assert.Equal(Cell.CellType.Bool, workbook.CurrentWorksheet.Cells["A1"].Formula.CachedValueType);
            }
        }

        [Fact(DisplayName = "Test of propagating a defined-name formula error to the referencing cell")]
        public void DefinedNameFormulaErrorWriteTest()
        {
            Workbook workbook = CreateWorkbook();
            DefinedName definedName = DefinedName.ResolveDefinedName("errorRef", "#REF!", workbook, null, null);
            workbook.AddDefinedName(definedName);
            workbook.CurrentWorksheet.AddCellReference(definedName, "A1");

            using (MemoryStream stream = SaveWorkbook(workbook))
            {
                XElement cellElement = ReadCellElement(stream, "A1");
                Assert.Equal("e", cellElement.Attribute("t")?.Value);
                Assert.Equal("errorRef", cellElement.Element(SpreadsheetNamespace + "f")?.Value);
                Assert.Equal("0", cellElement.Element(SpreadsheetNamespace + "v")?.Value);
            }
        }

        private static Workbook CreateWorkbook()
        {
            Workbook workbook = new Workbook(false);
            workbook.AddWorksheet("sheet1");
            return workbook;
        }

        private static Workbook CreateWorkbookWithCachedFormula(object cachedValue)
        {
            Workbook workbook = CreateWorkbook();
            AddFormulaWithCache(workbook, "A1", cachedValue);
            return workbook;
        }

        private static void AddFormulaWithCache(Workbook workbook, string address, object cachedValue)
        {
            workbook.CurrentWorksheet.AddCellFormula("1+1", address);
            workbook.CurrentWorksheet.Cells[address].Formula = new FormulaData("1+1", cachedValue);
        }

        private static MemoryStream SaveWorkbook(Workbook workbook)
        {
            MemoryStream stream = new MemoryStream();
            workbook.SaveAsStream(stream, true);
            stream.Position = 0;
            return stream;
        }

        private static MemoryStream CreateStandardsFormulaWorkbook()
        {
            Workbook workbook = CreateWorkbook();
            workbook.CurrentWorksheet.AddCell("shared", "A1");
            MemoryStream stream = SaveWorkbook(workbook);
            const string worksheetXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData><row r=\"1\">"
                + "<c r=\"A1\"><f>1+1</f><v>2</v></c>"
                + "<c r=\"B1\" t=\"n\"><f>1+1</f><v>3</v></c>"
                + "<c r=\"C1\" t=\"str\"><f>1+1</f><v>0</v></c>"
                + "<c r=\"D1\" t=\"b\"><f>1+1</f><v>1</v></c>"
                + "<c r=\"E1\" t=\"e\"><f>1+1</f><v>#DIV/0!</v></c>"
                + "<c r=\"F1\" t=\"d\"><f>1+1</f><v>2026-07-30T00:00:00Z</v></c>"
                + "<c r=\"G1\" t=\"s\"><f>1+1</f><v>0</v></c>"
                + "<c r=\"H1\" t=\"inlineStr\"><f>1+1</f><is><t>inline</t></is></c>"
                + "<c r=\"I1\" t=\"unsupported\"><f>1+1</f><v>unsupported</v></c>"
                + "<c r=\"J1\" t=\"str\"><f>1+1</f></c>"
                + "</row></sheetData></worksheet>";
            ReplaceZipEntry(stream, "xl/worksheets/sheet1.xml", worksheetXml);
            return stream;
        }

        private static void ReplaceZipEntry(MemoryStream stream, string path, string content)
        {
            stream.Position = 0;
            using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Update, true))
            {
                archive.GetEntry(path)?.Delete();
                ZipArchiveEntry entry = archive.CreateEntry(path);
                using (StreamWriter writer = new StreamWriter(entry.Open(), new UTF8Encoding(false)))
                {
                    writer.Write(content);
                }
            }
            stream.Position = 0;
        }

        private static XElement ReadCellElement(MemoryStream stream, string address)
        {
            stream.Position = 0;
            using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read, true))
            using (Stream entryStream = archive.GetEntry("xl/worksheets/sheet1.xml").Open())
            {
                XDocument document = XDocument.Load(entryStream);
                return document.Descendants(SpreadsheetNamespace + "c").Single(element => element.Attribute("r")?.Value == address);
            }
        }

        private static void AssertCellXml(MemoryStream stream, string address, string expectedType, string expectedValue)
        {
            XElement cellElement = ReadCellElement(stream, address);
            Assert.Equal(expectedType, cellElement.Attribute("t")?.Value);
            Assert.Equal(expectedValue, cellElement.Element(SpreadsheetNamespace + "v")?.Value);
        }
    }
}
