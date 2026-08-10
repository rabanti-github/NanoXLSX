using System;
using System.IO;
using System.Threading.Tasks;
using NanoXLSX.Enums;
using NanoXLSX.Extensions;
using NanoXLSX.Test.Writer_Reader.Utils;
using Xunit;

namespace NanoXLSX.Test.Writer_Reader.WorkbookTest
{
    public class SaveTest
    {
        [Fact(DisplayName = "Test of the Save function (file System)")]
        public void SaveTest1()
        {
            string fileName = TestUtils.GetRandomName();
            Workbook workbook = new Workbook(fileName, "test");
            FileInfo fi = new FileInfo(fileName);
            Assert.False(fi.Exists);
            workbook.Save();
            TestUtils.AssertExistingFile(fileName, true);
        }

        [Theory(DisplayName = "Test of the failing Save function (file System)")]
        [InlineData(null)]
        [InlineData("?")]
        [InlineData("")]
        public void SaveFailTest(string fileName)
        {
            Workbook workbook = new Workbook(fileName, "test");
            Assert.ThrowsAny<Exception>(() => workbook.Save());
        }

        [Fact(DisplayName = "Test of the SaveAsync function (file system)")]
        public async Task SaveAsyncTest()
        {
            string fileName = TestUtils.GetRandomName();
            Workbook workbook = new Workbook(fileName, "test");
            FileInfo fi = new FileInfo(fileName);
            Assert.False(fi.Exists);
            await workbook.SaveAsync();
            TestUtils.AssertExistingFile(fileName, true);
        }

        [Theory(DisplayName = "Test of the failing SaveAsync function (file System)")]
        [InlineData(null)]
        [InlineData("?")]
        [InlineData("")]
        public async Task SaveAsyncFailTest(string fileName)
        {
            Workbook workbook = new Workbook(fileName, "test");
            await Assert.ThrowsAnyAsync<Exception>(() => workbook.SaveAsync());
        }

        [Fact(DisplayName = "Test of the SaveAs function (file System)")]
        public void SaveAsTest()
        {
            string fileName = TestUtils.GetRandomName();
            Workbook workbook = new Workbook("test");
            FileInfo fi = new FileInfo(fileName);
            Assert.False(fi.Exists);
            workbook.SaveAs(fileName);
            TestUtils.AssertExistingFile(fileName, true);
        }

        [Theory(DisplayName = "Test of the failing SaveAs function (file System)")]
        [InlineData(null)]
        [InlineData("?")]
        [InlineData("")]
        public void SaveAsFailTest(string fileName)
        {
            Workbook workbook = new Workbook("test");
            Assert.ThrowsAny<Exception>(() => workbook.SaveAs(fileName));
        }

        [Fact(DisplayName = "Test of the SaveAsAsync function (file system)")]
        public async Task SaveAsAsyncTest()
        {
            string fileName = TestUtils.GetRandomName();
            Workbook workbook = new Workbook("test");
            FileInfo fi = new FileInfo(fileName);
            Assert.False(fi.Exists);
            await workbook.SaveAsAsync(fileName);
            TestUtils.AssertExistingFile(fileName, true);
        }

        [Theory(DisplayName = "Test of the failing SaveAsAsync function (file System)")]
        [InlineData(null)]
        [InlineData("?")]
        [InlineData("")]
        public async Task SaveAsAsyncFailTest(string fileName)
        {
            Workbook workbook = new Workbook("test");
            await Assert.ThrowsAnyAsync<Exception>(() => workbook.SaveAsAsync(fileName));
        }

        [Fact(DisplayName = "Test of the SaveAsStream function with a closing stream")]
        public void SaveAsStreamTest()
        {
            string fileName = TestUtils.GetRandomName();
            Workbook workbook = new Workbook("test");
            FileStream fs = new FileStream(fileName, FileMode.Create);
            Assert.Equal(0, fs.Length);
            workbook.SaveAsStream(fs);
            Assert.False(fs.CanWrite);
            TestUtils.AssertExistingFile(fileName, true);
        }

        [Fact(DisplayName = "Test of the failing SaveAsStream function with a already closed stream")]
        public void SaveAsStreamFailTest()
        {
            string fileName = TestUtils.GetRandomName();
            Workbook workbook = new Workbook("test");
            FileStream fs = new FileStream(fileName, FileMode.Create);
            fs.Write(new byte[] { 0, 0, 0, 0 }, 0, 4);
            fs.Close();
            Assert.ThrowsAny<Exception>(() => workbook.SaveAsStream(fs));
        }

        [Fact(DisplayName = "Test of the failing SaveAsStream function with a null stream")]
        public void SaveAsStreamFailTest2()
        {
            Workbook workbook = new Workbook("test");
            Assert.ThrowsAny<Exception>(() => workbook.SaveAsStream(null));
        }

        [Fact(DisplayName = "Test of the SaveAsStreamAsync function with a closing stream")]
        public async Task SaveAsStreamAsyncTest()
        {
            string fileName = TestUtils.GetRandomName();
            Workbook workbook = new Workbook("test");
            FileStream fs = new FileStream(fileName, FileMode.Create);
            Assert.Equal(0, fs.Length);
            await workbook.SaveAsStreamAsync(fs);
            Assert.False(fs.CanWrite);
            TestUtils.AssertExistingFile(fileName, true);
        }

        [Fact(DisplayName = "Test of the failing SaveAsStreamAsync function with a already closed stream")]
        public async Task SaveAsStreamAsyncFailTest()
        {
            string fileName = TestUtils.GetRandomName();
            Workbook workbook = new Workbook("test");
            FileStream fs = new FileStream(fileName, FileMode.Create);
            fs.Write(new byte[] { 0, 0, 0, 0 }, 0, 4);
            fs.Close();
            await Assert.ThrowsAnyAsync<Exception>(() => workbook.SaveAsStreamAsync(fs));
        }

        [Fact(DisplayName = "Test of the failing SaveAsStreamAsync function with a null stream")]
        public async Task SaveAsStreamAsyncFailTest2()
        {
            TestUtils.GetRandomName();
            Workbook workbook = new Workbook("test");
            await Assert.ThrowsAnyAsync<Exception>(() => workbook.SaveAsStreamAsync(null));
        }

        [Theory(DisplayName = "Test worksheet round-trip of typed error cells")]
        [InlineData(Errors.FormulaError.Null, "#NULL!")]
        [InlineData(Errors.FormulaError.DivisionByZero, "#DIV/0!")]
        [InlineData(Errors.FormulaError.Value, "#VALUE!")]
        [InlineData(Errors.FormulaError.Reference, "#REF!")]
        [InlineData(Errors.FormulaError.Name, "#NAME?")]
        [InlineData(Errors.FormulaError.Number, "#NUM!")]
        [InlineData(Errors.FormulaError.NotAvailable, "#N/A")]
        [InlineData(Errors.FormulaError.GettingData, "#GETTING_DATA")]
        public void SaveErrorCellTest(Errors.FormulaError error, string expectedValue)
        {
            Workbook workbook = new Workbook("worksheet1");
            workbook.CurrentWorksheet.AddCell(new Cell(error, Cell.CellType.Error, "A1"), "A1");

            using MemoryStream stream = new MemoryStream();
            workbook.SaveAsStream(stream, true);

            TestUtils.AssertZipEntry(stream, "xl/worksheets/sheet1.xml", "t=\"e\"");
            TestUtils.AssertZipEntry(stream, "xl/worksheets/sheet1.xml", "<v>" + expectedValue + "</v>");

            stream.Position = 0;
            Cell loadedCell = WorkbookReader.Load(stream).CurrentWorksheet.Cells["A1"];
            Assert.Equal(Cell.CellType.Error, loadedCell.DataType);
            Assert.Equal(error, loadedCell.Value);
            Assert.Null(loadedCell.Formula);
        }

        [Fact(DisplayName = "Test worksheet serialization compatibility of string-backed error cells")]
        public void SaveStringErrorCellTest()
        {
            Workbook workbook = new Workbook("worksheet1");
            workbook.CurrentWorksheet.AddCell(new Cell("#REF!", Cell.CellType.Error, "A1"), "A1");
            workbook.CurrentWorksheet.AddCell(new Cell("unsupported", Cell.CellType.Error, "B1"), "B1");

            using MemoryStream stream = new MemoryStream();
            workbook.SaveAsStream(stream, true);

            TestUtils.AssertZipEntry(stream, "xl/worksheets/sheet1.xml", "r=\"A1\"");
            TestUtils.AssertZipEntry(stream, "xl/worksheets/sheet1.xml", "t=\"e\"");
            TestUtils.AssertZipEntry(stream, "xl/worksheets/sheet1.xml", "<v>#REF!</v>");
            TestUtils.AssertZipEntry(stream, "xl/worksheets/sheet1.xml", "<v>#NAME?</v>");

            stream.Position = 0;
            Workbook loadedWorkbook = WorkbookReader.Load(stream);
            Assert.Equal(Errors.FormulaError.Reference, loadedWorkbook.CurrentWorksheet.Cells["A1"].Value);
            Assert.Equal(Errors.FormulaError.Name, loadedWorkbook.CurrentWorksheet.Cells["B1"].Value);
            Assert.Equal(0, loadedWorkbook.Features.FormulaCount);
        }

        // TODO consider move this test to another test class (currently for test coverage)
        [Fact(DisplayName = "Test worksheet serialization of string-backed formula time caches")]
        public void SaveFormulaTimeCacheTest()
        {
            Workbook workbook = new Workbook("worksheet1");
            workbook.CurrentWorksheet.AddCellFormula("B2", "B1");
            workbook.CurrentWorksheet.Cells["B1"].Formula.CachedValue = "0.5";
            workbook.CurrentWorksheet.Cells["B1"].Formula.CachedValueType = Cell.CellType.Time;

            using MemoryStream stream = new MemoryStream();
            workbook.SaveAsStream(stream, true);

            TestUtils.AssertZipEntry(stream, "xl/worksheets/sheet1.xml", "r=\"B1\"");
            TestUtils.AssertZipEntry(stream, "xl/worksheets/sheet1.xml", "t=\"normal\"");
            TestUtils.AssertZipEntry(stream, "xl/worksheets/sheet1.xml", ">B2</f>");
            TestUtils.AssertZipEntry(stream, "xl/worksheets/sheet1.xml", "<v>0.5</v>");
        }



    }
}
