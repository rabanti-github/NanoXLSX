using System;
using System.IO;
using System.Threading.Tasks;
using Xunit;

namespace NanoXLSX.Test.Workbooks
{
    public class SaveTest
    {
        [Fact(DisplayName = "Test of the Save function (file System)")]
        public void SaveTest1()
        {
            string fileName = WorkbookTest.GetRandomName();
            Workbook workbook = new Workbook(fileName, "test");
            FileInfo fi = new FileInfo(fileName);
            Assert.False(fi.Exists);
            workbook.Save();
            WorkbookTest.AssertExistingFile(fileName, true);
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
            string fileName = WorkbookTest.GetRandomName();
            Workbook workbook = new Workbook(fileName, "test");
            FileInfo fi = new FileInfo(fileName);
            Assert.False(fi.Exists);
            await workbook.SaveAsync();
            WorkbookTest.AssertExistingFile(fileName, true);
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
            string fileName = WorkbookTest.GetRandomName();
            Workbook workbook = new Workbook("test");
            FileInfo fi = new FileInfo(fileName);
            Assert.False(fi.Exists);
            workbook.SaveAs(fileName);
            WorkbookTest.AssertExistingFile(fileName, true);
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
            string fileName = WorkbookTest.GetRandomName();
            Workbook workbook = new Workbook("test");
            FileInfo fi = new FileInfo(fileName);
            Assert.False(fi.Exists);
            await workbook.SaveAsAsync(fileName);
            WorkbookTest.AssertExistingFile(fileName, true);
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
            string fileName = WorkbookTest.GetRandomName();
            Workbook workbook = new Workbook("test");
            FileStream fs = new FileStream(fileName, FileMode.Create);
            Assert.Equal(0, fs.Length);
            workbook.SaveAsStream(fs);
            Assert.False(fs.CanWrite);
            WorkbookTest.AssertExistingFile(fileName, true);
        }

        [Fact(DisplayName = "Test of the failing SaveAsStream function with a already closed stream")]
        public void SaveAsStreamFailTest()
        {
            string fileName = WorkbookTest.GetRandomName();
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
            string fileName = WorkbookTest.GetRandomName();
            Workbook workbook = new Workbook("test");
            FileStream fs = new FileStream(fileName, FileMode.Create);
            Assert.Equal(0, fs.Length);
            await workbook.SaveAsStreamAsync(fs);
            Assert.False(fs.CanWrite);
            WorkbookTest.AssertExistingFile(fileName, true);
        }

        [Fact(DisplayName = "Test of the failing SaveAsStreamAsync function with a already closed stream")]
        public async Task SaveAsStreamAsyncFailTest()
        {
            string fileName = WorkbookTest.GetRandomName();
            Workbook workbook = new Workbook("test");
            FileStream fs = new FileStream(fileName, FileMode.Create);
            fs.Write(new byte[] { 0, 0, 0, 0 }, 0, 4);
            fs.Close();
            await Assert.ThrowsAnyAsync<Exception>(() => workbook.SaveAsStreamAsync(fs));
        }

        [Fact(DisplayName = "Test of the failing SaveAsStreamAsync function with a null stream")]
        public async Task SaveAsStreamAsyncFailTest2()
        {
            WorkbookTest.GetRandomName();
            Workbook workbook = new Workbook("test");
            await Assert.ThrowsAnyAsync<Exception>(() => workbook.SaveAsStreamAsync(null));
        }



    }
}
