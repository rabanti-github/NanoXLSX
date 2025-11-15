using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using NanoXLSX.Extensions;
using NanoXLSX.Test.Writer_Reader.Utils;
using Xunit;

namespace NanoXLSX.Test.Writer_Reader.WorkbookReaderTest
{
    public class LoadTest
    {

        [Fact(DisplayName = "Test of the Load function with a file name")]
        public void LoadTest1()
        {
            Dictionary<string, object> data = CreateSampleData();
            string name = CreateWorksheet("test1", data);
            Workbook workbook = WorkbookReader.Load(name);
            Assert.Equal("test1", workbook.Worksheets[0].SheetName);
            foreach (KeyValuePair<string, object> item in data)
            {
                Assert.Equal(item.Value, workbook.Worksheets[0].GetCell(new Address(item.Key)).Value);
            }
            TestUtils.AssertExistingFile(name, true);
        }

        [Fact(DisplayName = "Test of the Load function with a stream")]
        public void LoadTest2()
        {
            Dictionary<string, object> data = CreateSampleData();
            string name = CreateWorksheet("test1", data);
            FileStream fs = new FileStream(name, FileMode.Open);
            Workbook workbook = WorkbookReader.Load(fs);
            Assert.Equal("test1", workbook.Worksheets[0].SheetName);
            foreach (KeyValuePair<string, object> item in data)
            {
                Assert.Equal(item.Value, workbook.Worksheets[0].GetCell(new Address(item.Key)).Value);
            }
            TestUtils.AssertExistingFile(name, true);
        }

        [Fact(DisplayName = "Test of the LoadAsync function with a file name")]
        public async Task LoadAsyncFileTest()
        {
            Dictionary<string, object> data = CreateSampleData();
            string name = CreateWorksheet("test1", data);
            Workbook workbook = await WorkbookReader.LoadAsync(name);
            Assert.Equal("test1", workbook.Worksheets[0].SheetName);
            foreach (KeyValuePair<string, object> item in data)
            {
                Assert.Equal(item.Value, workbook.Worksheets[0].GetCell(new Address(item.Key)).Value);
            }
            TestUtils.AssertExistingFile(name, true);
        }

        [Fact(DisplayName = "Test of the LoadAsync function with a stream")]
        public async Task LoadAsyncTestAsync()
        {
            Dictionary<string, object> data = CreateSampleData();
            string name = CreateWorksheet("test1", data);
            FileStream fs = new FileStream(name, FileMode.Open);
            Workbook workbook = await WorkbookReader.LoadAsync(fs);
            Assert.Equal("test1", workbook.Worksheets[0].SheetName);
            foreach (KeyValuePair<string, object> item in data)
            {
                Assert.Equal(item.Value, workbook.Worksheets[0].GetCell(new Address(item.Key)).Value);
            }
            TestUtils.AssertExistingFile(name, true);
        }

        private static string CreateWorksheet(string worksheetName, Dictionary<string, object> data)
        {
            string name = TestUtils.GetRandomName();
            Workbook workbook = new Workbook(worksheetName);
            foreach (KeyValuePair<string, object> cell in data)
            {
                workbook.CurrentWorksheet.AddCell(cell.Value, cell.Key);
            }
            workbook.SaveAs(name);
            return name;
        }

        private static Dictionary<string, object> CreateSampleData()
        {
            Dictionary<string, object> data = new Dictionary<string, object>();
            data.Add("A1", "test");
            data.Add("A2", 22);
            data.Add("A3", 11.1f);
            return data;
        }

    }
}
