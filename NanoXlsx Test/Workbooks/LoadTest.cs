using NanoXLSX;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace NanoXLSX_Test.Workbooks
{
    public class LoadTest
    {

        [Fact(DisplayName = "Test of the Load function with a file name")]
        public void LoadTest1()
        {
            Dictionary<string, object> data = CreateSampleData();
            string name = CreateWorksheet("test1", data);
            Workbook workbook = Workbook.Load(name);
            Assert.Equal("test1", workbook.Worksheets[0].SheetName);
            foreach (KeyValuePair<string, object> item in data)
            {
                Assert.Equal(item.Value, workbook.Worksheets[0].GetCell(new Address(item.Key)).Value);
            }
            WorkbookTest.AssertExistingFile(name, true);
        }

        [Fact(DisplayName = "Test of the Load function with a stream")]
        public void LoadTest2()
        {
            Dictionary<string, object> data = CreateSampleData();
            string name = CreateWorksheet("test1", data);
            FileStream fs = new FileStream(name, FileMode.Open);
            Workbook workbook = Workbook.Load(fs);
            Assert.Equal("test1", workbook.Worksheets[0].SheetName);
            foreach (KeyValuePair<string, object> item in data)
            {
                Assert.Equal(item.Value, workbook.Worksheets[0].GetCell(new Address(item.Key)).Value);
            }
            WorkbookTest.AssertExistingFile(name, true);
        }

        private static string CreateWorksheet(string worksheetName, Dictionary<string, object>data)
        {
            string name = WorkbookTest.GetRandomName();
            Workbook workbook = new Workbook(worksheetName);
            foreach(KeyValuePair<string, object> cell in data)
            {
                workbook.CurrentWorksheet.AddCell(cell.Value, cell.Key);
            }
            workbook.SaveAs(name);
            return name;
        }

        [Fact(DisplayName = "Test of the LoadAsync function with a stream")]
        public async Task LoadAsyncTestAsync()
        {
            Dictionary<string, object> data = CreateSampleData();
            string name = CreateWorksheet("test1", data);
            FileStream fs = new FileStream(name, FileMode.Open);
            Workbook workbook = await Workbook.LoadAsync(fs);
            Assert.Equal("test1", workbook.Worksheets[0].SheetName);
            foreach (KeyValuePair<string, object> item in data)
            {
                Assert.Equal(item.Value, workbook.Worksheets[0].GetCell(new Address(item.Key)).Value);
            }
            WorkbookTest.AssertExistingFile(name, true);
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
