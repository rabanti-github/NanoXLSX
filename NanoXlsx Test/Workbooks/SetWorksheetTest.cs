using NanoXLSX;
using NanoXLSX.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace NanoXLSX_Test.Workbooks
{
    public class SetWorksheetTest
    {
        [Fact(DisplayName = "Test of the SetCurrentWorksheet function by index")]
        public void SetCurrentWorksheetTest()
        {
            Workbook workbook = new Workbook();
            Assert.Null(workbook.CurrentWorksheet);
            workbook.AddWorksheet("test1");
            workbook.AddWorksheet("test2");
            workbook.AddWorksheet("test3");
            Assert.Equal("test3", workbook.CurrentWorksheet.SheetName);
            Worksheet worksheet = workbook.SetCurrentWorksheet(1);
            Assert.Equal("test2", workbook.CurrentWorksheet.SheetName);
            Assert.Equal("test2", worksheet.SheetName);
        }

        [Fact(DisplayName = "Test of the SetCurrentWorksheet function by name")]
        public void SetCurrentWorksheetTest2()
        {
            Workbook workbook = new Workbook();
            Assert.Null(workbook.CurrentWorksheet);
            workbook.AddWorksheet("test1");
            workbook.AddWorksheet("test2");
            workbook.AddWorksheet("test3");
            Assert.Equal("test3", workbook.CurrentWorksheet.SheetName);
            Worksheet worksheet = workbook.SetCurrentWorksheet("test2");
            Assert.Equal("test2", workbook.CurrentWorksheet.SheetName);
            Assert.Equal("test2", worksheet.SheetName);
        }

        [Fact(DisplayName = "Test of the SetCurrentWorksheet function by reference")]
        public void SetCurrentWorksheetTest3()
        {
            Workbook workbook = new Workbook();
            Assert.Null(workbook.CurrentWorksheet);
            workbook.AddWorksheet("test1");
            Worksheet worksheet = new Worksheet();
            worksheet.SetSheetName("test2");
            workbook.AddWorksheet(worksheet);
            workbook.AddWorksheet("test3");
            Assert.Equal("test3", workbook.CurrentWorksheet.SheetName);
            workbook.SetCurrentWorksheet(worksheet);
            Assert.Equal("test2", workbook.CurrentWorksheet.SheetName);
            Assert.Equal("test2", workbook.Worksheets[1].SheetName);
        }

        [Fact(DisplayName = "Test of the failing SetCurrentWorksheet function on an invalid name")]
        public void SetCurrentWorksheetFailTest()
        {
            Workbook workbook = new Workbook();
            Assert.Null(workbook.CurrentWorksheet);
            workbook.AddWorksheet("test1");
            string nullString = null;
            Assert.Throws<WorksheetException>(() => workbook.SetCurrentWorksheet(nullString));
            Assert.Throws<WorksheetException>(() => workbook.SetCurrentWorksheet(""));
            Assert.Throws<WorksheetException>(() => workbook.SetCurrentWorksheet("test2"));
        }

        [Fact(DisplayName = "Test of the failing SetCurrentWorksheet function on an invalid index")]
        public void SetCurrentWorksheetFailTest2()
        {
            Workbook workbook = new Workbook();
            Assert.Null(workbook.CurrentWorksheet);
            workbook.AddWorksheet("test1");
            Assert.Throws<RangeException>(() => workbook.SetCurrentWorksheet(-1));
            Assert.Throws<RangeException>(() => workbook.SetCurrentWorksheet(1));
        }

        [Fact(DisplayName = "Test of the failing SetCurrentWorksheet function on an invalid reference")]
        public void SetCurrentWorksheetFailTest3()
        {
            Workbook workbook = new Workbook();
            Assert.Null(workbook.CurrentWorksheet);
            workbook.AddWorksheet("test1");
            Worksheet worksheet = new Worksheet();
            worksheet.SetSheetName("test2");
            Worksheet nullWorksheet = null;
            Assert.Throws<WorksheetException>(() => workbook.SetCurrentWorksheet(nullWorksheet));
            Assert.Throws<WorksheetException>(() => workbook.SetCurrentWorksheet(worksheet));
        }


        [Fact(DisplayName = "Test of the SetSelectedWorksheet function by name")]
        public void SetSelectedWorksheetTest()
        {
            Workbook workbook = new Workbook();
            workbook.AddWorksheet("test1");
            workbook.AddWorksheet("test2");
            workbook.AddWorksheet("test3");
            Assert.Equal(0, workbook.SelectedWorksheet);
            workbook.SetSelectedWorksheet("test2");
            Assert.Equal(1, workbook.SelectedWorksheet);
        }

        [Fact(DisplayName = "Test of the SetSelectedWorksheet function by index")]
        public void SetSelectedWorksheetTest2()
        {
            Workbook workbook = new Workbook();
            workbook.AddWorksheet("test1");
            workbook.AddWorksheet("test2");
            workbook.AddWorksheet("test3");
            Assert.Equal(0, workbook.SelectedWorksheet);
            workbook.SetSelectedWorksheet(1);
            Assert.Equal(1, workbook.SelectedWorksheet);
        }

        [Fact(DisplayName = "Test of the SetSelectedWorksheet function by reference")]
        public void SetSelectedWorksheetTest3()
        {
            Workbook workbook = new Workbook();
            workbook.AddWorksheet("test1");
            Worksheet worksheet = new Worksheet();
            worksheet.SetSheetName("test2");
            workbook.AddWorksheet(worksheet);
            workbook.AddWorksheet("test3");
            Assert.Equal(0, workbook.SelectedWorksheet);
            workbook.SetSelectedWorksheet(worksheet);
            Assert.Equal(1, workbook.SelectedWorksheet);
        }

        [Fact(DisplayName = "Test of the failing SetSelectedWorksheet function on an invalid name")]
        public void SetSelectedWorksheetFailTest()
        {
            Workbook workbook = new Workbook();
            Assert.Equal(0, workbook.SelectedWorksheet);
            workbook.AddWorksheet("test1");
            string nullString = null;
            Assert.Throws<WorksheetException>(() => workbook.SetSelectedWorksheet(nullString));
            Assert.Throws<WorksheetException>(() => workbook.SetSelectedWorksheet(""));
            Assert.Throws<WorksheetException>(() => workbook.SetSelectedWorksheet("test2"));
        }

        [Fact(DisplayName = "Test of the failing SetSelectedWorksheet function on an invalid index")]
        public void SetSelectedWorksheetFailTest2()
        {
            Workbook workbook = new Workbook();
            Assert.Equal(0, workbook.SelectedWorksheet);
            workbook.AddWorksheet("test1");
            Assert.Throws<RangeException>(() => workbook.SetSelectedWorksheet(-1));
            Assert.Throws<RangeException>(() => workbook.SetSelectedWorksheet(1));
        }

        [Fact(DisplayName = "Test of the failing SetSelectedWorksheet function on an invalid reference")]
        public void SetSelectedWorksheetFailTest3()
        {
            Workbook workbook = new Workbook();
            Assert.Equal(0, workbook.SelectedWorksheet);
            workbook.AddWorksheet("test1");
            Worksheet worksheet = new Worksheet();
            worksheet.SetSheetName("test2");
            Worksheet nullWorksheet = null;
            Assert.Throws<WorksheetException>(() => workbook.SetSelectedWorksheet(nullWorksheet));
            Assert.Throws<WorksheetException>(() => workbook.SetSelectedWorksheet(worksheet));
        }

    }
}
