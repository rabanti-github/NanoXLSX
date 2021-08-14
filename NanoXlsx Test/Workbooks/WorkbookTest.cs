using NanoXLSX;
using NanoXLSX.Exceptions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;
using FormatException = NanoXLSX.Exceptions.FormatException;

namespace NanoXLSX_Test.Wprkbooks
{
    public class WorkbookTest
    {
        [Fact(DisplayName = "Test of the get function of the Shortener Property")]
        public void ShortenerTest()
        {
            Workbook workbook = new Workbook(false);
            Assert.NotNull(workbook.WS);
            workbook.AddWorksheet("Sheet1");
            workbook.WS.Value("Test");
            workbook.AddWorksheet("Sheet2");
            workbook.WS.Value("Test2");
            Assert.Equal("Test", workbook.Worksheets[0].GetCell(new Address("A1")).Value.ToString());
            Assert.Equal("Test2", workbook.Worksheets[1].GetCell(new Address("A1")).Value.ToString());
        }

        [Fact(DisplayName = "Test of the get function of the CurrentWorksheet property")]
        public void CurrentWorksheetTest()
        {
            Workbook workbook = new Workbook(false);
            Assert.Null(workbook.CurrentWorksheet);
            workbook.AddWorksheet("Test1");
            Assert.NotNull(workbook.CurrentWorksheet);
            Assert.Equal("Test1", workbook.CurrentWorksheet.SheetName);
            workbook.AddWorksheet("Test2");
            Assert.NotNull(workbook.CurrentWorksheet);
            Assert.Equal("Test2", workbook.CurrentWorksheet.SheetName);
        }


        [Fact(DisplayName = "Test of the Filename property")]
        public void FilenameTest()
        {
            string filename = GetRandomName();
            Workbook workbook = new Workbook(filename, "test");
            Assert.Equal(filename, workbook.Filename);
            workbook.Save();
            AssertExistingFile(filename, true);
            filename = GetRandomName();
            workbook.Filename = filename;
            workbook.Save();
            AssertExistingFile(filename, true);
        }

        [Fact(DisplayName = "Test of the get function of the LockStructureIfProtected property")]
        public void LockStructureIfProtectedTest()
        {
            Workbook workbook = new Workbook(false);
            Assert.False(workbook.LockStructureIfProtected);
            workbook.SetWorkbookProtection(true, false, true, "");
            Assert.True(workbook.LockStructureIfProtected);
            workbook.SetWorkbookProtection(false, false, false, "");
            Assert.False(workbook.LockStructureIfProtected);
        }

        [Fact(DisplayName = "Test of the get function of the LockWindowsIfProtected property")]
        public void LockWindowsIfProtectedTest()
        {
            Workbook workbook = new Workbook(false);
            Assert.False(workbook.LockWindowsIfProtected);
            workbook.SetWorkbookProtection(false, true, true, "");
            Assert.True(workbook.LockWindowsIfProtected);
            workbook.SetWorkbookProtection(false, false, false, "");
            Assert.False(workbook.LockWindowsIfProtected);
        }

        [Fact(DisplayName = "Test of the WorkbookMetadata property")]
        public void WorkbookMetadataTest()
        {
            Workbook workbook = new Workbook(false);
            Assert.NotNull(workbook.WorkbookMetadata); // Should be initialized
            workbook.WorkbookMetadata.Title = "Test";
            Assert.Equal("Test", workbook.WorkbookMetadata.Title);
            Metadata newMetaData = new Metadata();
            workbook.WorkbookMetadata = newMetaData;
            Assert.NotEqual("Test", workbook.WorkbookMetadata.Title);
        }

        [Fact(DisplayName = "Test of the get function of the SelectedWorksheet property")]
        public void SelectedWorksheetTest()
        {
            Workbook workbook = new Workbook("test1");
            Assert.Equal(0, workbook.SelectedWorksheet);
            workbook.AddWorksheet("test2");
            workbook.SetSelectedWorksheet(1);
            Assert.Equal(1, workbook.SelectedWorksheet);
        }

        [Fact(DisplayName = "Test of the get function of the UseWorkbookProtection property")]
        public void UseWorkbookProtectionTest()
        {
            Workbook workbook = new Workbook(false);
            Assert.False(workbook.UseWorkbookProtection);
            workbook.SetWorkbookProtection(true, true, true, "");
            Assert.True(workbook.UseWorkbookProtection);
            workbook.SetWorkbookProtection(false, false, false, "");
            Assert.False(workbook.UseWorkbookProtection);
        }

        [Fact(DisplayName = "Test of the get function of the WorkbookProtectionPassword property")]
        public void WorkbookProtectionPasswordTest()
        {
            Workbook workbook = new Workbook(false);
            Assert.Null(workbook.WorkbookProtectionPassword);
            workbook.SetWorkbookProtection(false, true, true, "test");
            Assert.Equal("test", workbook.WorkbookProtectionPassword);
            workbook.SetWorkbookProtection(false, false, false, "");
            Assert.Equal("", workbook.WorkbookProtectionPassword);
            workbook.SetWorkbookProtection(false, false, false, null);
            Assert.Null(workbook.WorkbookProtectionPassword);
        }

        [Fact(DisplayName = "Test of the get function of the Worksheets property")]
        public void WorksheetsTest()
        {
            Workbook workbook = new Workbook(false);
            Assert.Empty(workbook.Worksheets);
            workbook.AddWorksheet("test1");
            workbook.AddWorksheet("test2");
            Assert.Equal(2, workbook.Worksheets.Count);
            workbook.RemoveWorksheet("test2");
            Assert.Single(workbook.Worksheets);
        }


        [Fact(DisplayName = "Test of the Hidden property")]
        public void HiddenTest()
        {
            Workbook workbook = new Workbook(false);
            Assert.False(workbook.Hidden);
            workbook.Hidden = true;
            Assert.True(workbook.Hidden);
            workbook.Hidden = false;
            Assert.False(workbook.Hidden);
        }


        [Fact(DisplayName = "Test of the Workbook default constructor")]
        public void WorkbookConstructorTest()
        {
            Workbook workbook = new Workbook();
            Assert.Empty(workbook.Worksheets);
            Assert.NotNull(workbook.WorkbookMetadata);
            Assert.Null(workbook.CurrentWorksheet);
            Assert.Null(workbook.Filename);
            Assert.NotNull(workbook.WS);
            Assert.Empty(workbook.Worksheets);
        }

        [Theory(DisplayName = "Test of the Workbook constructor with an automatic option to create an initial worksheet")]
        [InlineData(true, "Sheet1")]
        [InlineData(false, null)]
        public void WorkbookConstructorTest2(bool givenValue, string expectedName)
        {
            Workbook workbook = new Workbook(givenValue);
            if (givenValue)
            {
                Assert.NotNull(workbook.CurrentWorksheet);
                Assert.Equal(expectedName, workbook.Worksheets[0].SheetName);
                Assert.Single(workbook.Worksheets);
            }
            else
            {
                Assert.Empty(workbook.Worksheets);
                Assert.Null(workbook.CurrentWorksheet);
            }
            Assert.NotNull(workbook.WorkbookMetadata);
            Assert.Null(workbook.Filename);
            Assert.NotNull(workbook.WS);
        }

        [Theory(DisplayName = "Test of the Workbook constructor with the name of the initially crated worksheet")]
        [InlineData("Sheet1", "Sheet1")]
        [InlineData("?", "_")]
        [InlineData("", "Sheet1")]
        [InlineData(null, "Sheet1")]
        public void WorkbookConstructorTest3(string givenName, string expectedName)
        {
            Workbook workbook = new Workbook(givenName);
            Assert.NotNull(workbook.CurrentWorksheet);
            Assert.Equal(expectedName, workbook.Worksheets[0].SheetName);
            Assert.Single(workbook.Worksheets);
            Assert.NotNull(workbook.WorkbookMetadata);
            Assert.Null(workbook.Filename);
            Assert.NotNull(workbook.WS);
        }

        [Theory(DisplayName = "Test of the Workbook constructor with the file name of the workbook and the name of the initially crated worksheet")]
        [InlineData("f1.xlsx", "Sheet1", "Sheet1")]
        [InlineData("", "?", "_")]
        [InlineData(null, "", "Sheet1")]
        [InlineData("?", null, "Sheet1")] 
        public void WorkbookConstructorTest4(string fileName, string givenSheetName, string expectedSheetName)
        {
            Workbook workbook = new Workbook(fileName, givenSheetName);
            Assert.NotNull(workbook.CurrentWorksheet);
            Assert.Equal(expectedSheetName, workbook.Worksheets[0].SheetName);
            Assert.Single(workbook.Worksheets);
            Assert.NotNull(workbook.WorkbookMetadata);
            Assert.NotNull(workbook.WS);
            Assert.Equal(fileName, workbook.Filename);
        }

        [Theory(DisplayName = "Test of the Workbook constructor with the file name of the workbook, the name of the initially created worksheet and a sanitizing option")]
        [InlineData(false, "f1.xlsx", "Sheet1", "Sheet1", false)]
        [InlineData(false, "", "?", null, true)]
        [InlineData(false, null, "", null, true)]
        [InlineData(false, "?", null, null, true)]
        [InlineData(true, "f1.xlsx", "Sheet1", "Sheet1", false)]
        [InlineData(true, "", "?", "_", false)]
        [InlineData(true, null, "", "Sheet1", false)]
        [InlineData(true, "?", null, "Sheet1", false)]
        public void WorkbookConstructorTest5(bool sanitize, string fileName, string givenSheetName, string expectedSheetName, bool expectException)
        {
            if (expectException)
            {
                Assert.Throws<FormatException>(() => new Workbook(fileName, givenSheetName, sanitize));
            }
            else
            {
                Workbook workbook = new Workbook(fileName, givenSheetName, sanitize);
                Assert.NotNull(workbook.CurrentWorksheet);
                Assert.Equal(expectedSheetName, workbook.Worksheets[0].SheetName);
                Assert.Single(workbook.Worksheets);
                Assert.NotNull(workbook.WorkbookMetadata);
                Assert.NotNull(workbook.WS);
                Assert.Equal(fileName, workbook.Filename);
            }
        }


        [Theory(DisplayName = "Test of the AddWorksheet function with the worksheet name")]
        [InlineData("test")]
        public void AddWorksheetTest(String name)
        {
            Workbook workbook = new Workbook();

        }



        private void AssertExistingFile(string expectedPath, bool deleteAfterAssertion)
        {
            FileInfo fi = new FileInfo(expectedPath);
            Assert.True(fi.Exists);
            if (deleteAfterAssertion)
            {
                try
                {
                    fi.Delete();
                }
                catch(Exception ex)
                {
                    Console.WriteLine("Could not delete " + expectedPath);
                }
            }
        }

        private static string GetRandomName()
        {
            return Path.GetTempFileName();
        }

    }
}
