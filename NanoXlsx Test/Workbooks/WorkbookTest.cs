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

namespace NanoXLSX_Test.Workbooks
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

        [Fact(DisplayName = "Test of the UseWorkbookProtection property")]
        public void UseWorkbookProtectionTest()
        {
            Workbook workbook = new Workbook(false);
            Assert.False(workbook.UseWorkbookProtection);
            workbook.UseWorkbookProtection = true;
            Assert.True(workbook.UseWorkbookProtection);
            workbook.UseWorkbookProtection = false;
            Assert.False(workbook.UseWorkbookProtection);
        }


        [Fact(DisplayName = "Test of the get function of the UseWorkbookProtection property using indirect measures")]
        public void UseWorkbookProtectionTest2()
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
        [InlineData("test", null)]
        [InlineData("test", "test2")]
        [InlineData("0", "_")]
        public void AddWorksheetTest(string name1, string name2)
        {
            Workbook workbook = new Workbook();
            Assert.Empty(workbook.Worksheets);
            workbook.AddWorksheet(name1);
            Assert.Single(workbook.Worksheets);
            Assert.Equal(name1, workbook.Worksheets[0].SheetName);
            Assert.Equal(name1, workbook.CurrentWorksheet.SheetName);
            if (name2 != null)
            {
                workbook.AddWorksheet(name2);
                Assert.Equal(2, workbook.Worksheets.Count);
                Assert.Equal(name2, workbook.Worksheets[1].SheetName);
                Assert.Equal(name2, workbook.CurrentWorksheet.SheetName);
            }
        }

        [Theory(DisplayName = "Test of the failing AddWorksheet function with an invalid worksheet name")]
        [InlineData("Sheet1", null)]
        [InlineData("Sheet1", "")]
        [InlineData("Sheet1", "?")]
        [InlineData("Sheet1", "Sheet1")]
        [InlineData("Sheet1", "--------------------------------")]
        public void AddWorksheetFailTest(string initialWorksheetName, string invalidName)
        {
            Workbook workbook = new Workbook();
            workbook.AddWorksheet(initialWorksheetName);
            Assert.ThrowsAny<Exception>(() => workbook.AddWorksheet(invalidName));
        }

        [Theory(DisplayName = "Test of the AddWorksheet function with the worksheet name and a sanitation option")]
        [InlineData("Sheet1", null, false, false, null)]
        [InlineData("test", "test", false, false, null)]
        [InlineData("Sheet1", "", false, false, null)]
        [InlineData("Sheet1", "--------------------------------", false, false, null)]
        [InlineData("Sheet1", "?", false, false, null)]
        [InlineData("Sheet1", "Sheet2", false, true, "Sheet2")]
        [InlineData("Sheet1", null, true, true, "Sheet2")]
        [InlineData("test", "test", true, true, "test1")]
        [InlineData("Sheet1", "", true, true, "Sheet2")]
        [InlineData("Sheet1", "--------------------------------", true, true, "-------------------------------")]
        [InlineData("Sheet1", "?", true, true, "_")]
        public void AddWorksheetTest2(string initialWorksheetName, string name2, bool sanitize, bool expectedValid, string expectedSheetName)
        {
            Workbook workbook = new Workbook();
            Assert.Empty(workbook.Worksheets);
            workbook.AddWorksheet(initialWorksheetName);
            Assert.Single(workbook.Worksheets);
            if (expectedValid)
            {
                workbook.AddWorksheet(name2, sanitize);
                Assert.Equal(2, workbook.Worksheets.Count);
                Assert.Equal(expectedSheetName, workbook.Worksheets[1].SheetName);
                Assert.Equal(expectedSheetName, workbook.CurrentWorksheet.SheetName);
            }
            else
            {
                Assert.ThrowsAny<Exception>(() => workbook.AddWorksheet(name2, sanitize));
            }
        }

        [Fact(DisplayName = "Test of the AddWorksheet function with a Worksheet object")]
        public void AddWorksheetTest3()
        {
            Workbook workbook = new Workbook();
            Assert.Empty(workbook.Worksheets);
            Worksheet worksheet = new Worksheet();
            worksheet.SheetName = "test";
            workbook.AddWorksheet(worksheet);
            Assert.Single(workbook.Worksheets);
            Assert.Equal("test", workbook.Worksheets[0].SheetName);
            Assert.Equal("test", workbook.CurrentWorksheet.SheetName);
        }

        [Fact(DisplayName = "Test of the failing AddWorksheet function with a null object")]
        public void AddWorksheetFailTest3()
        {
            Workbook workbook = new Workbook();
            Worksheet worksheet = null;
            Assert.ThrowsAny<Exception>(() => workbook.AddWorksheet(worksheet));
        }

        [Fact(DisplayName = "Test of the failing AddWorksheet function with a worksheet and an empty name")]
        public void AddWorksheetFailTest3b()
        {
            Workbook workbook = new Workbook();
            Worksheet worksheet = new Worksheet();
            Assert.ThrowsAny<Exception>(() => workbook.AddWorksheet(worksheet));
        }

        [Fact(DisplayName = "Test of the failing AddWorksheet function with a worksheet with an already defined name")]
        public void AddWorksheetFailTest3c()
        {
            Workbook workbook = new Workbook();
            workbook.AddWorksheet("Sheet1");
            Worksheet worksheet = new Worksheet();
            worksheet.SheetName = "Sheet1";
            Assert.ThrowsAny<Exception>(() => workbook.AddWorksheet(worksheet));
        }

        [Theory(DisplayName = "Test of the AddWorksheet function with the worksheet object and a sanitation option")]
        [InlineData("Sheet1", "Sheet1", false, false, null)]
        [InlineData("Sheet1", null, false, false, null)]
        [InlineData("Sheet1", "Sheet1", true, true, "Sheet2")]
        [InlineData("Sheet1", null, true, true, "Sheet2")]
        public void AddWorksheetTest4(string initialWorksheetName, string name2, bool sanitize, bool expectedValid, string expectedSheetName)
        {
            Workbook workbook = new Workbook();
            Assert.Empty(workbook.Worksheets);
            workbook.AddWorksheet(initialWorksheetName);
            Assert.Single(workbook.Worksheets);
            Worksheet worksheet = new Worksheet();
            if (name2 != null)
            {
                worksheet.SheetName = name2;
            }
            if (expectedValid)
            {
                workbook.AddWorksheet(worksheet, sanitize);
                Assert.Equal(2, workbook.Worksheets.Count);
                Assert.Equal(expectedSheetName, workbook.Worksheets[1].SheetName);
                Assert.Equal(expectedSheetName, workbook.CurrentWorksheet.SheetName);
            }
            else
            {
                Assert.ThrowsAny<Exception>(() => workbook.AddWorksheet(worksheet, sanitize));
            }
        }

        [Fact(DisplayName = "Test of the AddWorksheet function for a valid Sheet ID assignment with a name")]
        public void AddWorksheetTest5()
        {
            Workbook workbook = new Workbook();
            workbook.AddWorksheet("test");
            Assert.Equal(1, workbook.Worksheets[0].SheetID);
            workbook.AddWorksheet("test2");
            Assert.Equal(2, workbook.Worksheets[1].SheetID);
            workbook.RemoveWorksheet("test");
            workbook.AddWorksheet("test3");
            Assert.Equal(2, workbook.Worksheets[1].SheetID);
            workbook.RemoveWorksheet("test2");
            workbook.RemoveWorksheet("test3");
            workbook.AddWorksheet("test4");
            Assert.Equal(1, workbook.Worksheets[0].SheetID);
        }

        [Fact(DisplayName = "Test of the AddWorksheet function for a valid Sheet ID assignment with a worksheet object")]
        public void AddWorksheetTest6()
        {
            Workbook workbook = new Workbook();
            Worksheet worksheet1 = new Worksheet();
            worksheet1.SheetName = "test";
            workbook.AddWorksheet(worksheet1, true);
            Assert.Equal(1, workbook.Worksheets[0].SheetID);
            Worksheet worksheet2 = new Worksheet();
            worksheet2.SheetName = "test2";
            workbook.AddWorksheet(worksheet2, true);
            Assert.Equal(2, workbook.Worksheets[1].SheetID);
            workbook.RemoveWorksheet("test");
            Worksheet worksheet3 = new Worksheet();
            worksheet3.SheetName = "test3";
            workbook.AddWorksheet(worksheet3, true);
            Assert.Equal(2, workbook.Worksheets[1].SheetID);
            workbook.RemoveWorksheet("test2");
            workbook.RemoveWorksheet("test3");
            Worksheet worksheet4 = new Worksheet();
            workbook.AddWorksheet(worksheet4, true);
            Assert.Equal(1, workbook.Worksheets[0].SheetID);
        }

        [Theory(DisplayName = "Test of the RemoveWorksheet function by name")]
        [InlineData(2, 0, 1, 1, 0, 0)]
        [InlineData(2, 1, 0, 1, 0, 0)]
        [InlineData(2, 1, 1, 1, 0, 0)]
        [InlineData(2, 0, 0, 1, 0, 0)]
        [InlineData(1, 0, 0, 0, null, 0)]
        [InlineData(5, 2, 2, 2, 4, 3)]
        [InlineData(5, 0, 0, 4, 0, 0)]
        [InlineData(4, 3, 1, 3, 2, 1)]
        [InlineData(4, 3, 3, 3, 2, 2)]
        public void RemoveWorksheetTest(int worksheetCount, int currentWorksheetIndex, int selectedWorksheetIndex, int worksheetToRemoveIndex, int? expectedCurrentWorksheetIndex, int expectedSelectedWorksheetIndex)
        {
            Workbook workbook = new Workbook();
            string current = null;
            string toRemove = null;
            string expected = null;
            for(int i = 0; i < worksheetCount; i++)
            {
                string name = "Sheet" + (i + 1).ToString();
                workbook.AddWorksheet(name);
                if (i == currentWorksheetIndex)
                {
                    current = name;
                }
                if (i == worksheetToRemoveIndex)
                {
                    toRemove = name;
                }
                if (i == expectedCurrentWorksheetIndex)
                {
                    expected = name;
                }
            }
            AssertWorksheetRemoval<string>(workbook, workbook.RemoveWorksheet, worksheetCount, current, selectedWorksheetIndex, toRemove, expected, expectedSelectedWorksheetIndex);
        }

        [Theory(DisplayName = "Test of the RemoveWorksheet function by index")]
        [InlineData(2, 0, 1, 1, 0, 0)]
        [InlineData(2, 1, 0, 1, 0, 0)]
        [InlineData(2, 1, 1, 1, 0, 0)]
        [InlineData(2, 0, 0, 1, 0, 0)]
        [InlineData(1, 0, 0, 0, null, 0)]
        [InlineData(5, 2, 2, 2, 4, 3)]
        [InlineData(5, 0, 0, 4, 0, 0)]
        [InlineData(4, 3, 1, 3, 2, 1)]
        [InlineData(4, 3, 3, 3, 2, 2)]
        public void RemoveWorksheetTest2(int worksheetCount, int currentWorksheetIndex, int selectedWorksheetIndex, int worksheetToRemoveIndex, int? expectedCurrentWorksheetIndex, int expectedSelectedWorksheetIndex)
        {
            Workbook workbook = new Workbook();
            string current = null;
            int toRemove = -1;
            string expected = null;
            for (int i = 0; i < worksheetCount; i++)
            {
                string name = "Sheet" + (i + 1).ToString();
                workbook.AddWorksheet(name);
                if (i == currentWorksheetIndex)
                {
                    current = name;
                }
                if (i == worksheetToRemoveIndex)
                {
                    toRemove = i;
                }
                if (i == expectedCurrentWorksheetIndex)
                {
                    expected = name;
                }
            }
            AssertWorksheetRemoval<int>(workbook, workbook.RemoveWorksheet, worksheetCount, current, selectedWorksheetIndex, toRemove, expected, expectedSelectedWorksheetIndex);
        }

        [Theory(DisplayName = "Test of the failing RemoveWorksheet function on an non-existing name")]
        [InlineData("test", null)]
        [InlineData("test", "")]
        [InlineData("test", "Test")]
        [InlineData("test", "Sheet1")]
        public void RemoveWorksheetFailTest(string existingWorksheet, string absentWorksheet)
        {
            Workbook workbook = new Workbook();
            workbook.AddWorksheet(existingWorksheet);
            Assert.Throws<WorksheetException>(() => workbook.RemoveWorksheet(absentWorksheet));
        }

        [Theory(DisplayName = "Test of the failing RemoveWorksheet function on an non-existing index")]
        [InlineData("test", -1)]
        [InlineData("test", 1)]
        [InlineData("test", 99)]
        public void RemoveWorksheetFailTest2(string existingWorksheet, int absentIndex)
        {
            Workbook workbook = new Workbook();
            workbook.AddWorksheet(existingWorksheet);
            Assert.Throws<WorksheetException>(() => workbook.RemoveWorksheet(absentIndex));
        }

        [Theory(DisplayName = "Test of the SetWorkbookProtection function")]
        [InlineData(false, false, false, null, false, false, false)]
        [InlineData(true, false, false, "", false, false, false)]
        [InlineData(true, true, false, "test", true, false, true)]
        [InlineData(true, false, true, null, false, true, true)]
        [InlineData(true, true, true, " ", true, true, true)]
        [InlineData(false, true, false, "222", true, false, false)]
        [InlineData(false, false, true, "#*$", false, true, false)]
        [InlineData(false, true, true, "_-_", true, true, false)]

        public void SetWorkbookProtectionTest(bool state, bool protectWindows, bool protectStructure, string password, bool expectedLockWindowsState, bool expectedLockStructureState, bool expectedProtectionState)
        {
            Workbook workbook = new Workbook();
            workbook.SetWorkbookProtection(state, protectWindows, protectStructure, password);
            Assert.Equal(expectedLockWindowsState, workbook.LockWindowsIfProtected);
            Assert.Equal(expectedLockStructureState, workbook.LockStructureIfProtected);
            Assert.Equal(expectedProtectionState, workbook.UseWorkbookProtection);
            Assert.Equal(password, workbook.WorkbookProtectionPassword);
        }






        private void AssertWorksheetRemoval<T>(Workbook workbook, Action<T>removalFunction, int worksheetCount, string currentWorksheet, int selectedWorksheetIndex, T worksheetToRemove, string expectedCurrentWorksheet, int expectedSelectedWorksheetIndex)
        {
            workbook.SetCurrentWorksheet(currentWorksheet);
            workbook.SetSelectedWorksheet(selectedWorksheetIndex);
            Assert.Equal(worksheetCount, workbook.Worksheets.Count);
            Assert.Equal(currentWorksheet, workbook.CurrentWorksheet.SheetName);
            removalFunction.Invoke(worksheetToRemove);
            Assert.Equal(worksheetCount - 1, workbook.Worksheets.Count);
            if (expectedCurrentWorksheet == null)
            {
                Assert.Null(workbook.CurrentWorksheet);
            }
            else
            {
                Assert.Equal(expectedCurrentWorksheet, workbook.CurrentWorksheet.SheetName);
            }
            Assert.Equal(expectedSelectedWorksheetIndex, workbook.SelectedWorksheet);
        }



        public static void AssertExistingFile(string expectedPath, bool deleteAfterAssertion)
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

        public static string GetRandomName()
        {
            string path = Path.GetTempFileName();
            FileInfo fi = new FileInfo(path);
            if (fi.Exists)
            {
                fi.Delete();
            }
            return path.Replace(".tmp", ".xlsx");
        }

    }
}
