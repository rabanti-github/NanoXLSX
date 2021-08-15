using NanoXLSX;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace NanoXLSX_Test.Worksheets
{
    public class StaticTest
    {
        [Theory(DisplayName = "Test of the SanitizeWorksheetName function")]
        [InlineData("test", 0, null, "test")]
        [InlineData("Sheet2", 1, "Sheet", "Sheet2")]
        [InlineData("", 0, null, "Sheet1")]
        [InlineData(null, 0, null, "Sheet1")]
        [InlineData("a[b", 0, null, "a_b")]
        [InlineData("a]b", 0, null, "a_b")]
        [InlineData("a*b", 0, null, "a_b")]
        [InlineData("a?b", 0, null, "a_b")]
        [InlineData("a/b", 0, null, "a_b")]
        [InlineData("a\\b",0, null, "a_b")]
        [InlineData("--------------------------------", 0, null, "-------------------------------")]
        [InlineData("Sheet10", 20, "Sheet", "Sheet21")]
        [InlineData("*1", 1, "_", "_2")]
        [InlineData("------------------------------9", 9, "------------------------------", "-----------------------------10")]
        [InlineData("9999999999999999999999999999999", 9, "999999999999999999999999999999", "0")] // special case
        public void SanitizeWorksheetNameTest(String givenName, int numberOfExistingWorksheets, string existingWorksheetPrefix, string expectedName)
        {
            Workbook workbook = new Workbook(false);
            for(int i = 0; i < numberOfExistingWorksheets; i++)
            {
                workbook.AddWorksheet(existingWorksheetPrefix+ (i + 1).ToString());
            }
            string name = Worksheet.SanitizeWorksheetName(givenName, workbook);
            Assert.Equal(expectedName, name);
        }

    }
}
