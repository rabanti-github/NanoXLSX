using NanoXLSX;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace NanoXLSX_Test.Misc
{
    public class MetadataWriteReadTest
    {

        [Fact(DisplayName = "Test of the 'Application' property when writing and reading a workbook")]
        public void ReadApplicationTest()
        {
            Workbook workbook = new Workbook();
            workbook.WorkbookMetadata.Application = "testApp";
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.Equal("testApp", givenWorkbook.WorkbookMetadata.Application);
        }

        [Fact(DisplayName = "Test of the 'ApplicationVersion' property when writing and reading a workbook")]
        public void ReadApplicationVersionTest()
        {
            Workbook workbook = new Workbook();
            workbook.WorkbookMetadata.ApplicationVersion = "3.456";
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.Equal("3.456", givenWorkbook.WorkbookMetadata.ApplicationVersion);
        }

        [Fact(DisplayName = "Test of the 'Category' property when writing and reading a workbook")]
        public void ReadCategoryTest()
        {
            Workbook workbook = new Workbook();
            workbook.WorkbookMetadata.Category = "cat1";
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.Equal("cat1", givenWorkbook.WorkbookMetadata.Category);
        }

        [Fact(DisplayName = "Test of the 'Company' property when writing and reading a workbook")]
        public void ReadCompanyTest()
        {
            Workbook workbook = new Workbook();
            workbook.WorkbookMetadata.Company = "company1";
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.Equal("company1", givenWorkbook.WorkbookMetadata.Company);
        }

        [Fact(DisplayName = "Test of the 'ContentStatus' property when writing and reading a workbook")]
        public void ReadContentStatusTest()
        {
            Workbook workbook = new Workbook();
            workbook.WorkbookMetadata.ContentStatus = "status1";
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.Equal("status1", givenWorkbook.WorkbookMetadata.ContentStatus);
        }

        [Fact(DisplayName = "Test of the 'Creator' property when writing and reading a workbook")]
        public void ReadCreatorTest()
        {
            Workbook workbook = new Workbook();
            workbook.WorkbookMetadata.Creator = "creator1";
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.Equal("creator1", givenWorkbook.WorkbookMetadata.Creator);
        }

    }
}
