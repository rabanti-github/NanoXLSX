using NanoXLSX;
using Xunit;

namespace NanoXLSX_Test.Misc
{
    // Ensure that these tests are executed sequentially, since static repository methods may be called 
    [Collection(nameof(SequentialCollection))]
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

        [Fact(DisplayName = "Test of the 'Description' property when writing and reading a workbook")]
        public void ReadDescriptionTest()
        {
            Workbook workbook = new Workbook();
            workbook.WorkbookMetadata.Description = "description1";
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.Equal("description1", givenWorkbook.WorkbookMetadata.Description);
        }

        [Fact(DisplayName = "Test of the 'HyperlinkBase' property when writing and reading a workbook")]
        public void ReadHyperlinkBaseTest()
        {
            Workbook workbook = new Workbook();
            workbook.WorkbookMetadata.HyperlinkBase = "hyperlinkBase1";
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.Equal("hyperlinkBase1", givenWorkbook.WorkbookMetadata.HyperlinkBase);
        }

        [Fact(DisplayName = "Test of the 'Keywords' property when writing and reading a workbook")]
        public void ReadKeywordsTest()
        {
            Workbook workbook = new Workbook();
            workbook.WorkbookMetadata.Keywords = "keyword1;keyword2";
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.Equal("keyword1;keyword2", givenWorkbook.WorkbookMetadata.Keywords);
        }

        [Fact(DisplayName = "Test of the 'Manager' property when writing and reading a workbook")]
        public void ReadManagerTest()
        {
            Workbook workbook = new Workbook();
            workbook.WorkbookMetadata.Manager = "manager1";
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.Equal("manager1", givenWorkbook.WorkbookMetadata.Manager);
        }

        [Fact(DisplayName = "Test of the 'Subject' property when writing and reading a workbook")]
        public void ReadSubjectTest()
        {
            Workbook workbook = new Workbook();
            workbook.WorkbookMetadata.Subject = "subject1";
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.Equal("subject1", givenWorkbook.WorkbookMetadata.Subject);
        }

        [Fact(DisplayName = "Test of the 'Title' property when writing and reading a workbook")]
        public void ReadTitleTest()
        {
            Workbook workbook = new Workbook();
            workbook.WorkbookMetadata.Title = "title1";
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.Equal("title1", givenWorkbook.WorkbookMetadata.Title);
        }

        [Fact(DisplayName = "Test of writing and reading a workbook with a null WorkbookMetadata object")]
        public void ReadNullTest()
        {
            Workbook workbook = new Workbook();
            workbook.WorkbookMetadata = null;
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.NotNull(givenWorkbook.WorkbookMetadata);
            Assert.Null(givenWorkbook.WorkbookMetadata.Application);
            Assert.Null(givenWorkbook.WorkbookMetadata.Creator);
            Assert.Null(givenWorkbook.WorkbookMetadata.Title);
        }

    }
}
