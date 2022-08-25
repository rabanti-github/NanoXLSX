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
    }
}
