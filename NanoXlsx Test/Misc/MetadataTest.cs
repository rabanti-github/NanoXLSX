using NanoXLSX;
using NanoXLSX.Shared.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;
using FormatException = NanoXLSX.Shared.Exceptions.FormatException;

namespace NanoXLSX_Test.Misc
{
    public class MetadataTest
    {
        [Fact(DisplayName = "Test of the get and set function of the Application property")]
        public void ApplicationTest()
        {
            Metadata metadata = new Metadata();
            Assert.NotNull(metadata.Application);
            Assert.NotEmpty(metadata.Application);
            metadata.Application = "test";
            Assert.Equal("test", metadata.Application);
        }

        [Theory(DisplayName = "Test of the get and set function of the ApplicationVersion property")]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("0.1")]
        [InlineData("99999.99999")]
        public void ApplicationVersionTest(string version)
        {
            Metadata metadata = new Metadata();
            Assert.NotNull(metadata.ApplicationVersion);
            Assert.NotEmpty(metadata.ApplicationVersion);
            metadata.ApplicationVersion = version;
            Assert.Equal(version, metadata.ApplicationVersion);
        }

        [Theory(DisplayName = "Test of failing set function of the ApplicationVersion property on invalid versions")]
        [InlineData("1")]
        [InlineData("1.2.3")]
        [InlineData(" ")]
        [InlineData("xyz")]
        [InlineData("111111.1")]
        [InlineData("1.222222")]
        [InlineData("333333.333333")]
        public void ApplicationVersionFailTest(string version)
        {
            Metadata metadata = new Metadata();
            Assert.Throws<FormatException>(() => metadata.ApplicationVersion = version);
        }

        [Fact(DisplayName = "Test of the get and set function of the Category property")]
        public void CategoryTest()
        {
            Metadata metadata = new Metadata();
            Assert.Null(metadata.Category);
            metadata.Category = "test";
            Assert.Equal("test", metadata.Category);
        }

        [Fact(DisplayName = "Test of the get and set function of the Company property")]
        public void CompanyTest()
        {
            Metadata metadata = new Metadata();
            Assert.Null(metadata.Company);
            metadata.Company = "test";
            Assert.Equal("test", metadata.Company);
        }

        [Fact(DisplayName = "Test of the get and set function of the ContentStatus property")]
        public void ContentStatusTest()
        {
            Metadata metadata = new Metadata();
            Assert.Null(metadata.ContentStatus);
            metadata.ContentStatus = "test";
            Assert.Equal("test", metadata.ContentStatus);
        }

        [Fact(DisplayName = "Test of the get and set function of the Creator property")]
        public void CreatorTest()
        {
            Metadata metadata = new Metadata();
            Assert.Null(metadata.Creator);
            metadata.Creator = "test";
            Assert.Equal("test", metadata.Creator);
        }

        [Fact(DisplayName = "Test of the get and set function of the Description property")]
        public void DescriptionTest()
        {
            Metadata metadata = new Metadata();
            Assert.Null(metadata.Description);
            metadata.Description = "test";
            Assert.Equal("test", metadata.Description);
        }

        [Fact(DisplayName = "Test of the get and set function of the HyperlinkBase property")]
        public void HyperlinkBaseTest()
        {
            Metadata metadata = new Metadata();
            Assert.Null(metadata.HyperlinkBase);
            metadata.HyperlinkBase = "test";
            Assert.Equal("test", metadata.HyperlinkBase);
        }

        [Fact(DisplayName = "Test of the get and set function of the Keywords property")]
        public void KeywordsTest()
        {
            Metadata metadata = new Metadata();
            Assert.Null(metadata.Keywords);
            metadata.Keywords = "test";
            Assert.Equal("test", metadata.Keywords);
        }

        [Fact(DisplayName = "Test of the get and set function of the Manager property")]
        public void ManagerTest()
        {
            Metadata metadata = new Metadata();
            Assert.Null(metadata.Manager);
            metadata.Manager = "test";
            Assert.Equal("test", metadata.Manager);
        }

        [Fact(DisplayName = "Test of the get and set function of the Subject property")]
        public void SubjectTest()
        {
            Metadata metadata = new Metadata();
            Assert.Null(metadata.Subject);
            metadata.Subject = "test";
            Assert.Equal("test", metadata.Subject);
        }

        [Fact(DisplayName = "Test of the get and set function of the Title property")]
        public void TitleTest()
        {
            Metadata metadata = new Metadata();
            Assert.Null(metadata.Title);
            metadata.Title = "test";
            Assert.Equal("test", metadata.Title);
        }

        [Fact(DisplayName = "Test of the Constructor")]
        public void ConstructorTest()
        {
            Metadata metadata = new Metadata();
            Assert.NotNull(metadata);
            Assert.NotEmpty(metadata.Application);
            Assert.NotEmpty(metadata.ApplicationVersion);
        }

        [Theory(DisplayName = "Test of the ParseVersion function")]
        [InlineData(1, 2, 2, 5, "1.225")]
        [InlineData(4, 2, 2, 0, "4.22")]
        [InlineData(11, 2, 0, 0, "11.2")]
        [InlineData(112, 0, 0, 0, "112.0")]
        [InlineData(0, 0, 0, 0, "0.0")]
        [InlineData(0, 4, 5, 1, "0.451")]
        [InlineData(0, 0, 2, 1, "0.021")]
        [InlineData(0, 0, 0, 1, "0.001")]
        [InlineData(9999, 666, 555, 444, "9999.66655")]
        [InlineData(99999, 0, 0, 1234567, "99999.00123")]
        public void ParseVersionTest(int major, int minor, int build, int revision, string expectedVersion)
        {
            string version = Metadata.ParseVersion(major, minor, build, revision);
            Assert.Equal(expectedVersion, version);
        }

        [Theory(DisplayName = "Test of the failingParseVersion function")]
        [InlineData(111111, 1, 1, 1)]
        [InlineData(-1, 1, 1, 1)]
        [InlineData(1, -1, 1, 1)]
        [InlineData(1, 1, -1, 1)]
        [InlineData(1, 1, 1, -1)]
        public void ParseVersionFailTest(int major, int minor, int build, int revision)
        {
            Assert.Throws<FormatException>(() => Metadata.ParseVersion(major, minor, build, revision));
        }

    }
}
