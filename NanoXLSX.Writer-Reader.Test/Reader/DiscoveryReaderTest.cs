using System;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using NanoXLSX.Interfaces.Reader;
using NanoXLSX.Internal;
using NanoXLSX.Internal.Readers;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;
using Xunit;

namespace NanoXLSX.Test.Writer_Reader.ReaderTest
{
    [Collection(nameof(SequentialCollection2))]
    public class DiscoveryReaderTest : IDisposable
    {
        public void Dispose()
        {
            PlugInLoader.DisposePlugins();
        }

        [Fact(DisplayName = "Discovery reader registration uses the discovery UUID")]
        public void RegistrationTest()
        {
            NanoXlsxPlugInAttribute attribute = typeof(DiscoveryReader).GetCustomAttribute<NanoXlsxPlugInAttribute>();

            Assert.NotNull(attribute);
            Assert.Equal(PlugInUUID.DiscoveryReader, attribute.PlugInUUID);
        }

        [Fact(DisplayName = "Discovery reader can be replaced through the plugin loader")]
        public void ReplacementRegistrationTest()
        {
            PlugInLoader.InjectPlugins(new System.Collections.Generic.List<Type> { typeof(ReplacementDiscoveryReader) });

            IDiscoveryReader reader = PlugInLoader.GetPlugIn<IDiscoveryReader>(PlugInUUID.DiscoveryReader, new DiscoveryReader());

            Assert.IsType<ReplacementDiscoveryReader>(reader);
        }

        [Fact(DisplayName = "Discovery reader prepares temporary catalog and leaves archive open")]
        public void PrepareCatalogTest()
        {
            using (MemoryStream stream = CreateArchive())
            using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read, true))
            {
                Workbook workbook = new Workbook("Sheet1");
                ReaderOptions options = new ReaderOptions { EnforceStrictValidation = true };
                DiscoveryReader reader = new DiscoveryReader();

                reader.Init(archive, workbook, options);
                reader.Execute();

                RelationshipCatalog catalog = workbook.AuxiliaryData.GetData<RelationshipCatalog>(PlugInUUID.DiscoveryReader, PlugInUUID.DiscoveryCatalogEntity);
                Assert.NotNull(catalog);
                Assert.True(catalog.IsComplete);
                Assert.Empty(catalog.Relationships);
                Assert.Same(workbook, reader.Workbook);
                Assert.Same(options, reader.Options);
                Assert.Single(archive.Entries);
                using (Stream entryStream = archive.Entries[0].Open())
                {
                    Assert.True(entryStream.CanRead);
                }

                workbook.AuxiliaryData.ClearTemporaryData();
                Assert.Null(workbook.AuxiliaryData.GetData<RelationshipCatalog>(PlugInUUID.DiscoveryReader, PlugInUUID.DiscoveryCatalogEntity));
            }
        }

        [Fact(DisplayName = "Discovery reader requires a ZIP archive")]
        public void ArchiveInitializationRequiredTest()
        {
            DiscoveryReader reader = new DiscoveryReader();
            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => reader.Execute());

            Assert.Equal("The discovery reader was not initialized with a ZIP archive.", exception.Message);
        }

        [Fact(DisplayName = "Discovery reader requires a workbook")]
        public void WorkbookInitializationRequiredTest()
        {
            using (MemoryStream stream = CreateArchive())
            using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read, true))
            {
                DiscoveryReader reader = new DiscoveryReader();
                reader.Init(archive, null, new ReaderOptions());

                InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => reader.Execute());

                Assert.Equal("The discovery reader was not initialized with a workbook.", exception.Message);
            }
        }

        [Theory(DisplayName = "Built-in document readers expose exact relationship types")]
        [InlineData(typeof(MetadataAppReader), "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties")]
        [InlineData(typeof(MetadataCoreReader), "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties")]
        [InlineData(typeof(SharedStringsReader), "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings")]
        [InlineData(typeof(StyleReader), "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles")]
        [InlineData(typeof(ThemeReader), "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")]
        [InlineData(typeof(WorkbookReader), "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument")]
        [InlineData(typeof(WorksheetReader), "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")]
        public void DocumentTypeTest(Type readerType, string expectedDocumentType)
        {
            IDocumentReader reader = (IDocumentReader)Activator.CreateInstance(readerType, true);
            Assert.Equal(expectedDocumentType, reader.DocumentType);
        }

        private static MemoryStream CreateArchive()
        {
            MemoryStream stream = new MemoryStream();
            using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Create, true))
            {
                archive.CreateEntry("xl/workbook.xml");
            }
            stream.Position = 0;
            return stream;
        }

        [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.DiscoveryReader, PlugInOrder = 100)]
        private class ReplacementDiscoveryReader : DiscoveryReader
        {
        }
    }
}
