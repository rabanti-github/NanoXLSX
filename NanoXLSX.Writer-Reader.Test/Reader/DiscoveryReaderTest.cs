using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Text;
using System.IO.Packaging;
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

        [Fact(DisplayName = "Discovery reader scans all relationship parts and resolves OPC targets")]
        public void DiscoverRelationshipPartsTest()
        {
            using (MemoryStream stream = CreateRelationshipArchive(
                ("xl/worksheets/_rels/sheet1.xml.rels", RelationshipsXml(
                    RelationshipXml("rIdDrawing", "http://example.test/relationships/drawing", "../drawings/drawing1.xml"))),
                ("xl/drawings/_rels/drawing1.xml.rels", RelationshipsXml(
                    RelationshipXml("rIdChart", "http://example.test/relationships/chart", "../charts/chart1.xml"))),
                ("xl/externalLinks/_rels/externalLink1.xml.rels", RelationshipsXml(
                    RelationshipXml("rIdExternal", "http://example.test/relationships/external", "https://example.test/resource", "External"))),
                ("_rels/.rels", RelationshipsXml(
                    RelationshipXml("rIdWorkbook", new WorkbookReader().DocumentType, "/xl/workbook.xml"))),
                ("xl/_rels/workbook.xml.rels", RelationshipsXml(
                    RelationshipXml("rIdSheet", new WorksheetReader().DocumentType, "worksheets/sheet1.xml"),
                    RelationshipXml("rIdExternalLink", "http://example.test/relationships/externalLink", "externalLinks/externalLink1.xml")))))
            using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read, true))
            {
                Workbook workbook = new Workbook("Sheet1");
                DiscoveryReader reader = new DiscoveryReader();
                reader.Init(archive, workbook, new ReaderOptions());

                reader.Execute();

                RelationshipCatalog catalog = workbook.AuxiliaryData.GetData<RelationshipCatalog>(PlugInUUID.DiscoveryReader, PlugInUUID.DiscoveryCatalogEntity);
                Assert.Equal(6, catalog.Relationships.Count);
                Assert.Equal("_rels/.rels", catalog.Relationships[0].RelationshipPartPath);
                Assert.Equal("xl/_rels/workbook.xml.rels", catalog.Relationships[1].RelationshipPartPath);
                Assert.Equal("xl/drawings/_rels/drawing1.xml.rels", catalog.Relationships[3].RelationshipPartPath);
                Assert.Equal("xl/externalLinks/_rels/externalLink1.xml.rels", catalog.Relationships[4].RelationshipPartPath);
                Assert.Equal("xl/worksheets/_rels/sheet1.xml.rels", catalog.Relationships[5].RelationshipPartPath);
                Assert.Equal(string.Empty, catalog.Relationships[0].SourcePartPath);
                Assert.Equal("xl/workbook.xml", catalog.Relationships[0].ResolvedTargetPath);
                Assert.Equal("xl/worksheets/sheet1.xml", catalog.Relationships[1].ResolvedTargetPath);
                Assert.Equal("xl/externalLinks/externalLink1.xml", catalog.Relationships[2].ResolvedTargetPath);
                Assert.Equal("xl/charts/chart1.xml", catalog.Relationships[3].ResolvedTargetPath);
                Assert.Equal(TargetMode.External, catalog.Relationships[4].TargetMode);
                Assert.Equal("https://example.test/resource", catalog.Relationships[4].Target);
                Assert.Null(catalog.Relationships[4].ResolvedTargetPath);
                Assert.Equal("xl/drawings/drawing1.xml", catalog.Relationships[5].ResolvedTargetPath);
                Assert.True(catalog.IsComplete);
            }
        }

        [Fact(DisplayName = "Tolerant discovery skips invalid entries and retains the first duplicate")]
        public void TolerantValidationTest()
        {
            string relationships = RelationshipsXml(
                RelationshipXml("rId1", "http://example.test/relationships/type", "first.xml"),
                RelationshipXml("rId1", "http://example.test/relationships/type", "second.xml"),
                "<Relationship Id=\"rId2\" Type=\"http://example.test/relationships/type\" />");
            using (MemoryStream stream = CreateRelationshipArchive(("_rels/.rels", relationships)))
            using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read, true))
            {
                Workbook workbook = new Workbook("Sheet1");
                DiscoveryReader reader = new DiscoveryReader();
                reader.Init(archive, workbook, new ReaderOptions());

                reader.Execute();

                RelationshipCatalog catalog = workbook.AuxiliaryData.GetData<RelationshipCatalog>(PlugInUUID.DiscoveryReader, PlugInUUID.DiscoveryCatalogEntity);
                Assert.Single(catalog.Relationships);
                Assert.Equal("first.xml", catalog.Relationships[0].Target);
                Assert.Equal(2, catalog.Issues.Count);
                Assert.False(catalog.IsComplete);
            }
        }

        [Fact(DisplayName = "Tolerant discovery discards a malformed relationship part transactionally")]
        public void TolerantMalformedPartTest()
        {
            string malformed = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + RelationshipXml("rId1", "http://example.test/relationships/type", "retained-only-if-invalid.xml");
            using (MemoryStream stream = CreateRelationshipArchive(("_rels/.rels", malformed)))
            using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read, true))
            {
                Workbook workbook = new Workbook("Sheet1");
                DiscoveryReader reader = new DiscoveryReader();
                reader.Init(archive, workbook, new ReaderOptions());

                reader.Execute();

                RelationshipCatalog catalog = workbook.AuxiliaryData.GetData<RelationshipCatalog>(PlugInUUID.DiscoveryReader, PlugInUUID.DiscoveryCatalogEntity);
                Assert.Empty(catalog.Relationships);
                Assert.Single(catalog.Issues);
                Assert.False(catalog.IsComplete);
            }
        }

        [Fact(DisplayName = "Strict discovery rejects a duplicate relationship identifier")]
        public void StrictDuplicateTest()
        {
            string relationships = RelationshipsXml(
                RelationshipXml("rId1", "http://example.test/relationships/type", "first.xml"),
                RelationshipXml("rId1", "http://example.test/relationships/type", "second.xml"));
            using (MemoryStream stream = CreateRelationshipArchive(("_rels/.rels", relationships)))
            using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read, true))
            {
                DiscoveryReader reader = new DiscoveryReader();
                reader.Init(archive, new Workbook("Sheet1"), new ReaderOptions { EnforceStrictValidation = true });

                NanoXLSX.Exceptions.IOException exception = Assert.Throws<NanoXLSX.Exceptions.IOException>(() => reader.Execute());

                Assert.Contains("rId1", exception.Message);
                Assert.Contains("duplicated", exception.Message);
            }
        }

        public static IEnumerable<object[]> InvalidRelationshipCases
        {
            get
            {
                yield return new object[] { "<Relationship Type=\"http://example.test/type\" Target=\"part.xml\" />", "Id attribute" };
                yield return new object[] { "<Relationship Id=\"rIdInvalid\" Target=\"part.xml\" />", "Type attribute" };
                yield return new object[] { "<Relationship Id=\"rIdInvalid\" Type=\"http://example.test/type\" />", "Target attribute" };
                yield return new object[] { RelationshipXml("invalid:id", "http://example.test/type", "part.xml"), "NCName" };
                yield return new object[] { RelationshipXml("rIdInvalid", "relative-type", "part.xml"), "absolute URI" };
                yield return new object[] { RelationshipXml("rIdInvalid", "http://example.test/type", "http://[", "External"), "valid URI" };
                yield return new object[] { RelationshipXml("rIdInvalid", "http://example.test/type", "part.xml", "Remote"), "TargetMode" };
                yield return new object[] { RelationshipXml("rIdInvalid", "http://example.test/type", "https://example.test/part.xml"), "absolute URI" };
            }
        }

        [Theory(DisplayName = "Tolerant discovery skips every invalid relationship entry category")]
        [MemberData(nameof(InvalidRelationshipCases))]
        public void TolerantInvalidEntryMatrixTest(string invalidRelationship, string expectedReason)
        {
            string relationships = RelationshipsXml(
                RelationshipXml("rIdValid", "http://example.test/type", "valid.xml"),
                invalidRelationship);
            using (MemoryStream stream = CreateRelationshipArchive(("_rels/.rels", relationships)))
            using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read, true))
            {
                Workbook workbook = new Workbook("Sheet1");
                DiscoveryReader reader = new DiscoveryReader();
                reader.Init(archive, workbook, null);

                reader.Execute();

                RelationshipCatalog catalog = workbook.AuxiliaryData.GetData<RelationshipCatalog>(PlugInUUID.DiscoveryReader, PlugInUUID.DiscoveryCatalogEntity);
                Assert.Single(catalog.Relationships);
                RelationshipDiscoveryIssue issue = Assert.Single(catalog.Issues);
                Assert.Contains(expectedReason, issue.Reason);
                Assert.False(catalog.IsComplete);
            }
        }

        [Theory(DisplayName = "Strict discovery rejects every invalid relationship entry category")]
        [MemberData(nameof(InvalidRelationshipCases))]
        public void StrictInvalidEntryMatrixTest(string invalidRelationship, string expectedReason)
        {
            using (MemoryStream stream = CreateRelationshipArchive(("_rels/.rels", RelationshipsXml(invalidRelationship))))
            using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read, true))
            {
                DiscoveryReader reader = new DiscoveryReader();
                reader.Init(archive, new Workbook("Sheet1"), new ReaderOptions { EnforceStrictValidation = true });

                NanoXLSX.Exceptions.IOException exception = Assert.Throws<NanoXLSX.Exceptions.IOException>(() => reader.Execute());

                Assert.Contains(expectedReason, exception.Message);
            }
        }

        [Fact(DisplayName = "Strict discovery rejects malformed relationship XML")]
        public void StrictMalformedPartTest()
        {
            string malformed = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">";
            using (MemoryStream stream = CreateRelationshipArchive(("_rels/.rels", malformed)))
            using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read, true))
            {
                DiscoveryReader reader = new DiscoveryReader();
                reader.Init(archive, new Workbook("Sheet1"), new ReaderOptions { EnforceStrictValidation = true });

                NanoXLSX.Exceptions.IOException exception = Assert.Throws<NanoXLSX.Exceptions.IOException>(() => reader.Execute());

                Assert.Contains("_rels/.rels", exception.Message);
                Assert.Contains("could not be parsed", exception.Message);
            }
        }

        [Theory(DisplayName = "Discovery validates relationship-part paths according to validation mode")]
        [InlineData(false)]
        [InlineData(true)]
        public void InvalidRelationshipPartPathTest(bool strict)
        {
            using (MemoryStream stream = CreateRelationshipArchive(("_rels/bad name.rels", RelationshipsXml(
                RelationshipXml("rId1", "http://example.test/type", "part.xml")))))
            using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read, true))
            {
                Workbook workbook = new Workbook("Sheet1");
                DiscoveryReader reader = new DiscoveryReader();
                reader.Init(archive, workbook, new ReaderOptions { EnforceStrictValidation = strict });

                if (strict)
                {
                    NanoXLSX.Exceptions.IOException exception = Assert.Throws<NanoXLSX.Exceptions.IOException>(() => reader.Execute());
                    Assert.Contains("relationship-part path", exception.Message);
                }
                else
                {
                    reader.Execute();
                    RelationshipCatalog catalog = workbook.AuxiliaryData.GetData<RelationshipCatalog>(PlugInUUID.DiscoveryReader, PlugInUUID.DiscoveryCatalogEntity);
                    Assert.Empty(catalog.Relationships);
                    Assert.Single(catalog.Issues);
                }
            }
        }

        [Theory(DisplayName = "Malformed unconsumed relationship parts follow strict or tolerant loading mode")]
        [InlineData(false)]
        [InlineData(true)]
        public void MalformedOptionalPartIntegrationTest(bool strict)
        {
            const string malformed = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">";
            using (MemoryStream stream = CreateWorkbookWithAdditionalRelationships(
                string.Empty,
                ("xl/custom/_rels/optional.xml.rels", malformed)))
            {
                ReaderOptions options = new ReaderOptions { EnforceStrictValidation = strict };
                if (strict)
                {
                    Assert.Throws<NanoXLSX.Exceptions.IOException>(() => NanoXLSX.Extensions.WorkbookReader.Load(stream, options));
                }
                else
                {
                    Workbook workbook = NanoXLSX.Extensions.WorkbookReader.Load(stream, options);
                    Assert.Single(workbook.Worksheets);
                }
            }
        }

        [Fact(DisplayName = "XLSX reader resolves renamed built-in parts through discovery")]
        public void RenamedBuiltInPartsTest()
        {
            using (MemoryStream stream = CreateRenamedWorkbookArchive())
            {
                Workbook workbook = NanoXLSX.Extensions.WorkbookReader.Load(stream);

                Assert.Single(workbook.Worksheets);
                Assert.Equal("discovered", workbook.Worksheets[0].GetCell(new Address("A1")).Value.ToString());
                Assert.Null(workbook.AuxiliaryData.GetData<RelationshipCatalog>(PlugInUUID.DiscoveryReader, PlugInUUID.DiscoveryCatalogEntity));
            }
        }

        [Theory(DisplayName = "XLSX reader resolves independently renamed optional built-in parts")]
        [InlineData("xl/theme/theme1.xml", "xl/custom/look3.xml", "theme/theme1.xml", "custom/look3.xml")]
        [InlineData("xl/sharedStrings.xml", "xl/custom/strings8.xml", "sharedStrings.xml", "custom/strings8.xml")]
        [InlineData("docProps/app.xml", "info/app9.xml", "docProps/app.xml", "info/app9.xml")]
        [InlineData("docProps/core.xml", "info/core9.xml", "docProps/core.xml", "info/core9.xml")]
        public void IndependentlyRenamedOptionalPartTest(string sourcePath, string renamedPath, string relationshipTarget, string renamedRelationshipTarget)
        {
            using (MemoryStream stream = CreateWorkbookWithRenamedPart(sourcePath, renamedPath, relationshipTarget, renamedRelationshipTarget))
            {
                Workbook workbook = NanoXLSX.Extensions.WorkbookReader.Load(stream);

                Assert.Single(workbook.Worksheets);
                Assert.Equal("discovered", workbook.Worksheets[0].GetCell(new Address("A1")).Value.ToString());
            }
        }

        [Fact(DisplayName = "Discovery-aware registry reader executes once per matching target")]
        public void DiscoveryPackageReaderDispatchTest()
        {
            PlugInLoader.InjectPlugins(new List<Type> { typeof(WorksheetDiscoveryPackageReader) });
            Workbook sourceWorkbook = new Workbook("Sheet1");
            sourceWorkbook.AddWorksheet("Sheet2");
            using (MemoryStream stream = new MemoryStream())
            {
                sourceWorkbook.SaveAsStream(stream, true);
                stream.Position = 0;

                Workbook workbook = NanoXLSX.Extensions.WorkbookReader.Load(stream);

                Assert.Equal(2, workbook.AuxiliaryData.GetData<int>(WorksheetDiscoveryPackageReader.PLUGIN_ID, 0));
                Assert.Equal("xl/worksheets/sheet2.xml", workbook.AuxiliaryData.GetData<string>(WorksheetDiscoveryPackageReader.PLUGIN_ID, 1));
                Assert.Null(workbook.AuxiliaryData.GetData<RelationshipCatalog>(PlugInUUID.DiscoveryReader, PlugInUUID.DiscoveryCatalogEntity));
            }
        }

        [Fact(DisplayName = "Discovery catalog is available to registry and prepending queues in order")]
        public void DiscoveryQueueOrderTest()
        {
            PlugInLoader.InjectPlugins(new List<Type> { typeof(CatalogRegistryReader), typeof(CatalogPrependingReader) });
            using (MemoryStream stream = CreateWrittenWorkbookArchive())
            {
                Workbook workbook = NanoXLSX.Extensions.WorkbookReader.Load(stream);

                Assert.Equal("registry", workbook.AuxiliaryData.GetData<string>(CatalogRegistryReader.PLUGIN_ID, 0));
                Assert.Equal("prepending", workbook.AuxiliaryData.GetData<string>(CatalogRegistryReader.PLUGIN_ID, 1));
            }
        }

        [Fact(DisplayName = "Discovery package dispatch compares relationship types ordinally and skips external targets")]
        public void ExactDocumentTypeDispatchTest()
        {
            const string exactType = "http://example.test/relationships/custom";
            string relationships = RelationshipXml("rIdExact", exactType, "custom/exact1.xml")
                + RelationshipXml("rIdCase", "HTTP://example.test/relationships/custom", "custom/case2.xml")
                + RelationshipXml("rIdEscaped", "http://example.test/relationships/%63ustom", "custom/escaped3.xml")
                + RelationshipXml("rIdExternal", exactType, "https://example.test/external", "External");
            using (MemoryStream stream = CreateWorkbookWithAdditionalRelationships(
                relationships,
                ("xl/custom/exact1.xml", "exact"),
                ("xl/custom/case2.xml", "case"),
                ("xl/custom/escaped3.xml", "escaped")))
            {
                PlugInLoader.InjectPlugins(new List<Type> { typeof(ExactTypeDiscoveryPackageReader) });

                Workbook workbook = NanoXLSX.Extensions.WorkbookReader.Load(stream);

                Assert.Equal(1, workbook.AuxiliaryData.GetData<int>(ExactTypeDiscoveryPackageReader.PLUGIN_ID, 0));
                Assert.Equal("xl/custom/exact1.xml", workbook.AuxiliaryData.GetData<string>(ExactTypeDiscoveryPackageReader.PLUGIN_ID, 1));
            }
        }

        [Fact(DisplayName = "XLSX reader processes only the first discovered theme relationship")]
        public void MultipleThemeRelationshipTest()
        {
            string secondThemeRelationship = RelationshipXml("rIdSecondTheme", new ThemeReader().DocumentType, "theme/theme2.xml");
            using (MemoryStream stream = CreateWorkbookWithAdditionalRelationships(
                secondThemeRelationship,
                ("xl/theme/theme2.xml", "SECOND_THEME")))
            {
                PlugInLoader.InjectPlugins(new List<Type> { typeof(ThemeSelectionReader) });

                Workbook workbook = NanoXLSX.Extensions.WorkbookReader.Load(stream);

                Assert.Equal(1, workbook.AuxiliaryData.GetData<int>(ThemeSelectionReader.PLUGIN_ID, 0));
                Assert.DoesNotContain("SECOND_THEME", workbook.AuxiliaryData.GetData<string>(ThemeSelectionReader.PLUGIN_ID, 1));
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

        private static MemoryStream CreateWrittenWorkbookArchive()
        {
            Workbook workbook = new Workbook("Sheet1");
            workbook.CurrentWorksheet.AddCell("discovered", "A1");
            MemoryStream stream = new MemoryStream();
            workbook.SaveAsStream(stream, true);
            stream.Position = 0;
            return stream;
        }

        private static MemoryStream CreateRelationshipArchive(params (string Path, string Xml)[] entries)
        {
            MemoryStream stream = new MemoryStream();
            using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Create, true))
            {
                foreach ((string Path, string Xml) entry in entries)
                {
                    ZipArchiveEntry archiveEntry = archive.CreateEntry(entry.Path);
                    using (Stream entryStream = archiveEntry.Open())
                    {
                        byte[] content = Encoding.UTF8.GetBytes(entry.Xml);
                        entryStream.Write(content, 0, content.Length);
                    }
                }
            }
            stream.Position = 0;
            return stream;
        }

        private static string RelationshipsXml(params string[] relationships)
        {
            return "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + string.Concat(relationships)
                + "</Relationships>";
        }

        private static string RelationshipXml(string id, string type, string target, string targetMode = null)
        {
            string targetModeAttribute = targetMode == null ? string.Empty : " TargetMode=\"" + targetMode + "\"";
            return "<Relationship Id=\"" + id + "\" Type=\"" + type + "\" Target=\"" + target + "\"" + targetModeAttribute + " />";
        }

        private static MemoryStream CreateRenamedWorkbookArchive()
        {
            Workbook workbook = new Workbook("Sheet1");
            workbook.CurrentWorksheet.AddCell("discovered", "A1");
            MemoryStream originalStream = new MemoryStream();
            workbook.SaveAsStream(originalStream, true);
            originalStream.Position = 0;

            Dictionary<string, string> renamedPaths = new Dictionary<string, string>(StringComparer.Ordinal)
            {
                { "xl/workbook.xml", "custom/book7.xml" },
                { "xl/_rels/workbook.xml.rels", "custom/_rels/book7.xml.rels" },
                { "xl/worksheets/sheet1.xml", "parts/grid42.xml" },
                { "xl/styles.xml", "parts/design9.xml" },
                { "xl/theme/theme1.xml", "parts/look3.xml" },
                { "xl/sharedStrings.xml", "parts/strings8.xml" }
            };
            MemoryStream renamedStream = new MemoryStream();
            using (originalStream)
            using (ZipArchive originalArchive = new ZipArchive(originalStream, ZipArchiveMode.Read, true))
            using (ZipArchive renamedArchive = new ZipArchive(renamedStream, ZipArchiveMode.Create, true))
            {
                foreach (ZipArchiveEntry originalEntry in originalArchive.Entries)
                {
                    string targetPath = renamedPaths.TryGetValue(originalEntry.FullName, out string renamedPath)
                        ? renamedPath
                        : originalEntry.FullName;
                    ZipArchiveEntry renamedEntry = renamedArchive.CreateEntry(targetPath);
                    using (Stream source = originalEntry.Open())
                    using (Stream target = renamedEntry.Open())
                    {
                        if (originalEntry.FullName.EndsWith(".xml", StringComparison.Ordinal)
                            || originalEntry.FullName.EndsWith(".rels", StringComparison.Ordinal))
                        {
                            using (StreamReader reader = new StreamReader(source, Encoding.UTF8, true, 1024, true))
                            using (StreamWriter writer = new StreamWriter(target, new UTF8Encoding(false), 1024, true))
                            {
                                string content = RenameWorkbookContent(reader.ReadToEnd());
                                writer.Write(content);
                            }
                        }
                        else
                        {
                            source.CopyTo(target);
                        }
                    }
                }
            }
            renamedStream.Position = 0;
            return renamedStream;
        }

        private static MemoryStream CreateWorkbookWithRenamedPart(string sourcePath, string renamedPath, string relationshipTarget, string renamedRelationshipTarget)
        {
            Workbook workbook = new Workbook("Sheet1");
            workbook.CurrentWorksheet.AddCell("discovered", "A1");
            MemoryStream originalStream = new MemoryStream();
            workbook.SaveAsStream(originalStream, true);
            originalStream.Position = 0;

            bool sourceFound = false;
            MemoryStream renamedStream = new MemoryStream();
            using (originalStream)
            using (ZipArchive originalArchive = new ZipArchive(originalStream, ZipArchiveMode.Read, true))
            using (ZipArchive renamedArchive = new ZipArchive(renamedStream, ZipArchiveMode.Create, true))
            {
                foreach (ZipArchiveEntry originalEntry in originalArchive.Entries)
                {
                    bool isRenamedEntry = string.Equals(originalEntry.FullName, sourcePath, StringComparison.Ordinal);
                    sourceFound |= isRenamedEntry;
                    ZipArchiveEntry renamedEntry = renamedArchive.CreateEntry(isRenamedEntry ? renamedPath : originalEntry.FullName);
                    using (Stream source = originalEntry.Open())
                    using (Stream target = renamedEntry.Open())
                    {
                        if (originalEntry.FullName.EndsWith(".xml", StringComparison.Ordinal)
                            || originalEntry.FullName.EndsWith(".rels", StringComparison.Ordinal))
                        {
                            using (StreamReader reader = new StreamReader(source, Encoding.UTF8, true, 1024, true))
                            using (StreamWriter writer = new StreamWriter(target, new UTF8Encoding(false), 1024, true))
                            {
                                string content = reader.ReadToEnd()
                                    .Replace(relationshipTarget, renamedRelationshipTarget)
                                    .Replace("PartName=\"/" + sourcePath + "\"", "PartName=\"/" + renamedPath + "\"");
                                writer.Write(content);
                            }
                        }
                        else
                        {
                            source.CopyTo(target);
                        }
                    }
                }
            }
            Assert.True(sourceFound, "The source workbook did not contain the expected part '" + sourcePath + "'.");
            renamedStream.Position = 0;
            return renamedStream;
        }

        private static MemoryStream CreateWorkbookWithAdditionalRelationships(string additionalRelationships, params (string Path, string Content)[] additionalEntries)
        {
            using (MemoryStream originalStream = CreateWrittenWorkbookArchive())
            using (ZipArchive originalArchive = new ZipArchive(originalStream, ZipArchiveMode.Read, true))
            {
                MemoryStream resultStream = new MemoryStream();
                using (ZipArchive resultArchive = new ZipArchive(resultStream, ZipArchiveMode.Create, true))
                {
                    foreach (ZipArchiveEntry originalEntry in originalArchive.Entries)
                    {
                        ZipArchiveEntry resultEntry = resultArchive.CreateEntry(originalEntry.FullName);
                        using (Stream source = originalEntry.Open())
                        using (Stream target = resultEntry.Open())
                        {
                            if (originalEntry.FullName == "xl/_rels/workbook.xml.rels")
                            {
                                using (StreamReader reader = new StreamReader(source, Encoding.UTF8, true, 1024, true))
                                using (StreamWriter writer = new StreamWriter(target, new UTF8Encoding(false), 1024, true))
                                {
                                    string content = reader.ReadToEnd().Replace("</Relationships>", additionalRelationships + "</Relationships>");
                                    writer.Write(content);
                                }
                            }
                            else
                            {
                                source.CopyTo(target);
                            }
                        }
                    }
                    foreach ((string Path, string Content) entry in additionalEntries)
                    {
                        ZipArchiveEntry resultEntry = resultArchive.CreateEntry(entry.Path);
                        using (Stream target = resultEntry.Open())
                        {
                            byte[] content = Encoding.UTF8.GetBytes(entry.Content);
                            target.Write(content, 0, content.Length);
                        }
                    }
                }
                resultStream.Position = 0;
                return resultStream;
            }
        }

        private static string RenameWorkbookContent(string content)
        {
            return content
                .Replace("/xl/workbook.xml", "/custom/book7.xml")
                .Replace("xl/workbook.xml", "custom/book7.xml")
                .Replace("/xl/worksheets/sheet1.xml", "/parts/grid42.xml")
                .Replace("worksheets/sheet1.xml", "../parts/grid42.xml")
                .Replace("/xl/styles.xml", "/parts/design9.xml")
                .Replace("styles.xml", "../parts/design9.xml")
                .Replace("/xl/theme/theme1.xml", "/parts/look3.xml")
                .Replace("theme/theme1.xml", "../parts/look3.xml")
                .Replace("/xl/sharedStrings.xml", "/parts/strings8.xml")
                .Replace("sharedStrings.xml", "../parts/strings8.xml");
        }

        [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.DiscoveryReader, PlugInOrder = 100)]
        private class ReplacementDiscoveryReader : DiscoveryReader
        {
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = PLUGIN_ID, QueueUUID = PlugInUUID.ReaderPackageRegistryQueue, PlugInOrder = 100)]
        private class WorksheetDiscoveryPackageReader : IDiscoveryPackageReader
        {
            public const string PLUGIN_ID = "DISCOVERY_WORKSHEET_PACKAGE_TEST";
            private Stream previousStream;

            public string StreamEntryName { get { return null; } }
            public string DocumentType { get { return new WorksheetReader().DocumentType; } }
            public RelationshipInfo CurrentRelationship { get; set; }
            public Workbook Workbook { get; set; }
            public NanoXLSX.Interfaces.IOptions Options { get; set; }
            public Action<Stream, Workbook, string, NanoXLSX.Interfaces.IOptions, int?> InlinePluginHandler { get; set; }

            public void Init(Stream stream, Workbook workbook, NanoXLSX.Interfaces.IOptions options, Action<Stream, Workbook, string, NanoXLSX.Interfaces.IOptions, int?> inlinePluginHandler)
            {
                Assert.NotNull(stream);
                Assert.NotSame(previousStream, stream);
                previousStream = stream;
                Workbook = workbook;
                Options = options;
                InlinePluginHandler = inlinePluginHandler;
            }

            public void Execute()
            {
                int count = Workbook.AuxiliaryData.GetData<int>(PLUGIN_ID, 0);
                Workbook.AuxiliaryData.SetData(PLUGIN_ID, 0, count + 1, true);
                Workbook.AuxiliaryData.SetData(PLUGIN_ID, 1, CurrentRelationship.ResolvedTargetPath, true);
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = PLUGIN_ID, QueueUUID = PlugInUUID.ReaderPackageRegistryQueue, PlugInOrder = 10)]
        private class CatalogRegistryReader : IPluginQueueReader
        {
            public const string PLUGIN_ID = "DISCOVERY_CATALOG_ORDER_TEST";
            public Workbook Workbook { get; set; }
            public NanoXLSX.Interfaces.IOptions Options { get; set; }
            public Action<Stream, Workbook, string, NanoXLSX.Interfaces.IOptions, int?> InlinePluginHandler { get; set; }

            public void Init(Stream stream, Workbook workbook, NanoXLSX.Interfaces.IOptions options, Action<Stream, Workbook, string, NanoXLSX.Interfaces.IOptions, int?> inlinePluginHandler)
            {
                Workbook = workbook;
            }

            public void Execute()
            {
                Assert.NotNull(Workbook.AuxiliaryData.GetData<RelationshipCatalog>(PlugInUUID.DiscoveryReader, PlugInUUID.DiscoveryCatalogEntity));
                Workbook.AuxiliaryData.SetData(PLUGIN_ID, 0, "registry", true);
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "DISCOVERY_CATALOG_PREPENDING_TEST", QueueUUID = PlugInUUID.ReaderPrependingQueue, PlugInOrder = 10)]
        private class CatalogPrependingReader : IPluginQueueReader
        {
            public Workbook Workbook { get; set; }
            public NanoXLSX.Interfaces.IOptions Options { get; set; }
            public Action<Stream, Workbook, string, NanoXLSX.Interfaces.IOptions, int?> InlinePluginHandler { get; set; }

            public void Init(Stream stream, Workbook workbook, NanoXLSX.Interfaces.IOptions options, Action<Stream, Workbook, string, NanoXLSX.Interfaces.IOptions, int?> inlinePluginHandler)
            {
                Workbook = workbook;
            }

            public void Execute()
            {
                Assert.NotNull(Workbook.AuxiliaryData.GetData<RelationshipCatalog>(PlugInUUID.DiscoveryReader, PlugInUUID.DiscoveryCatalogEntity));
                Assert.Equal("registry", Workbook.AuxiliaryData.GetData<string>(CatalogRegistryReader.PLUGIN_ID, 0));
                Workbook.AuxiliaryData.SetData(CatalogRegistryReader.PLUGIN_ID, 1, "prepending", true);
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = PLUGIN_ID, QueueUUID = PlugInUUID.ReaderPackageRegistryQueue, PlugInOrder = 20)]
        private class ExactTypeDiscoveryPackageReader : IDiscoveryPackageReader
        {
            public const string PLUGIN_ID = "DISCOVERY_EXACT_TYPE_TEST";
            public string StreamEntryName { get { return null; } }
            public string DocumentType { get { return "http://example.test/relationships/custom"; } }
            public RelationshipInfo CurrentRelationship { get; set; }
            public Workbook Workbook { get; set; }
            public NanoXLSX.Interfaces.IOptions Options { get; set; }
            public Action<Stream, Workbook, string, NanoXLSX.Interfaces.IOptions, int?> InlinePluginHandler { get; set; }

            public void Init(Stream stream, Workbook workbook, NanoXLSX.Interfaces.IOptions options, Action<Stream, Workbook, string, NanoXLSX.Interfaces.IOptions, int?> inlinePluginHandler)
            {
                Assert.NotNull(stream);
                Workbook = workbook;
            }

            public void Execute()
            {
                int count = Workbook.AuxiliaryData.GetData<int>(PLUGIN_ID, 0);
                Workbook.AuxiliaryData.SetData(PLUGIN_ID, 0, count + 1, true);
                Workbook.AuxiliaryData.SetData(PLUGIN_ID, 1, CurrentRelationship.ResolvedTargetPath, true);
            }
        }

        [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.ThemeReader, PlugInOrder = 100)]
        private class ThemeSelectionReader : IDocumentReader
        {
            public const string PLUGIN_ID = "DISCOVERY_THEME_SELECTION_TEST";
            private Stream stream;

            public string DocumentType { get { return new ThemeReader().DocumentType; } }
            public Workbook Workbook { get; set; }
            public NanoXLSX.Interfaces.IOptions Options { get; set; }
            public Action<Stream, Workbook, string, NanoXLSX.Interfaces.IOptions, int?> InlinePluginHandler { get; set; }

            public void Init(Stream stream, Workbook workbook, NanoXLSX.Interfaces.IOptions options, Action<Stream, Workbook, string, NanoXLSX.Interfaces.IOptions, int?> inlinePluginHandler)
            {
                this.stream = stream;
                Workbook = workbook;
            }

            public void Execute()
            {
                string content;
                using (stream)
                using (StreamReader reader = new StreamReader(stream))
                {
                    content = reader.ReadToEnd();
                }
                int count = Workbook.AuxiliaryData.GetData<int>(PLUGIN_ID, 0);
                Workbook.AuxiliaryData.SetData(PLUGIN_ID, 0, count + 1, true);
                Workbook.AuxiliaryData.SetData(PLUGIN_ID, 1, content, true);
            }
        }
    }
}
