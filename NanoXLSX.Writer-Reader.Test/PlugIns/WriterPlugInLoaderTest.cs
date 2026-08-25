using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.IO.Compression;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using NanoXLSX.Interfaces.Reader;
using NanoXLSX.Interfaces.Writer;
using NanoXLSX.Internal.Structures;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;
using NanoXLSX.Utils.Xml;
using Xunit;
using IOException = NanoXLSX.Exceptions.IOException;

namespace NanoXLSX.Test.Writer_Reader.PlugInsTest
{
    // Ensure that these tests are executed sequentially, since static repository methods may be called
    [Collection(nameof(SequentialCollection2))]
    public class WriterPlugInLoaderTest : IDisposable
    {
        private const string ContentType = "application/vnd.openxmlformats-package.test-file+xml";
        private const string RelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tests";
        private const string SingleKey = "single-package";
        private const string RootKey = "root-package";
        private const string MultipleKey1 = "multiple-package-1";
        private const string MultipleKey2 = "multiple-package-2";
        private const string RelationshipSourceKey = "relationship-source";
        private const string SecondRelationshipSourceKey = "second-relationship-source";
        private const string RelationshipTargetKey = "relationship-target";
        private const string ExternalRelationshipId = "rIdExternal";
        private const string InternalRelationshipId = "rIdInternal";
        private const string ExternalRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath";
        private const string InternalRelationshipType = "http://example.org/relationships/internal";
        private const string ExternalRelationshipTarget = "file:///C:/temp/external.xlsx";
        private static readonly HashSet<Type> ActivePluginTypes = new HashSet<Type>();

        public void Dispose()
        {
            PlugInLoaderTestIsolation.Reset();
            SingleRegistry.Reset();
            SingleWriter.Reset();
            MultipleWriter.Reset();
            AppendingWriter.Reset();
            EmptyRegistry.Reset();
            EmptyWriter.Reset();
            NullableKeyWriter.Reset();
            ReaderInWriterQueue.Reset();
            WriterInPackageRegistryQueue.Reset();
            ActivePluginTypes.Clear();
        }

        [Fact(DisplayName = "Test of the plug-in handling initializer (dummy; should not crash)")]
        public void InitializeTest()
        {
            PlugInLoader.Initialize();
        }

        [Theory(DisplayName = "Separate registry and indexed writer plug-ins can write one package part")]
        [InlineData(typeof(SingleRegistry), typeof(SingleWriter), "xl/theme/test.xml", "single-package-0")]
        [InlineData(typeof(RootRegistry), typeof(RootWriter), "xl/rootTest.xml", "root-package-0")]
        public void SinglePackagePartTest(Type registryType, Type writerType, string expectedEntryPath, string expectedValue)
        {
            InjectPlugins(registryType, writerType);

            using (MemoryStream stream = SaveWorkbook())
            {
                AssertZipEntry(stream, expectedEntryPath, expectedValue);
            }
        }

        [Fact(DisplayName = "Registry and indexed writer plug-ins use an explicit initialization lifecycle")]
        public void PackagePluginLifecycleTest()
        {
            InjectPlugins(typeof(SingleRegistry), typeof(SingleWriter));

            using (MemoryStream stream = SaveWorkbook())
            {
                AssertZipEntry(stream, "xl/theme/test.xml", "single-package-0");
            }

            Assert.Equal(1, SingleRegistry.InitCount);
            Assert.Equal(1, SingleRegistry.ExecuteCount);
            Assert.Equal(1, SingleWriter.InitCount);
            Assert.Equal(1, SingleWriter.ExecuteCount);
            Assert.Equal(new[] { 0 }, SingleWriter.ExecutedIndexes);
        }

        [Fact(DisplayName = "Indexed writers correlate multiple package parts by key rather than position")]
        public void MultiplePackagePartsByKeyTest()
        {
            InjectPlugins(typeof(MultipleRegistry), typeof(MultipleWriter));

            using (MemoryStream stream = SaveWorkbook())
            {
                AssertZipEntry(stream, "xl/custom/multiple1.xml", "multiple-package-1");
                AssertZipEntry(stream, "xl/custom/multiple2.xml", "multiple-package-2");
            }

            Assert.Equal(new[] { 0, 1 }, MultipleWriter.ExecutedIndexes);
        }

        [Fact(DisplayName = "Multiple registry and indexed writer pairs can coexist")]
        public void MultiplePackagePluginPairsTest()
        {
            InjectPlugins(
                typeof(SingleRegistry),
                typeof(SingleWriter),
                typeof(MultipleRegistry),
                typeof(MultipleWriter));

            using (MemoryStream stream = SaveWorkbook())
            {
                AssertZipEntry(stream, "xl/theme/test.xml", "single-package-0");
                AssertZipEntry(stream, "xl/custom/multiple1.xml", "multiple-package-1");
                AssertZipEntry(stream, "xl/custom/multiple2.xml", "multiple-package-2");
            }
        }

        [Fact(DisplayName = "Registered package parts can own external and internal relationships")]
        public void PackagePartRelationshipsTest()
        {
            InjectPlugins(typeof(RelationshipRegistry), typeof(RelationshipWriter));

            using (MemoryStream stream = SaveWorkbook())
            {
                AssertZipEntry(stream, "xl/externalLinks/externalLink1.xml", ExternalRelationshipId);
                AssertZipEntry(stream, "xl/custom/relationshipTarget.xml", "relationship-target");

                XDocument sourceDocument = ReadZipXml(stream, "xl/externalLinks/externalLink1.xml");
                XNamespace officeRelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                Assert.Equal(
                    new[] { ExternalRelationshipId, InternalRelationshipId },
                    sourceDocument.Root.Elements("reference").Select(element => (string)element.Attribute(officeRelationshipNamespace + "id")));

                XDocument relationships = ReadZipXml(stream, "xl/externalLinks/_rels/externalLink1.xml.rels");
                XNamespace relationshipNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
                XElement externalRelationship = relationships.Root.Elements(relationshipNamespace + "Relationship")
                    .Single(element => (string)element.Attribute("Id") == ExternalRelationshipId);
                XElement internalRelationship = relationships.Root.Elements(relationshipNamespace + "Relationship")
                    .Single(element => (string)element.Attribute("Id") == InternalRelationshipId);

                Assert.Equal(ExternalRelationshipType, (string)externalRelationship.Attribute("Type"));
                Assert.Equal(ExternalRelationshipTarget, (string)externalRelationship.Attribute("Target"));
                Assert.Equal("External", (string)externalRelationship.Attribute("TargetMode"));
                Assert.Equal(InternalRelationshipType, (string)internalRelationship.Attribute("Type"));
                Assert.Equal("/xl/custom/relationshipTarget.xml", (string)internalRelationship.Attribute("Target"));
                Assert.Null(internalRelationship.Attribute("TargetMode"));

                XDocument secondRelationships = ReadZipXml(stream, "xl/custom/_rels/secondSource.xml.rels");
                XElement secondExternalRelationship = secondRelationships.Root.Elements(relationshipNamespace + "Relationship").Single();
                Assert.Equal(ExternalRelationshipId, (string)secondExternalRelationship.Attribute("Id"));
                Assert.Equal("https://example.org/second-target", (string)secondExternalRelationship.Attribute("Target"));

                AssertZipEntryMissing(stream, "xl/custom/_rels/relationshipTarget.xml.rels");
            }
        }

        [Fact(DisplayName = "Ordinary writers execute while unrelated plug-in types in the writer queue are ignored")]
        public void MixedAppendingQueueTest()
        {
            InjectPlugins(
                typeof(SingleRegistry),
                typeof(SingleWriter),
                typeof(ReaderInWriterQueue),
                typeof(AppendingWriter));

            using (MemoryStream stream = SaveWorkbook())
            {
                AssertZipEntry(stream, "xl/theme/test.xml", "single-package-0");
            }

            Assert.Equal(1, AppendingWriter.InitCount);
            Assert.Equal(1, AppendingWriter.ExecuteCount);
            Assert.Equal(0, ReaderInWriterQueue.ExecuteCount);
        }

        [Fact(DisplayName = "The package registry queue ignores plug-ins that are not package registries")]
        public void PackageRegistryQueueTypeFilterTest()
        {
            InjectPlugins(
                typeof(WriterInPackageRegistryQueue),
                typeof(SingleRegistry),
                typeof(SingleWriter));

            using (MemoryStream stream = SaveWorkbook())
            {
                AssertZipEntry(stream, "xl/theme/test.xml", "single-package-0");
            }

            Assert.Equal(0, WriterInPackageRegistryQueue.InitCount);
            Assert.Equal(0, WriterInPackageRegistryQueue.ExecuteCount);
        }

        [Fact(DisplayName = "Empty registries and indexed writers are valid no-ops")]
        public void EmptyPackagePluginsTest()
        {
            InjectPlugins(typeof(EmptyRegistry), typeof(EmptyWriter));

            using (MemoryStream stream = SaveWorkbook())
            {
                Assert.True(stream.Length > 0);
            }

            Assert.Equal(1, EmptyRegistry.InitCount);
            Assert.Equal(1, EmptyRegistry.ExecuteCount);
            Assert.Equal(1, EmptyWriter.InitCount);
            Assert.Equal(0, EmptyWriter.ExecuteCount);
        }

        [Fact(DisplayName = "An indexed iteration can explicitly select no package part")]
        public void NullPackagePartKeyTest()
        {
            InjectPlugins(typeof(NullableKeyWriter));

            using (MemoryStream stream = SaveWorkbook())
            {
                Assert.True(stream.Length > 0);
            }

            Assert.Equal(1, NullableKeyWriter.ExecuteCount);
        }

        [Theory(DisplayName = "Invalid package registry definitions fail with contextual errors")]
        [InlineData(typeof(InconsistentRegistry), "Inconsistent package registry plug-in detected")]
        [InlineData(typeof(NullCollectionRegistry), "Null collection in package registry plug-in")]
        [InlineData(typeof(BlankKeyRegistry), "Blank package part index")]
        [InlineData(typeof(DuplicateKeyRegistry), "Duplicate package part index")]
        [InlineData(typeof(InvalidDefinitionRegistry), "Invalid package part definition")]
        public void InvalidPackageRegistryTest(Type registryType, string expectedMessage)
        {
            InjectPlugins(registryType);

            using (MemoryStream stream = new MemoryStream())
            {
                IOException exception = Assert.Throws<IOException>(() => new Workbook().SaveAsStream(stream, true));
                Assert.Contains(expectedMessage, exception.Message);
                Assert.Contains(registryType.Name, exception.Message);
            }
        }

        [Theory(DisplayName = "Invalid package relationship definitions fail with contextual errors")]
        [InlineData(typeof(NullRelationshipCollectionRegistry), "Null collection in package registry plug-in")]
        [InlineData(typeof(InconsistentRelationshipCollectionRegistry), "Inconsistent package registry plug-in")]
        [InlineData(typeof(NullPackagePartRelationshipsRegistry), "Null package relationship collection")]
        [InlineData(typeof(NullPackageRelationshipRegistry), "Null package relationship")]
        [InlineData(typeof(BlankRelationshipIdRegistry), "Blank package relationship ID")]
        [InlineData(typeof(InvalidRelationshipIdRegistry), "Invalid package relationship ID")]
        [InlineData(typeof(DuplicateRelationshipIdRegistry), "Duplicate package relationship ID")]
        [InlineData(typeof(BlankRelationshipTypeRegistry), "Invalid package relationship type")]
        [InlineData(typeof(RelativeRelationshipTypeRegistry), "Invalid package relationship type")]
        [InlineData(typeof(BlankRelationshipTargetRegistry), "Invalid package relationship target")]
        [InlineData(typeof(InvalidRelationshipTargetRegistry), "Invalid package relationship target")]
        [InlineData(typeof(InvalidRelationshipTargetModeRegistry), "Invalid package relationship target mode")]
        [InlineData(typeof(AbsoluteInternalRelationshipTargetRegistry), "Absolute internal package relationship target")]
        [InlineData(typeof(NetworkPathInternalRelationshipTargetRegistry), "Network-path internal package relationship target")]
        public void InvalidPackageRelationshipTest(Type registryType, string expectedMessage)
        {
            InjectPlugins(registryType);

            using (MemoryStream stream = new MemoryStream())
            {
                IOException exception = Assert.Throws<IOException>(() => new Workbook().SaveAsStream(stream, true));
                Assert.Contains(expectedMessage, exception.Message);
                Assert.Contains(registryType.Name, exception.Message);
            }
        }

        [Theory(DisplayName = "Invalid indexed writer output fails with contextual errors")]
        [InlineData(typeof(UnknownKeyWriter), "Unknown package part index")]
        [InlineData(typeof(BlankKeyWriter), "Blank package part index")]
        [InlineData(typeof(MissingXmlWriter), "Missing XML element")]
        [InlineData(typeof(InvalidMaxIndexWriter), "Invalid maximum index")]
        public void InvalidIndexedWriterTest(Type writerType, string expectedMessage)
        {
            InjectPlugins(typeof(SingleRegistry), writerType);

            using (MemoryStream stream = new MemoryStream())
            {
                IOException exception = Assert.Throws<IOException>(() => new Workbook().SaveAsStream(stream, true));
                Assert.Contains(expectedMessage, exception.Message);
                Assert.Contains(writerType.Name, exception.Message);
            }
        }

        [Fact(DisplayName = "Indexed writers are rejected in the prepending queue")]
        public void IndexedWriterInPrependingQueueTest()
        {
            InjectPlugins(typeof(PrependingIndexedWriter));

            using (MemoryStream stream = new MemoryStream())
            {
                IOException exception = Assert.Throws<IOException>(() => new Workbook().SaveAsStream(stream, true));
                Assert.Contains("cannot be executed in the writer prepending queue", exception.Message);
                Assert.Contains(nameof(PrependingIndexedWriter), exception.Message);
            }
        }

        private static MemoryStream SaveWorkbook()
        {
            MemoryStream stream = new MemoryStream();
            new Workbook().SaveAsStream(stream, true);
            stream.Position = 0;
            return stream;
        }

        private static void InjectPlugins(params Type[] pluginTypes)
        {
            ActivePluginTypes.UnionWith(pluginTypes);
            PlugInLoader.InjectPlugins(new List<Type>(pluginTypes));
        }

        private static bool IsPluginActive(Type pluginType)
        {
            return ActivePluginTypes.Contains(pluginType);
        }

        private static void AssertZipEntry(Stream stream, string entryPath, string expectedValue)
        {
            stream.Position = 0;
            using (ZipArchive zip = new ZipArchive(stream, ZipArchiveMode.Read, true))
            {
                ZipArchiveEntry entry = zip.GetEntry(entryPath);
                Assert.NotNull(entry);
                using (StreamReader reader = new StreamReader(entry.Open()))
                {
                    Assert.Contains(expectedValue, reader.ReadToEnd());
                }
            }
        }

        private static XDocument ReadZipXml(Stream stream, string entryPath)
        {
            stream.Position = 0;
            using (ZipArchive zip = new ZipArchive(stream, ZipArchiveMode.Read, true))
            {
                ZipArchiveEntry entry = zip.GetEntry(entryPath);
                Assert.NotNull(entry);
                using (Stream entryStream = entry.Open())
                {
                    return XDocument.Load(entryStream);
                }
            }
        }

        private static void AssertZipEntryMissing(Stream stream, string entryPath)
        {
            stream.Position = 0;
            using (ZipArchive zip = new ZipArchive(stream, ZipArchiveMode.Read, true))
            {
                Assert.Null(zip.GetEntry(entryPath));
            }
        }

        internal abstract class PackageRegistryBase : IPluginPackageRegistry
        {
            protected PackageRegistryBase()
            {
                OrderNumberList = new List<int>();
                PackagePartPathList = new List<string>();
                PackagePartFileNameList = new List<string>();
                ContentTypeList = new List<string>();
                RelationshipTypeList = new List<string>();
                ArePackagePartRootList = new List<bool>();
                UniquePackagePartIndexList = new List<string>();
                PackagePartRelationshipList = new List<IReadOnlyList<IPluginPackageRelationship>>();
            }

            protected List<int> OrderNumberList { get; }
            protected List<string> PackagePartPathList { get; }
            protected List<string> PackagePartFileNameList { get; }
            protected List<string> ContentTypeList { get; }
            protected List<string> RelationshipTypeList { get; }
            protected List<bool> ArePackagePartRootList { get; }
            protected List<string> UniquePackagePartIndexList { get; }
            protected List<IReadOnlyList<IPluginPackageRelationship>> PackagePartRelationshipList { get; }
            protected bool IsActive => IsPluginActive(GetType());

            public virtual IReadOnlyList<int> OrderNumbers => OrderNumberList;
            public virtual IReadOnlyList<string> PackagePartPaths => PackagePartPathList;
            public virtual IReadOnlyList<string> PackagePartFileNames => PackagePartFileNameList;
            public virtual IReadOnlyList<string> ContentTypes => ContentTypeList;
            public virtual IReadOnlyList<string> RelationshipTypes => RelationshipTypeList;
            public virtual IReadOnlyList<bool> ArePackagePartsRoot => ArePackagePartRootList;
            public virtual IReadOnlyList<string> UniquePackagePartIndices => UniquePackagePartIndexList;
            public virtual IReadOnlyList<IReadOnlyList<IPluginPackageRelationship>> PackagePartRelationships => PackagePartRelationshipList;
            public Workbook Workbook { get; set; }

            public virtual void Init(IBaseWriter baseWriter)
            {
                Workbook = baseWriter.Workbook;
            }

            public virtual void Execute()
            {
            }

            protected void AddDefinition(int orderNumber, string path, string fileName, string uniqueIndex, bool isRoot = false, IReadOnlyList<IPluginPackageRelationship> relationships = null)
            {
                OrderNumberList.Add(orderNumber);
                PackagePartPathList.Add(path);
                PackagePartFileNameList.Add(fileName);
                ContentTypeList.Add(ContentType);
                RelationshipTypeList.Add(RelationshipType);
                ArePackagePartRootList.Add(isRoot);
                UniquePackagePartIndexList.Add(uniqueIndex);
                PackagePartRelationshipList.Add(relationships ?? Array.Empty<IPluginPackageRelationship>());
            }
        }

        internal sealed class PluginPackageRelationship : IPluginPackageRelationship
        {
            public string RelationshipId { get; }
            public string RelationshipType { get; }
            public string Target { get; }
            public TargetMode TargetMode { get; }

            internal PluginPackageRelationship(string relationshipId, string relationshipType, string target, TargetMode targetMode)
            {
                RelationshipId = relationshipId;
                RelationshipType = relationshipType;
                Target = target;
                TargetMode = targetMode;
            }
        }

        internal abstract class IndexedWriterBase : IPluginIndexedWriter
        {
            private readonly IReadOnlyList<string> packagePartIndices;
            private readonly IReadOnlyList<string> values;
            private XmlElement xmlElement;

            protected IndexedWriterBase(IReadOnlyList<string> packagePartIndices, IReadOnlyList<string> values)
            {
                this.packagePartIndices = packagePartIndices;
                this.values = values;
                CurrentIndex = -1;
            }

            public int CurrentIndex { get; set; }
            public virtual string CurrentUniquePackagePartIndex => packagePartIndices[CurrentIndex];
            public virtual int MaxIndex => IsPluginActive(GetType()) ? packagePartIndices.Count - 1 : -1;
            public Workbook Workbook { get; set; }
            public XmlElement XmlElement => xmlElement;

            public virtual void Init(IBaseWriter baseWriter)
            {
                Workbook = baseWriter.Workbook;
            }

            public virtual void Execute()
            {
                xmlElement = CreateElement(values[CurrentIndex]);
            }

            protected static XmlElement CreateElement(string value)
            {
                XmlElement element = XmlElement.CreateElement("test");
                element.InnerValue = value;
                return element;
            }

            protected void SetXmlElement(XmlElement element)
            {
                xmlElement = element;
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_SINGLE_REGISTRY", QueueUUID = PlugInUUID.WriterPackageRegistryQueue)]
        internal class SingleRegistry : PackageRegistryBase
        {
            public static int InitCount { get; private set; }
            public static int ExecuteCount { get; private set; }

            public override void Init(IBaseWriter baseWriter)
            {
                InitCount++;
                base.Init(baseWriter);
            }

            public override void Execute()
            {
                ExecuteCount++;
                if (IsActive)
                {
                    AddDefinition(PackagePartDefinition.POST_WORkSHEET_PACKAGE_PART_START_INDEX + 1, "xl/theme/", "test.xml", SingleKey);
                }
            }

            public static void Reset()
            {
                InitCount = 0;
                ExecuteCount = 0;
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_SINGLE_WRITER", QueueUUID = PlugInUUID.WriterAppendingQueue, PlugInOrder = 10)]
        internal class SingleWriter : IndexedWriterBase
        {
            public static readonly List<int> ExecutedIndexes = new List<int>();
            public static int InitCount { get; private set; }
            public static int ExecuteCount { get; private set; }

            public SingleWriter() : base(new[] { SingleKey }, new[] { "single-package-0" })
            {
            }

            public override void Init(IBaseWriter baseWriter)
            {
                InitCount++;
                base.Init(baseWriter);
            }

            public override void Execute()
            {
                ExecuteCount++;
                ExecutedIndexes.Add(CurrentIndex);
                base.Execute();
            }

            public static void Reset()
            {
                InitCount = 0;
                ExecuteCount = 0;
                ExecutedIndexes.Clear();
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_ROOT_REGISTRY", QueueUUID = PlugInUUID.WriterPackageRegistryQueue)]
        internal class RootRegistry : PackageRegistryBase
        {
            public RootRegistry()
            {
                if (IsActive)
                {
                    AddDefinition(99, "xl/", "rootTest.xml", RootKey, true);
                }
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_ROOT_WRITER", QueueUUID = PlugInUUID.WriterAppendingQueue)]
        internal class RootWriter : IndexedWriterBase
        {
            public RootWriter() : base(new[] { RootKey }, new[] { "root-package-0" })
            {
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_MULTIPLE_REGISTRY", QueueUUID = PlugInUUID.WriterPackageRegistryQueue, PlugInOrder = 10)]
        internal class MultipleRegistry : PackageRegistryBase
        {
            public MultipleRegistry()
            {
                if (IsActive)
                {
                    AddDefinition(PackagePartDefinition.POST_WORkSHEET_PACKAGE_PART_START_INDEX + 2, "xl/custom/", "multiple1.xml", MultipleKey1);
                    AddDefinition(PackagePartDefinition.POST_WORkSHEET_PACKAGE_PART_START_INDEX + 3, "xl/custom/", "multiple2.xml", MultipleKey2);
                }
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_MULTIPLE_WRITER", QueueUUID = PlugInUUID.WriterAppendingQueue, PlugInOrder = 20)]
        internal class MultipleWriter : IndexedWriterBase
        {
            public static readonly List<int> ExecutedIndexes = new List<int>();

            public MultipleWriter() : base(
                new[] { MultipleKey2, MultipleKey1 },
                new[] { "multiple-package-2", "multiple-package-1" })
            {
            }

            public override void Execute()
            {
                ExecutedIndexes.Add(CurrentIndex);
                base.Execute();
            }

            public static void Reset()
            {
                ExecutedIndexes.Clear();
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_RELATIONSHIP_REGISTRY", QueueUUID = PlugInUUID.WriterPackageRegistryQueue)]
        internal class RelationshipRegistry : PackageRegistryBase
        {
            public override void Execute()
            {
                AddDefinition(
                    PackagePartDefinition.POST_WORkSHEET_PACKAGE_PART_START_INDEX + 4,
                    "xl/externalLinks/",
                    "externalLink1.xml",
                    RelationshipSourceKey,
                    false,
                    new IPluginPackageRelationship[]
                    {
                        new PluginPackageRelationship(ExternalRelationshipId, ExternalRelationshipType, ExternalRelationshipTarget, TargetMode.External),
                        new PluginPackageRelationship(InternalRelationshipId, InternalRelationshipType, "/xl/custom/relationshipTarget.xml", TargetMode.Internal)
                    });
                AddDefinition(
                    PackagePartDefinition.POST_WORkSHEET_PACKAGE_PART_START_INDEX + 5,
                    "xl/custom/",
                    "secondSource.xml",
                    SecondRelationshipSourceKey,
                    false,
                    new IPluginPackageRelationship[]
                    {
                        new PluginPackageRelationship(ExternalRelationshipId, ExternalRelationshipType, "https://example.org/second-target", TargetMode.External)
                    });
                AddDefinition(
                    PackagePartDefinition.POST_WORkSHEET_PACKAGE_PART_START_INDEX + 6,
                    "xl/custom/",
                    "relationshipTarget.xml",
                    RelationshipTargetKey);
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_RELATIONSHIP_WRITER", QueueUUID = PlugInUUID.WriterAppendingQueue)]
        internal class RelationshipWriter : IndexedWriterBase
        {
            public RelationshipWriter() : base(
                new[] { RelationshipSourceKey, SecondRelationshipSourceKey, RelationshipTargetKey },
                new[]
                {
                    ExternalRelationshipId + " " + InternalRelationshipId,
                    ExternalRelationshipId,
                    "relationship-target"
                })
            {
            }

            public override void Execute()
            {
                if (CurrentIndex != 0)
                {
                    base.Execute();
                    return;
                }

                XmlElement element = XmlElement.CreateElement("externalLink");
                element.AddNameSpaceAttribute("r", "xmlns", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                element.AddChildElementWithAttribute("reference", "id", ExternalRelationshipId, "", "r");
                element.AddChildElementWithAttribute("reference", "id", InternalRelationshipId, "", "r");
                SetXmlElement(element);
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_EMPTY_REGISTRY", QueueUUID = PlugInUUID.WriterPackageRegistryQueue)]
        internal class EmptyRegistry : PackageRegistryBase
        {
            public static int InitCount { get; private set; }
            public static int ExecuteCount { get; private set; }

            public override void Init(IBaseWriter baseWriter)
            {
                InitCount++;
                base.Init(baseWriter);
            }

            public override void Execute()
            {
                ExecuteCount++;
            }

            public static void Reset()
            {
                InitCount = 0;
                ExecuteCount = 0;
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_EMPTY_WRITER", QueueUUID = PlugInUUID.WriterAppendingQueue)]
        internal class EmptyWriter : IndexedWriterBase
        {
            public static int InitCount { get; private set; }
            public static int ExecuteCount { get; private set; }

            public EmptyWriter() : base(Array.Empty<string>(), Array.Empty<string>())
            {
            }

            public override void Init(IBaseWriter baseWriter)
            {
                InitCount++;
                base.Init(baseWriter);
            }

            [ExcludeFromCodeCoverage]
            public override void Execute()
            {
                ExecuteCount++;
            }

            public static void Reset()
            {
                InitCount = 0;
                ExecuteCount = 0;
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_NULLABLE_KEY_WRITER", QueueUUID = PlugInUUID.WriterAppendingQueue)]
        internal class NullableKeyWriter : IndexedWriterBase
        {
            public static int ExecuteCount { get; private set; }

            public NullableKeyWriter() : base(new string[] { null }, new[] { "ignored" })
            {
            }

            public override void Execute()
            {
                ExecuteCount++;
                base.Execute();
            }

            public static void Reset()
            {
                ExecuteCount = 0;
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_INCONSISTENT_REGISTRY", QueueUUID = PlugInUUID.WriterPackageRegistryQueue)]
        internal class InconsistentRegistry : PackageRegistryBase
        {
            public InconsistentRegistry()
            {
                if (IsActive)
                {
                    AddDefinition(1, "xl/", "inconsistent.xml", "inconsistent");
                    PackagePartFileNameList.Clear();
                }
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_NULL_COLLECTION_REGISTRY", QueueUUID = PlugInUUID.WriterPackageRegistryQueue)]
        internal class NullCollectionRegistry : PackageRegistryBase
        {
            public override IReadOnlyList<string> ContentTypes => IsActive ? null : base.ContentTypes;
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_BLANK_KEY_REGISTRY", QueueUUID = PlugInUUID.WriterPackageRegistryQueue)]
        internal class BlankKeyRegistry : PackageRegistryBase
        {
            public BlankKeyRegistry()
            {
                if (IsActive)
                {
                    AddDefinition(1, "xl/", "blank.xml", " ");
                }
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_DUPLICATE_KEY_REGISTRY", QueueUUID = PlugInUUID.WriterPackageRegistryQueue)]
        internal class DuplicateKeyRegistry : PackageRegistryBase
        {
            public DuplicateKeyRegistry()
            {
                if (IsActive)
                {
                    AddDefinition(1, "xl/custom/", "duplicate1.xml", "duplicate");
                    AddDefinition(2, "xl/custom/", "duplicate2.xml", "duplicate");
                }
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_INVALID_DEFINITION_REGISTRY", QueueUUID = PlugInUUID.WriterPackageRegistryQueue)]
        internal class InvalidDefinitionRegistry : PackageRegistryBase
        {
            public InvalidDefinitionRegistry()
            {
                if (IsActive)
                {
                    AddDefinition(1, " ", "invalid.xml", "invalid-definition");
                }
            }
        }

        internal abstract class InvalidRelationshipRegistryBase : PackageRegistryBase
        {
            protected InvalidRelationshipRegistryBase(params IPluginPackageRelationship[] relationships)
            {
                if (IsActive)
                {
                    AddDefinition(1, "xl/custom/", "invalidRelationship.xml", "invalid-relationship", false, relationships);
                }
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_NULL_RELATIONSHIP_COLLECTION_REGISTRY", QueueUUID = PlugInUUID.WriterPackageRegistryQueue)]
        internal class NullRelationshipCollectionRegistry : PackageRegistryBase
        {
            public override IReadOnlyList<IReadOnlyList<IPluginPackageRelationship>> PackagePartRelationships => IsActive ? null : base.PackagePartRelationships;
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_INCONSISTENT_RELATIONSHIP_COLLECTION_REGISTRY", QueueUUID = PlugInUUID.WriterPackageRegistryQueue)]
        internal class InconsistentRelationshipCollectionRegistry : PackageRegistryBase
        {
            public InconsistentRelationshipCollectionRegistry()
            {
                if (IsActive)
                {
                    AddDefinition(1, "xl/custom/", "inconsistentRelationship.xml", "inconsistent-relationship");
                    PackagePartRelationshipList.Clear();
                }
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_NULL_PACKAGE_PART_RELATIONSHIPS_REGISTRY", QueueUUID = PlugInUUID.WriterPackageRegistryQueue)]
        internal class NullPackagePartRelationshipsRegistry : PackageRegistryBase
        {
            public NullPackagePartRelationshipsRegistry()
            {
                if (IsActive)
                {
                    AddDefinition(1, "xl/custom/", "nullRelationships.xml", "null-relationships");
                    PackagePartRelationshipList[0] = null;
                }
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_NULL_PACKAGE_RELATIONSHIP_REGISTRY", QueueUUID = PlugInUUID.WriterPackageRegistryQueue)]
        internal class NullPackageRelationshipRegistry : InvalidRelationshipRegistryBase
        {
            public NullPackageRelationshipRegistry() : base((IPluginPackageRelationship)null)
            {
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_BLANK_RELATIONSHIP_ID_REGISTRY", QueueUUID = PlugInUUID.WriterPackageRegistryQueue)]
        internal class BlankRelationshipIdRegistry : InvalidRelationshipRegistryBase
        {
            public BlankRelationshipIdRegistry() : base(new PluginPackageRelationship(" ", ExternalRelationshipType, ExternalRelationshipTarget, TargetMode.External))
            {
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_INVALID_RELATIONSHIP_ID_REGISTRY", QueueUUID = PlugInUUID.WriterPackageRegistryQueue)]
        internal class InvalidRelationshipIdRegistry : InvalidRelationshipRegistryBase
        {
            public InvalidRelationshipIdRegistry() : base(new PluginPackageRelationship("1 invalid", ExternalRelationshipType, ExternalRelationshipTarget, TargetMode.External))
            {
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_DUPLICATE_RELATIONSHIP_ID_REGISTRY", QueueUUID = PlugInUUID.WriterPackageRegistryQueue)]
        internal class DuplicateRelationshipIdRegistry : InvalidRelationshipRegistryBase
        {
            public DuplicateRelationshipIdRegistry() : base(
                new PluginPackageRelationship("rId1", ExternalRelationshipType, ExternalRelationshipTarget, TargetMode.External),
                new PluginPackageRelationship("rId1", ExternalRelationshipType, "https://example.org/duplicate", TargetMode.External))
            {
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_BLANK_RELATIONSHIP_TYPE_REGISTRY", QueueUUID = PlugInUUID.WriterPackageRegistryQueue)]
        internal class BlankRelationshipTypeRegistry : InvalidRelationshipRegistryBase
        {
            public BlankRelationshipTypeRegistry() : base(new PluginPackageRelationship("rId1", " ", ExternalRelationshipTarget, TargetMode.External))
            {
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_RELATIVE_RELATIONSHIP_TYPE_REGISTRY", QueueUUID = PlugInUUID.WriterPackageRegistryQueue)]
        internal class RelativeRelationshipTypeRegistry : InvalidRelationshipRegistryBase
        {
            public RelativeRelationshipTypeRegistry() : base(new PluginPackageRelationship("rId1", "relative/type", ExternalRelationshipTarget, TargetMode.External))
            {
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_BLANK_RELATIONSHIP_TARGET_REGISTRY", QueueUUID = PlugInUUID.WriterPackageRegistryQueue)]
        internal class BlankRelationshipTargetRegistry : InvalidRelationshipRegistryBase
        {
            public BlankRelationshipTargetRegistry() : base(new PluginPackageRelationship("rId1", ExternalRelationshipType, " ", TargetMode.External))
            {
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_INVALID_RELATIONSHIP_TARGET_REGISTRY", QueueUUID = PlugInUUID.WriterPackageRegistryQueue)]
        internal class InvalidRelationshipTargetRegistry : InvalidRelationshipRegistryBase
        {
            public InvalidRelationshipTargetRegistry() : base(new PluginPackageRelationship("rId1", ExternalRelationshipType, "http://[invalid", TargetMode.External))
            {
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_INVALID_RELATIONSHIP_TARGET_MODE_REGISTRY", QueueUUID = PlugInUUID.WriterPackageRegistryQueue)]
        internal class InvalidRelationshipTargetModeRegistry : InvalidRelationshipRegistryBase
        {
            public InvalidRelationshipTargetModeRegistry() : base(new PluginPackageRelationship("rId1", ExternalRelationshipType, ExternalRelationshipTarget, (TargetMode)99))
            {
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_ABSOLUTE_INTERNAL_RELATIONSHIP_TARGET_REGISTRY", QueueUUID = PlugInUUID.WriterPackageRegistryQueue)]
        internal class AbsoluteInternalRelationshipTargetRegistry : InvalidRelationshipRegistryBase
        {
            public AbsoluteInternalRelationshipTargetRegistry() : base(new PluginPackageRelationship("rId1", InternalRelationshipType, "https://example.org/internal", TargetMode.Internal))
            {
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_NETWORK_PATH_INTERNAL_RELATIONSHIP_TARGET_REGISTRY", QueueUUID = PlugInUUID.WriterPackageRegistryQueue)]
        internal class NetworkPathInternalRelationshipTargetRegistry : InvalidRelationshipRegistryBase
        {
            public NetworkPathInternalRelationshipTargetRegistry() : base(new PluginPackageRelationship("rId1", InternalRelationshipType, "//example.org/internal", TargetMode.Internal))
            {
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_UNKNOWN_KEY_WRITER", QueueUUID = PlugInUUID.WriterAppendingQueue)]
        internal class UnknownKeyWriter : IndexedWriterBase
        {
            public UnknownKeyWriter() : base(new[] { "unknown" }, new[] { "unknown" })
            {
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_BLANK_KEY_WRITER", QueueUUID = PlugInUUID.WriterAppendingQueue)]
        internal class BlankKeyWriter : IndexedWriterBase
        {
            public BlankKeyWriter() : base(new[] { " " }, new[] { "blank" })
            {
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_MISSING_XML_WRITER", QueueUUID = PlugInUUID.WriterAppendingQueue)]
        internal class MissingXmlWriter : IndexedWriterBase
        {
            public MissingXmlWriter() : base(new[] { SingleKey }, new[] { "missing" })
            {
            }

            public override void Execute()
            {
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_INVALID_MAX_INDEX_WRITER", QueueUUID = PlugInUUID.WriterAppendingQueue)]
        internal class InvalidMaxIndexWriter : IndexedWriterBase
        {
            public InvalidMaxIndexWriter() : base(Array.Empty<string>(), Array.Empty<string>())
            {
            }

            public override int MaxIndex => IsPluginActive(GetType()) ? -2 : -1;
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_PREPENDING_INDEXED_WRITER", QueueUUID = PlugInUUID.WriterPrependingQueue)]
        internal class PrependingIndexedWriter : IndexedWriterBase
        {
            public PrependingIndexedWriter() : base(new[] { SingleKey }, new[] { "prepending" })
            {
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_ORDINARY_APPEND", QueueUUID = PlugInUUID.WriterAppendingQueue, PlugInOrder = 30)]
        internal class AppendingWriter : IPluginWriter
        {
            public static int InitCount { get; private set; }
            public static int ExecuteCount { get; private set; }
            public Workbook Workbook { get; set; }
            [ExcludeFromCodeCoverage]
            public XmlElement XmlElement => null;

            public void Execute()
            {
                ExecuteCount++;
            }

            public void Init(IBaseWriter baseWriter)
            {
                InitCount++;
                Workbook = baseWriter.Workbook;
            }

            public static void Reset()
            {
                InitCount = 0;
                ExecuteCount = 0;
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_READER_IN_WRITER_QUEUE", QueueUUID = PlugInUUID.WriterAppendingQueue, PlugInOrder = 25)]
        internal class ReaderInWriterQueue : IPluginReader
        {
            public static int ExecuteCount { get; private set; }
            [ExcludeFromCodeCoverage]
            public Workbook Workbook { get; set; }

            [ExcludeFromCodeCoverage]
            public void Execute()
            {
                ExecuteCount++;
            }

            public static void Reset()
            {
                ExecuteCount = 0;
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_WRITER_IN_PACKAGE_REGISTRY_QUEUE", QueueUUID = PlugInUUID.WriterPackageRegistryQueue, PlugInOrder = -10)]
        internal class WriterInPackageRegistryQueue : IPluginWriter
        {
            public static int InitCount { get; private set; }
            public static int ExecuteCount { get; private set; }
            [ExcludeFromCodeCoverage]
            public Workbook Workbook { get; set; }
            [ExcludeFromCodeCoverage]
            public XmlElement XmlElement => null;

            [ExcludeFromCodeCoverage]
            public void Execute()
            {
                ExecuteCount++;
            }

            [ExcludeFromCodeCoverage]
            public void Init(IBaseWriter baseWriter)
            {
                InitCount++;
                Workbook = baseWriter.Workbook;
            }

            public static void Reset()
            {
                InitCount = 0;
                ExecuteCount = 0;
            }
        }
    }
}
