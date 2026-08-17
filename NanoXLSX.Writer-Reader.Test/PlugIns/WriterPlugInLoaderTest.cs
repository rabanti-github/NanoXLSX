using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.IO.Compression;
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
            }

            protected List<int> OrderNumberList { get; }
            protected List<string> PackagePartPathList { get; }
            protected List<string> PackagePartFileNameList { get; }
            protected List<string> ContentTypeList { get; }
            protected List<string> RelationshipTypeList { get; }
            protected List<bool> ArePackagePartRootList { get; }
            protected List<string> UniquePackagePartIndexList { get; }
            protected bool IsActive => IsPluginActive(GetType());

            public virtual IReadOnlyList<int> OrderNumbers => OrderNumberList;
            public virtual IReadOnlyList<string> PackagePartPaths => PackagePartPathList;
            public virtual IReadOnlyList<string> PackagePartFileNames => PackagePartFileNameList;
            public virtual IReadOnlyList<string> ContentTypes => ContentTypeList;
            public virtual IReadOnlyList<string> RelationshipTypes => RelationshipTypeList;
            public virtual IReadOnlyList<bool> ArePackagePartsRoot => ArePackagePartRootList;
            public virtual IReadOnlyList<string> UniquePackagePartIndices => UniquePackagePartIndexList;
            public Workbook Workbook { get; set; }

            public virtual void Init(IBaseWriter baseWriter)
            {
                Workbook = baseWriter.Workbook;
            }

            public virtual void Execute()
            {
            }

            protected void AddDefinition(int orderNumber, string path, string fileName, string uniqueIndex, bool isRoot = false)
            {
                OrderNumberList.Add(orderNumber);
                PackagePartPathList.Add(path);
                PackagePartFileNameList.Add(fileName);
                ContentTypeList.Add(ContentType);
                RelationshipTypeList.Add(RelationshipType);
                ArePackagePartRootList.Add(isRoot);
                UniquePackagePartIndexList.Add(uniqueIndex);
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
