using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.IO.Compression;
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

        public void Dispose()
        {
            PlugInLoader.DisposePlugins();
            SinglePackage.Reset();
            MultiplePackage.Reset();
            AppendingWriter.Reset();
        }

        [Fact(DisplayName = "Test of the plug-in handling initializer (dummy; should not crash)")]
        public void InitializeTest()
        {
            PlugInLoader.Initialize();
        }

        [Theory(DisplayName = "Package writer validation rejects inconsistent definitions in both queues")]
        [InlineData(typeof(InconsistentRegistryPackage), "InconsistentRegistryPackage")]
        [InlineData(typeof(InconsistentAppendingPackage), "InconsistentAppendingPackage")]
        public void InconsistentPackageWriterTest(Type pluginType, string expectedPluginName)
        {
            PlugInLoader.InjectPlugins(new List<Type> { pluginType });

            using (MemoryStream stream = new MemoryStream())
            {
                IOException exception = Assert.Throws<IOException>(() => new Workbook().SaveAsStream(stream, true));
                Assert.Contains("Inconsistent package writer plug-in detected: " + expectedPluginName, exception.Message);
            }
        }

        [Theory(DisplayName = "A package writer can register and write one package part")]
        [InlineData(typeof(SinglePackage), "xl/theme/test.xml", "single-package-0")]
        [InlineData(typeof(SingleRootPackage), "xl/rootTest.xml", "single-root-package-0")]
        public void SinglePackagePartTest(Type pluginType, string expectedEntryPath, string expectedValue)
        {
            PlugInLoader.InjectPlugins(new List<Type> { pluginType });

            using (MemoryStream stream = SaveWorkbook())
            {
                AssertZipEntry(stream, expectedEntryPath, expectedValue);
            }
        }

        [Fact(DisplayName = "A package writer can register and write multiple package parts")]
        public void MultiplePackagePartsTest()
        {
            MultiplePackage.Reset();
            PlugInLoader.InjectPlugins(new List<Type> { typeof(MultiplePackage) });

            using (MemoryStream stream = SaveWorkbook())
            {
                AssertZipEntry(stream, "xl/custom/multiple1.xml", "multiple-package-0");
                AssertZipEntry(stream, "xl/custom/multiple2.xml", "multiple-package-1");
            }

            Assert.Equal(new[] { 0, 1 }, MultiplePackage.ExecutedIndexes);
        }

        [Fact(DisplayName = "Multiple package writers can each write one or multiple package parts")]
        public void MultiplePackageWritersTest()
        {
            SinglePackage.Reset();
            MultiplePackage.Reset();
            PlugInLoader.InjectPlugins(new List<Type> { typeof(SinglePackage), typeof(MultiplePackage) });

            using (MemoryStream stream = SaveWorkbook())
            {
                AssertZipEntry(stream, "xl/theme/test.xml", "single-package-0");
                AssertZipEntry(stream, "xl/custom/multiple1.xml", "multiple-package-0");
                AssertZipEntry(stream, "xl/custom/multiple2.xml", "multiple-package-1");
            }

            Assert.Equal(new[] { 0 }, SinglePackage.ExecutedIndexes);
            Assert.Equal(new[] { 0, 1 }, MultiplePackage.ExecutedIndexes);
        }

        [Fact(DisplayName = "The appending queue executes ordinary writers but does not re-execute package writers")]
        public void MixedAppendingQueueTest()
        {
            SinglePackage.Reset();
            AppendingWriter.Reset();
            PlugInLoader.InjectPlugins(new List<Type> { typeof(SinglePackage), typeof(AppendingWriter) });

            using (MemoryStream stream = SaveWorkbook())
            {
                AssertZipEntry(stream, "xl/theme/test.xml", "single-package-0");
            }

            // The package writer executes once in the registry queue. It must not execute again in the appending queue.
            Assert.Equal(1, SinglePackage.ExecuteCount);
            Assert.Equal(1, AppendingWriter.InitCount);
            Assert.Equal(1, AppendingWriter.ExecuteCount);
        }

        private static MemoryStream SaveWorkbook()
        {
            MemoryStream stream = new MemoryStream();
            new Workbook().SaveAsStream(stream, true);
            stream.Position = 0;
            return stream;
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

        internal abstract class PackageWriterBase : IPluginPackageWriter
        {
            private readonly List<XmlElement> xmlElements;

            protected PackageWriterBase(IReadOnlyList<string> elementPrefixes)
            {
                xmlElements = new List<XmlElement>();
                for (int i = 0; i < elementPrefixes.Count; i++)
                {
                    xmlElements.Add(CreateElement(elementPrefixes[i] + "-" + i));
                }
            }

            public abstract List<int> OrderNumbers { get; }
            public int CurrentIndex { get; set; } = -1;
            public abstract List<string> PackagePartPaths { get; }
            public abstract List<string> PackagePartFileNames { get; }
            public abstract List<string> ContentTypes { get; }
            public abstract List<string> RelationshipTypes { get; }
            public abstract List<bool> ArePackagePartsRoot { get; }
            public List<XmlElement> XmlElements => xmlElements;

            [ExcludeFromCodeCoverage]
            public Workbook Workbook { get; set; }

            // IPluginPackageWriter.XmlElements is used instead of this property.
            [ExcludeFromCodeCoverage]
            public XmlElement XmlElement => CreateElement("ignored");

            public virtual void Execute()
            {
                XmlElements[CurrentIndex] = CreateElement(ElementPrefix(CurrentIndex) + "-" + CurrentIndex);
            }

            [ExcludeFromCodeCoverage]
            void IPluginWriter.Init(IBaseWriter baseWriter)
            {
                Workbook = baseWriter.Workbook;
            }

            protected abstract string ElementPrefix(int index);

            protected static XmlElement CreateElement(string value)
            {
                XmlElement element = XmlElement.CreateElement("test");
                element.InnerValue = value;
                return element;
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_SINGLE_REGISTRY", QueueUUID = PlugInUUID.WriterPackageRegistryQueue)]
        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_SINGLE_APPEND", QueueUUID = PlugInUUID.WriterAppendingQueue, PlugInOrder = 10)]
        internal class SinglePackage : PackageWriterBase
        {
            public static readonly List<int> ExecutedIndexes = new List<int>();
            public static int ExecuteCount { get; private set; }

            public SinglePackage() : base(new[] { "single-package" }) { }

            public override List<int> OrderNumbers => new List<int> { PackagePartDefinition.POST_WORkSHEET_PACKAGE_PART_START_INDEX + 1 };
            public override List<string> PackagePartPaths => new List<string> { "xl/theme/" };
            public override List<string> PackagePartFileNames => new List<string> { "test.xml" };
            public override List<string> ContentTypes => new List<string> { ContentType };
            public override List<string> RelationshipTypes => new List<string> { RelationshipType };
            public override List<bool> ArePackagePartsRoot => new List<bool> { false };

            public override void Execute()
            {
                ExecuteCount++;
                ExecutedIndexes.Add(CurrentIndex);
                base.Execute();
            }

            protected override string ElementPrefix(int index) => "single-package";

            public static void Reset()
            {
                ExecuteCount = 0;
                ExecutedIndexes.Clear();
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_ROOT_REGISTRY", QueueUUID = PlugInUUID.WriterPackageRegistryQueue)]
        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_ROOT_APPEND", QueueUUID = PlugInUUID.WriterAppendingQueue)]
        internal class SingleRootPackage : PackageWriterBase
        {
            public SingleRootPackage() : base(new[] { "single-root-package" }) { }

            public override List<int> OrderNumbers => new List<int> { 99 };
            public override List<string> PackagePartPaths => new List<string> { "xl/" };
            public override List<string> PackagePartFileNames => new List<string> { "rootTest.xml" };
            public override List<string> ContentTypes => new List<string> { ContentType };
            public override List<string> RelationshipTypes => new List<string> { RelationshipType };
            public override List<bool> ArePackagePartsRoot => new List<bool> { true };
            protected override string ElementPrefix(int index) => "single-root-package";
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_MULTIPLE_REGISTRY", QueueUUID = PlugInUUID.WriterPackageRegistryQueue, PlugInOrder = 10)]
        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_MULTIPLE_APPEND", QueueUUID = PlugInUUID.WriterAppendingQueue, PlugInOrder = 20)]
        internal class MultiplePackage : PackageWriterBase
        {
            public static readonly List<int> ExecutedIndexes = new List<int>();

            public MultiplePackage() : base(new[] { "multiple-package", "multiple-package" }) { }

            public override List<int> OrderNumbers => new List<int>
            {
                PackagePartDefinition.POST_WORkSHEET_PACKAGE_PART_START_INDEX + 2,
                PackagePartDefinition.POST_WORkSHEET_PACKAGE_PART_START_INDEX + 3
            };
            public override List<string> PackagePartPaths => new List<string> { "xl/custom/", "xl/custom/" };
            public override List<string> PackagePartFileNames => new List<string> { "multiple1.xml", "multiple2.xml" };
            public override List<string> ContentTypes => new List<string> { ContentType, ContentType };
            public override List<string> RelationshipTypes => new List<string> { RelationshipType, RelationshipType };
            public override List<bool> ArePackagePartsRoot => new List<bool> { false, false };

            public override void Execute()
            {
                ExecutedIndexes.Add(CurrentIndex);
                base.Execute();
            }

            protected override string ElementPrefix(int index) => "multiple-package";
            public static void Reset() => ExecutedIndexes.Clear();
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_INCONSISTENT_REGISTRY", QueueUUID = PlugInUUID.WriterPackageRegistryQueue)]
        internal class InconsistentRegistryPackage : PackageWriterBase
        {
            public InconsistentRegistryPackage() : base(new[] { "inconsistent" }) { }
            public override List<int> OrderNumbers => new List<int> { 1 };
            public override List<string> PackagePartPaths => new List<string> { "xl/" };
            public override List<string> PackagePartFileNames => new List<string>(); // should not be empty when consistent
            public override List<string> ContentTypes => new List<string> { ContentType };
            public override List<string> RelationshipTypes => new List<string> { RelationshipType };
            public override List<bool> ArePackagePartsRoot => new List<bool> { false };
            [ExcludeFromCodeCoverage]
            protected override string ElementPrefix(int index) => "inconsistent";
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_INCONSISTENT_APPEND", QueueUUID = PlugInUUID.WriterAppendingQueue)]
        internal class InconsistentAppendingPackage : PackageWriterBase
        {
            public InconsistentAppendingPackage() : base(new[] { "inconsistent" }) { }
            public override List<int> OrderNumbers => new List<int> { 1 };
            public override List<string> PackagePartPaths => new List<string> { "xl/" };
            public override List<string> PackagePartFileNames => new List<string>();// should not be empty when consistent
            public override List<string> ContentTypes => new List<string> { ContentType };
            public override List<string> RelationshipTypes => new List<string> { RelationshipType };
            public override List<bool> ArePackagePartsRoot => new List<bool> { false };
            [ExcludeFromCodeCoverage]
            protected override string ElementPrefix(int index) => "inconsistent";
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

            void IPluginWriter.Init(IBaseWriter baseWriter)
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
