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

namespace NanoXLSX.Test.Writer_Reader.PlugInsTest
{
    // Ensure that these tests are executed sequentially, since static repository methods may be called 
    [Collection(nameof(SequentialCollection2))]
    public class PluginLoaderTest : IDisposable
    {
        public void Dispose()
        {
            PlugInLoader.DisposePlugins();
        }



        [Fact(DisplayName = "Test of the plug-in handling initializer (dummy; should not crash)")]
        public void InitializeTest()
        {
            PlugInLoader.Initialize();
        }

        [Theory(DisplayName = "Test of the plug-in handling for the registration of package parts")]
        [InlineData(typeof(TestPackage))]
        [InlineData(typeof(TestRootPackage))]
        public void PackageRegistrationTest(Type pluginType)
        {
            List<Type> plugins = new List<Type>();
            // Note: These plug-ins may lead to an invalid XLSX file, depending on the RId and metadata of the packed file. It is just to test the plug-in functionality
            plugins.Add(pluginType);
            PlugInLoader.InjectPlugins(plugins);
            Workbook wb = new Workbook();
            using (MemoryStream ms = new MemoryStream())
            {
                IPlugInPackageWriter dummy = (IPlugInPackageWriter)Activator.CreateInstance(pluginType);
                string expectedPath = dummy.PackagePartPath;
                string expectedFileName = dummy.PackagePartFileName;
                wb.SaveAsStream(ms, true);
                ms.Position = 0;
                using (var zip = new ZipArchive(ms, ZipArchiveMode.Read))
                {
                    var expectedEntryPath = $"{expectedPath}{expectedFileName}";
                    var entry = zip.GetEntry(expectedEntryPath);
                    Assert.NotNull(entry);
                }
            }
        }

        [Fact(DisplayName = "Test of the plug-in handling for the registration of multiple package parts in a queue")]
        public void PackageRegistrationTest2()
        {
            List<Type> plugins = new List<Type>();
            // Note: These plug-ins may lead to an invalid XLSX file, depending on the RId and metadata of the packed file. It is just to test the plug-in functionality
            plugins.Add(typeof(TestPackage));
            plugins.Add(typeof(TestPackage2));
            PlugInLoader.InjectPlugins(plugins);
            Workbook wb = new Workbook();
            using (MemoryStream ms = new MemoryStream())
            {
                IPlugInPackageWriter dummy1 = new TestPackage();
                IPlugInPackageWriter dummy2 = new TestPackage2();
                string expectedPath1 = dummy1.PackagePartPath;
                string expectedFileName1 = dummy1.PackagePartFileName;
                string expectedPath2 = dummy2.PackagePartPath;
                string expectedFileName2 = dummy2.PackagePartFileName;
                wb.SaveAsStream(ms, true);
                ms.Position = 0;
                using (var zip = new ZipArchive(ms, ZipArchiveMode.Read))
                {
                    var expectedEntryPath = $"{expectedPath1}{expectedFileName1}";
                    var entry = zip.GetEntry(expectedEntryPath);
                    Assert.NotNull(entry);
                    var expectedEntryPath2 = $"{expectedPath2}{expectedFileName2}";
                    var entry2 = zip.GetEntry(expectedEntryPath2);
                    Assert.NotNull(entry2);
                }
            }
        }




        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_PLUGIN_1", QueueUUID = PlugInUUID.WRITER_PACKAGE_REGISTRY_QUEUE)]
        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_PLUGIN_2", QueueUUID = PlugInUUID.WRITER_APPENDING_QUEUE)]
        internal class TestPackage : IPlugInPackageWriter
        {
            private Workbook workbook;

            public int OrderNumber => PackagePartDefinition.POST_WORSHEET_PACKAGE_PART_START_INDEX + 1;

            public string PackagePartPath => "xl/theme/";

            public string PackagePartFileName => "test.xml";

            public string ContentType => @"application/vnd.openxmlformats-package.test-file+xml";

            public string RelationshipType => @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/tests";

            public bool IsRootPackagePart => false;

            [ExcludeFromCodeCoverage]
            public Workbook Workbook { get => workbook; set => workbook = value; }

            public XmlElement XmlElement
            {
                get
                {
                    XmlElement element = XmlElement.CreateElement("test");
                    element.InnerValue = "test";
                    return element;
                }
            }

            public void Execute()
            {
                //NoOp
            }

            void IPlugInWriter.Init(IBaseWriter baseWriter)
            {
                this.workbook = baseWriter.Workbook;
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_PLUGIN_3", QueueUUID = PlugInUUID.WRITER_PACKAGE_REGISTRY_QUEUE)]
        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_PLUGIN_4", QueueUUID = PlugInUUID.WRITER_APPENDING_QUEUE)]
        internal class TestPackage2 : IPlugInPackageWriter
        {
            private Workbook workbook;

            public int OrderNumber => PackagePartDefinition.POST_WORSHEET_PACKAGE_PART_START_INDEX + 2;

            public string PackagePartPath => "xl/theme/";

            public string PackagePartFileName => "test2.xml";

            public string ContentType => @"application/vnd.openxmlformats-package.test-file+xml";

            public string RelationshipType => @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/tests";

            public bool IsRootPackagePart => false;

            [ExcludeFromCodeCoverage]
            public Workbook Workbook { get => workbook; set => workbook = value; }

            public XmlElement XmlElement
            {
                get
                {
                    XmlElement element = XmlElement.CreateElement("test");
                    element.InnerValue = "test2";
                    return element;
                }
            }

            public void Execute()
            {
                //NoOp
            }

            void IPlugInWriter.Init(IBaseWriter baseWriter)
            {
                this.workbook = baseWriter.Workbook;
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_PLUGIN_5", QueueUUID = PlugInUUID.WRITER_PACKAGE_REGISTRY_QUEUE)]
        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_PLUGIN_6", QueueUUID = PlugInUUID.WRITER_APPENDING_QUEUE)]
        public class TestRootPackage : IPlugInPackageWriter
        {
            private Workbook workbook;

            public int OrderNumber => 99;

            public string PackagePartPath => "xl/";

            public string PackagePartFileName => "rootTest.xml";

            public string ContentType => @"application/vnd.openxmlformats-package.test-file+xml";

            public string RelationshipType => @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/tests";

            public bool IsRootPackagePart => true;

            [ExcludeFromCodeCoverage]
            public Workbook Workbook { get => workbook; set => workbook = value; }

            public XmlElement XmlElement
            {
                get
                {
                    XmlElement element = XmlElement.CreateElement("test");
                    element.InnerValue = "test";
                    return element;
                }
            }

            public void Execute()
            {
                //NoOp
            }

            void IPlugInWriter.Init(IBaseWriter baseWriter)
            {
                this.workbook = baseWriter.Workbook;
            }
        }
    }
}
