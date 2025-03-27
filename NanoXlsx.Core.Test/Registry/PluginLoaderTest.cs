using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NanoXLSX.Interfaces.Writer;
using NanoXLSX.Registry;
using NanoXLSX.Utils.Xml;
using Xunit;

namespace NanoXLSX.Test.Core.RegistryTest
{
    public class PluginLoaderTest
    {
        [Fact(DisplayName = "Test of the plug-in handling for the registration of package parts")]
        public void PackageRegistrationTest()
        {
            List<Type> plugins = new List<Type>();
            // Note: These plug-ins will not lead to a valid XLSX file. It is just to test the plug-in functionality
            plugins.Add(typeof(TestPackageDefinition));
            plugins.Add(typeof(TestPackage));
            PlugInLoader.InjectPlugins(plugins);
            Workbook wb = new Workbook();
            using (MemoryStream ms = new MemoryStream())
            {
                TestPackage dummy = new TestPackage();
                string expectedPath = dummy.PackagePath;
                string expectedFileName = dummy.PackageFileName; 
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

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_PLUGIN_1", QueueUUID = PlugInUUID.WRITER_PACKAGE_REGISTRY_QUEUE)]
        public class TestPackageDefinition : IPlugInWriterRegistration
        {
            public int OrderNumber => 99;

            public string PackagePartPath => "xl/";

            public string PackagePartFileName => "test.xml";

            public string ContentType => @"application/vnd.openxmlformats-package.test-file+xml";

            public string RelationshipType => @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/tests";

            public bool IsRootPackagePart => false;

            public void Execute()
            {
                //NoOp
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_PLUGIN_2", QueueUUID = PlugInUUID.WRITER_APPENDING_QUEUE)]
        public class TestPackage : IPlugInPackageWriter
        {
            private Workbook workbook;
            public string PackagePath => "xl/";

            public string PackageFileName => "test.xml";

            public Workbook Workbook { get => workbook; set => workbook = value; }

            public void Execute()
            {
                // NoOp
            }

            public XmlElement GetElement()
            {
                XmlElement element = XmlElement.CreateElement("test");
                element.InnerValue = "test";
                return element;
            }

            void IPlugInWriter.Init(IBaseWriter baseWriter)
            {
                this.workbook = baseWriter.Workbook;
            }
        }

    }
}
