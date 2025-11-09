using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using NanoXLSX.Interfaces.Plugin;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;
using Xunit;

namespace NanoXLSX.Test.Writer_Reader.PlugInsTest
{
    // Ensure that these tests are executed sequentially, since static repository methods may be called 
    [Collection(nameof(SequentialCollection4))]
    public class ReaderPlugInLoaderTest : IDisposable
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

        [Theory(DisplayName = "Test of the plug-in handling for the registration of a single reader package")]
        [InlineData(typeof(TestReaderPackage), "TEST_READER_PLUGIN_1")]
        [InlineData(typeof(TestReaderPackageWithoutStream), "TEST_READER_PLUGIN_NO_STREAM")]
        public void ReaderPackageRegistrationTest(Type pluginType, string pluginUuid)
        {
            List<Type> plugins = new List<Type>();
            plugins.Add(pluginType);
            PlugInLoader.InjectPlugins(plugins);

            Workbook wb = CreateTestWorkbook();
            using (MemoryStream ms = new MemoryStream())
            {
                wb.SaveAsStream(ms, true);
                ms.Position = 0;

                Workbook loadedWorkbook = WorkbookReader.Load(ms);

                string testData = loadedWorkbook.AuxiliaryData.GetData<string>(pluginUuid, 0);
                Assert.NotNull(testData);
                Assert.Equal("executed", testData);
            }
        }

        [Fact(DisplayName = "Test of the plug-in handling for the registration of multiple reader packages in a queue")]
        public void ReaderPackageRegistrationTest2()
        {
            List<Type> plugins = new List<Type>();
            plugins.Add(typeof(TestReaderPackage));
            plugins.Add(typeof(TestReaderPackage2));
            PlugInLoader.InjectPlugins(plugins);

            Workbook wb = CreateTestWorkbook();
            using (MemoryStream ms = new MemoryStream())
            {
                wb.SaveAsStream(ms, true);
                ms.Position = 0;

                Workbook loadedWorkbook = WorkbookReader.Load(ms);

                string testData1 = loadedWorkbook.AuxiliaryData.GetData<string>("TEST_READER_PLUGIN_1", 0);
                string testData2 = loadedWorkbook.AuxiliaryData.GetData<string>("TEST_READER_PLUGIN_2", 0);

                Assert.NotNull(testData1);
                Assert.Equal("executed", testData1);
                Assert.NotNull(testData2);
                Assert.Equal("executed", testData2);
            }
        }

        [Fact(DisplayName = "Test of the plug-in handling for reader queue execution order")]
        public void ReaderQueueExecutionOrderTest()
        {
            List<Type> plugins = new List<Type>();
            plugins.Add(typeof(TestReaderPackageOrder2));
            plugins.Add(typeof(TestReaderPackageOrder1));
            PlugInLoader.InjectPlugins(plugins);

            Workbook wb = CreateTestWorkbook();
            using (MemoryStream ms = new MemoryStream())
            {
                wb.SaveAsStream(ms, true);
                ms.Position = 0;

                Workbook loadedWorkbook = WorkbookReader.Load(ms);

                int order1 = loadedWorkbook.AuxiliaryData.GetData<int>("TEST_READER_PLUGIN_ORDER_1", 0);
                int order2 = loadedWorkbook.AuxiliaryData.GetData<int>("TEST_READER_PLUGIN_ORDER_2", 0);

                Assert.True(order1 < order2, $"Plugin with lower order number should execute first (order1={order1}, order2={order2})");
            }
        }

        [Fact(DisplayName = "Test of reader plugin with non-existent stream entry")]
        public void ReaderPackageWithMissingStreamTest()
        {
            List<Type> plugins = new List<Type>();
            plugins.Add(typeof(TestReaderPackageNonExistentStream));
            PlugInLoader.InjectPlugins(plugins);

            Workbook wb = CreateTestWorkbook();
            using (MemoryStream ms = new MemoryStream())
            {
                wb.SaveAsStream(ms, true);
                ms.Position = 0;

                Workbook loadedWorkbook = WorkbookReader.Load(ms);

                // Plugin should be skipped if stream is not found
                string testData = loadedWorkbook.AuxiliaryData.GetData<string>("TEST_READER_PLUGIN_MISSING_STREAM", 0);
                Assert.Null(testData);
            }
        }

        [Fact(DisplayName = "Test of reader plugin with existing stream entry")]
        public void ReaderPackageWithExistingStreamTest()
        {
            List<Type> plugins = new List<Type>();
            plugins.Add(typeof(TestReaderPackageExistingStream));
            PlugInLoader.InjectPlugins(plugins);

            Workbook wb = CreateTestWorkbook();
            using (MemoryStream ms = new MemoryStream())
            {
                wb.SaveAsStream(ms, true);
                ms.Position = 0;

                Workbook loadedWorkbook = WorkbookReader.Load(ms);

                // Plugin should be executed with the stream content
                string testData = loadedWorkbook.AuxiliaryData.GetData<string>("TEST_READER_PLUGIN_EXISTING_STREAM", 0);
                Assert.NotNull(testData);
                Assert.Equal("stream_processed", testData);
            }
        }

        private Workbook CreateTestWorkbook()
        {
            Workbook wb = new Workbook("worksheet1");
            wb.CurrentWorksheet.AddNextCell("Test");
            wb.CurrentWorksheet.AddNextCell(123);
            return wb;
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_READER_PLUGIN_1", QueueUUID = PlugInUUID.ReaderPrependingQueue, PlugInOrder = 199)]
        internal class TestReaderPackage : IPlugInPackageReader
        {
            public string StreamEntryName => null;
            public Workbook Workbook { get; set; }

            public void Execute()
            {
                string testData = "executed";
                Workbook.AuxiliaryData.SetData("TEST_READER_PLUGIN_1", 0, testData, true);
            }

            public void Init(MemoryStream stream, Workbook workbook, IOptions options)
            {
                this.Workbook = workbook;
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_READER_PLUGIN_2", QueueUUID = PlugInUUID.ReaderAppendingQueue, PlugInOrder = 200)]
        internal class TestReaderPackage2 : IPlugInPackageReader
        {
            public string StreamEntryName => null;
            public Workbook Workbook { get; set; }

            public void Execute()
            {
                string testData = "executed";
                Workbook.AuxiliaryData.SetData("TEST_READER_PLUGIN_2", 0, testData, true);
            }

            public void Init(MemoryStream stream, Workbook workbook, IOptions options)
            {
                this.Workbook = workbook;
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_READER_PLUGIN_ORDER_1", QueueUUID = PlugInUUID.ReaderAppendingQueue, PlugInOrder = 1)]
        internal class TestReaderPackageOrder1 : IPlugInPackageReader
        {
            public string StreamEntryName => null;
            public Workbook Workbook { get; set; }

            public void Execute()
            {
                var attribute = (NanoXlsxQueuePlugInAttribute)Attribute.GetCustomAttribute(
                                   this.GetType(), typeof(NanoXlsxQueuePlugInAttribute));
                int pluginOrder = attribute.PlugInOrder;
                Workbook.AuxiliaryData.SetData("TEST_READER_PLUGIN_ORDER_1", 0, pluginOrder, true);
            }

            public void Init(MemoryStream stream, Workbook workbook, IOptions options)
            {
                this.Workbook = workbook;
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_READER_PLUGIN_ORDER_2", QueueUUID = PlugInUUID.ReaderAppendingQueue, PlugInOrder = 2000)]
        internal class TestReaderPackageOrder2 : IPlugInPackageReader
        {
            public string StreamEntryName => null;
            public Workbook Workbook { get; set; }

            public void Execute()
            {
                var attribute = (NanoXlsxQueuePlugInAttribute)Attribute.GetCustomAttribute(
                                   this.GetType(), typeof(NanoXlsxQueuePlugInAttribute));
                int pluginOrder = attribute.PlugInOrder;
                Workbook.AuxiliaryData.SetData("TEST_READER_PLUGIN_ORDER_2", 0, pluginOrder, true);
            }

            public void Init(MemoryStream stream, Workbook workbook, IOptions options)
            {
                this.Workbook = workbook;
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_READER_PLUGIN_NO_STREAM", QueueUUID = PlugInUUID.ReaderAppendingQueue, PlugInOrder = 1000)]
        internal class TestReaderPackageWithoutStream : IPlugInPackageReader
        {
            public string StreamEntryName => null;
            public Workbook Workbook { get; set; }

            public void Execute()
            {
                string testData = "executed";
                Workbook.AuxiliaryData.SetData("TEST_READER_PLUGIN_NO_STREAM", 0, testData, true);
            }

            public void Init(MemoryStream stream, Workbook workbook, IOptions options)
            {
                Assert.Null(stream);
                this.Workbook = workbook;
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_READER_PLUGIN_MISSING_STREAM", QueueUUID = PlugInUUID.ReaderAppendingQueue, PlugInOrder = 100)]
        internal class TestReaderPackageNonExistentStream : IPlugInPackageReader
        {
            public string StreamEntryName => "xl/nonexistent/file.xml";
            [ExcludeFromCodeCoverage]
            public Workbook Workbook { get; set; }
            [ExcludeFromCodeCoverage]
            public void Execute()
            {
                // Should not be called if stream doesn't exist
                string testData = "executed";
                Workbook.AuxiliaryData.SetData("TEST_READER_PLUGIN_MISSING_STREAM", 0, testData);
            }
            [ExcludeFromCodeCoverage]
            public void Init(MemoryStream stream, Workbook workbook, IOptions options)
            {
                this.Workbook = workbook;
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_READER_PLUGIN_EXISTING_STREAM", QueueUUID = PlugInUUID.ReaderAppendingQueue, PlugInOrder = 1)]
        internal class TestReaderPackageExistingStream : IPlugInPackageReader
        {
            public string StreamEntryName => "xl/workbook.xml";
            public Workbook Workbook { get; set; }

            public void Execute()
            {
                // This should be called with the stream
                string testData = "stream_processed";
                Workbook.AuxiliaryData.SetData("TEST_READER_PLUGIN_EXISTING_STREAM", 0, testData, true);
            }

            public void Init(MemoryStream stream, Workbook workbook, IOptions options)
            {
                Assert.NotNull(stream);
                this.Workbook = workbook;
            }
        }
    }
}
