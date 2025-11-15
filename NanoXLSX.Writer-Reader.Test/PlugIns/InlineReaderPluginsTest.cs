using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using NanoXLSX.Extensions;
using NanoXLSX.Interfaces.Reader;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;
using NanoXLSX.Styles;
using NanoXLSX.Test.Writer_Reader.Utils;
using NanoXLSX.Utils;
using Xunit;

namespace NanoXLSX.Test.Writer_Reader.PlugIns
{
    // Ensure that these tests are executed sequentially, since static repository methods may be called 
    [Collection(nameof(SequentialCollection3))]
    public class InlineReaderPluginsTest : IDisposable
    {
        public void Dispose()
        {
            PlugInLoader.DisposePlugins();
        }

        private const string VALUE_ID = "valueId";

        [Theory(DisplayName = "Test of the plug-in handling for inline reader plug-ins")]
        [InlineData(typeof(InlineAppMetadataReader), nameof(AppMetadataTestData), "NanoXLSXv3", PlugInUUID.MetadataAppInlineReader)]
        [InlineData(typeof(InlineCoreMetadataReader), nameof(CoreMetadataTestData), "UnitTest", PlugInUUID.MetadataCoreInlineReader)]
        [InlineData(typeof(InlineSharedStringReader), nameof(SharedStringTestData), "TestValue", PlugInUUID.SharedStringsInlineReader)]
        [InlineData(typeof(InlineThemeReader), nameof(ThemeTestData), "Theme1", PlugInUUID.ThemeInlineReader)]
        [InlineData(typeof(InlineStyleReader), nameof(StyleTestData), "Papyrus", PlugInUUID.StyleInlineReader)]
        [InlineData(typeof(RelationshipReader), nameof(DummyTestData), "/xl/worksheets/sheet1.xml", PlugInUUID.RelationshipInlineReader)]
        [InlineData(typeof(SharedStringsReader), nameof(DummyTestData), "TestValue2", PlugInUUID.SharedStringsInlineReader)]
        [InlineData(typeof(WorkbookTestReader), nameof(WorkbookTestData), "Worksheet01", PlugInUUID.WorkbookInlineReader)]
        [InlineData(typeof(WorksheetReader), nameof(WorksheetTestData), "27", PlugInUUID.WorksheetInlineReader)]
        public void InlineReaderPluginTest(Type readerType, string setupMethodName, string expectedReferenceValue, string pluginUuid)
        {
            Workbook wb = new Workbook("sheet1");
            var setupMethod = typeof(InlineReaderPluginsTest).GetMethod(setupMethodName,
            BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
            setupMethod.Invoke(null, new object[] { wb, expectedReferenceValue });

            List<Type> plugins = new List<Type>
            {
                readerType
            };
            PlugInLoader.InjectPlugins(plugins);

            wb.CurrentWorksheet.AddCell(expectedReferenceValue, "A1"); // Write reference value
            using (MemoryStream ms = new MemoryStream())
            {
                wb.SaveAsStream(ms, true);
                ms.Position = 0;
                Workbook wb2 = WorkbookReader.Load(ms);
                Assert.NotNull(wb2);
                string singleValue = wb2.AuxiliaryData.GetData<string>(pluginUuid, 0);
                if (singleValue != null)
                {
                    Assert.Equal(expectedReferenceValue, wb2.AuxiliaryData.GetData<string>(pluginUuid, 0));
                    return;
                }
                else
                {
                    List<string> values = wb2.AuxiliaryData.GetData<List<string>>(pluginUuid, 1);
                    Assert.Contains(expectedReferenceValue, values);
                    return;
                }
            }

        }

        public static void AppMetadataTestData(Workbook wb, string expectedValue)
        {
            wb.WorkbookMetadata.Application = expectedValue; // Application is in App metadata
        }

        public static void CoreMetadataTestData(Workbook wb, string expectedValue)
        {
            wb.WorkbookMetadata.Category = expectedValue; // Category is in Core metadata
        }
        public static void SharedStringTestData(Workbook wb, string expectedValue)
        {
            wb.CurrentWorksheet.AddNextCell(expectedValue);
        }
        public static void ThemeTestData(Workbook wb, string expectedValue)
        {
            wb.WorkbookTheme.Name = expectedValue;
        }
        public static void StyleTestData(Workbook wb, string expectedValue)
        {
            Style style = new Style();
            style.CurrentFont.Name = expectedValue;
            wb.CurrentWorksheet.AddCell(expectedValue, "A2", style); // A1 will be occupied by the unit test method
        }
        public static void DummyTestData(Workbook wb, string expectedValue)
        {
            // NoOp
        }
        public static void WorkbookTestData(Workbook wb, string expectedValue)
        {
            wb.CurrentWorksheet.SetSheetName(expectedValue); // Located in workbook.xml
        }
        public static void WorksheetTestData(Workbook wb, string expectedValue)
        {
            wb.CurrentWorksheet.DefaultRowHeight = ParserUtils.ParseInt(expectedValue);
        }


        [NanoXlsxQueuePlugIn(PlugInUUID = "MetadatAppReaderPlugIn1", QueueUUID = PlugInUUID.MetadataAppInlineReader)]
        public class InlineAppMetadataReader : IInlinePlugInReader
        {
            private const string TEST_NODE = "Application";
            private MemoryStream stream;
            public Workbook Workbook { get; set; }

            public void Execute()
            {
                string testValue = TestUtils.ReadFirstNodeValue(stream, TEST_NODE);
                Workbook.AuxiliaryData.SetData(PlugInUUID.MetadataAppInlineReader, 0, testValue, true);
                this.stream.Position = 0;
            }

            public void Init(ref MemoryStream stream, Workbook workbook, int? index = null)
            {
                this.stream = stream;
                this.stream.Position = 0;
                this.Workbook = workbook;
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "MetadatCoreReaderPlugIn1", QueueUUID = PlugInUUID.MetadataCoreInlineReader)]
        public class InlineCoreMetadataReader : IInlinePlugInReader
        {
            private const string TEST_NODE = "category";
            private MemoryStream stream;
            public Workbook Workbook { get; set; }

            public void Execute()
            {
                string testValue = TestUtils.ReadFirstNodeValue(stream, TEST_NODE);
                Workbook.AuxiliaryData.SetData(PlugInUUID.MetadataCoreInlineReader, 0, testValue, true);
                this.stream.Position = 0;
            }

            public void Init(ref MemoryStream stream, Workbook workbook, int? index = null)
            {
                this.stream = stream;
                this.stream.Position = 0;
                this.Workbook = workbook;
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "SharedStringReaderPlugIn1", QueueUUID = PlugInUUID.SharedStringsInlineReader)]
        public class InlineSharedStringReader : IInlinePlugInReader
        {
            private const string TEST_NODE = "t";
            private MemoryStream stream;
            public Workbook Workbook { get; set; }

            public void Execute()
            {
                string testValue = TestUtils.ReadFirstNodeValue(stream, TEST_NODE);
                Workbook.AuxiliaryData.SetData(PlugInUUID.SharedStringsInlineReader, 0, testValue, true);
                this.stream.Position = 0;
            }

            public void Init(ref MemoryStream stream, Workbook workbook, int? index = null)
            {
                this.stream = stream;
                this.stream.Position = 0;
                this.Workbook = workbook;
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "ThemeReaderPlugIn1", QueueUUID = PlugInUUID.ThemeInlineReader)]
        public class InlineThemeReader : IInlinePlugInReader
        {
            private const string TEST_NODE = "theme";
            private const string TEST_ATTRIBUTE = "name";
            private MemoryStream stream;
            public Workbook Workbook { get; set; }

            public void Execute()
            {
                string testValue = TestUtils.ReadFirstAttributeValue(stream, TEST_NODE, TEST_ATTRIBUTE);
                Workbook.AuxiliaryData.SetData(PlugInUUID.ThemeInlineReader, 0, testValue, true);
                this.stream.Position = 0;
            }

            public void Init(ref MemoryStream stream, Workbook workbook, int? index = null)
            {
                this.stream = stream;
                this.stream.Position = 0;
                this.Workbook = workbook;
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "StyleReaderPlugIn1", QueueUUID = PlugInUUID.StyleInlineReader)]
        public class InlineStyleReader : IInlinePlugInReader
        {
            private const string TEST_NODE = "name";
            private const string TEST_ATTRIBUTE = "val";
            private MemoryStream stream;
            public Workbook Workbook { get; set; }

            public void Execute()
            {
                List<string> testValues = TestUtils.ReadAllAttributeValues(stream, TEST_NODE, TEST_ATTRIBUTE);
                Workbook.AuxiliaryData.SetData(PlugInUUID.StyleInlineReader, 1, testValues, true);
                this.stream.Position = 0;
            }

            public void Init(ref MemoryStream stream, Workbook workbook, int? index = null)
            {
                this.stream = stream;
                this.stream.Position = 0;
                this.Workbook = workbook;
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "RelationshipReaderPlugIn1", QueueUUID = PlugInUUID.RelationshipInlineReader)]
        public class RelationshipReader : IInlinePlugInReader
        {
            private const string TEST_NODE = "Relationship";
            private const string TEST_ATTRIBUTE = "Target";
            private MemoryStream stream;
            public Workbook Workbook { get; set; }

            public void Execute()
            {
                List<string> testValues = TestUtils.ReadAllAttributeValues(stream, TEST_NODE, TEST_ATTRIBUTE);
                Workbook.AuxiliaryData.SetData(PlugInUUID.RelationshipInlineReader, 1, testValues, true);
                this.stream.Position = 0;
            }

            public void Init(ref MemoryStream stream, Workbook workbook, int? index = null)
            {
                this.stream = stream;
                this.stream.Position = 0;
                this.Workbook = workbook;
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "SharedStringsReaderPlugIn1", QueueUUID = PlugInUUID.SharedStringsInlineReader)]
        public class SharedStringsReader : IInlinePlugInReader
        {
            private const string TEST_NODE = "t";
            private MemoryStream stream;
            public Workbook Workbook { get; set; }

            public void Execute()
            {
                string testValue = TestUtils.ReadFirstNodeValue(stream, TEST_NODE);
                Workbook.AuxiliaryData.SetData(PlugInUUID.SharedStringsInlineReader, 0, testValue, true);
                this.stream.Position = 0;
            }

            public void Init(ref MemoryStream stream, Workbook workbook, int? index = null)
            {
                this.stream = stream;
                this.stream.Position = 0;
                this.Workbook = workbook;
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "WorkbookReaderPlugIn1", QueueUUID = PlugInUUID.WorkbookInlineReader)]
        public class WorkbookTestReader : IInlinePlugInReader
        {
            private const string TEST_NODE = "sheet";
            private const string TEST_ATTRIBUTE = "name";
            private MemoryStream stream;
            public Workbook Workbook { get; set; }

            public void Execute()
            {
                string testValue = TestUtils.ReadFirstAttributeValue(stream, TEST_NODE, TEST_ATTRIBUTE);
                Workbook.AuxiliaryData.SetData(PlugInUUID.WorkbookInlineReader, 0, testValue, true);
                this.stream.Position = 0;
            }

            public void Init(ref MemoryStream stream, Workbook workbook, int? index = null)
            {
                this.stream = stream;
                this.stream.Position = 0;
                this.Workbook = workbook;
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "WorksheetReaderPlugIn1", QueueUUID = PlugInUUID.WorksheetInlineReader)]
        public class WorksheetReader : IInlinePlugInReader
        {
            private const string TEST_NODE = "sheetFormatPr";
            private const string TEST_ATTRIBUTE = "defaultRowHeight";
            private MemoryStream stream;
            public Workbook Workbook { get; set; }

            public void Execute()
            {
                string testValue = TestUtils.ReadFirstAttributeValue(stream, TEST_NODE, TEST_ATTRIBUTE);
                Workbook.AuxiliaryData.SetData(PlugInUUID.WorksheetInlineReader, 0, testValue, true);
                this.stream.Position = 0;
            }

            public void Init(ref MemoryStream stream, Workbook workbook, int? index = null)
            {
                this.stream = stream;
                this.stream.Position = 0;
                this.Workbook = workbook;
            }
        }

    }
}
