using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Reflection;
using System.Xml;
using NanoXLSX.Interfaces.Reader;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;
using NanoXLSX.Styles;
using NanoXLSX.Test.Writer_Reader.Utils;
using NanoXLSX.Utils;
using NanoXLSX.Utils.Xml;
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
        [InlineData(typeof(InlineAppMetadataReader), nameof(AppMetadataTestData), "NanoXLSXv3", PlugInUUID.METADATA_APP_INLINE_READER)]
        [InlineData(typeof(InlineCoreMetadataReader), nameof(CoreMetadataTestData), "UnitTest", PlugInUUID.METADATA_CORE_INLINE_READER)]
        [InlineData(typeof(InlineSharedStringReader), nameof(SharedStringTestData), "TestValue", PlugInUUID.SHARED_STRINGS_INLINE_READER)]
        [InlineData(typeof(InlineThemeReader), nameof(ThemeTestData), "Theme1", PlugInUUID.THEME_INLINE_READER)]
        [InlineData(typeof(InlineStyleReader), nameof(StyleTestData), "Papyrus", PlugInUUID.STYLE_INLINE_READER)]
        [InlineData(typeof(RelationshipReader), nameof(DummyTestData), "/xl/worksheets/sheet1.xml", PlugInUUID.RELATIONSHIP_INLINE_READER)]
        [InlineData(typeof(SharedStringsReader), nameof(DummyTestData), "TestValue2", PlugInUUID.SHARED_STRINGS_INLINE_READER)]
        [InlineData(typeof(WorkbookTestReader), nameof(WorkbookTestData), "Worksheet01", PlugInUUID.WORKBOOK_INLINE_READER)]
        [InlineData(typeof(WorksheetReader), nameof(WorksheetTestData), "27", PlugInUUID.WORKSHEET_INLINE_READER)]
        public void InlineReaderPluginTest(Type readerType, string setupMethodName, string expectedReferenceValue, string pluginUuid)
        {
            Workbook wb = new Workbook("sheet1");
            var setupMethod = typeof(InlineReaderPluginsTest).GetMethod(setupMethodName,
            BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
            setupMethod.Invoke(null, new object[] { wb, expectedReferenceValue });

            List<Type> plugins = new List<Type>();
            plugins.Add(readerType);
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
                Assert.True(false, "No suitable data could be found");
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


        [NanoXlsxQueuePlugIn(PlugInUUID = "MetadatAppReaderPlugIn1", QueueUUID = PlugInUUID.METADATA_APP_INLINE_READER)]
        public class InlineAppMetadataReader : IInlinePlugInReader
        {
            private const string TEST_NODE = "Application";
            private MemoryStream stream;
            public Workbook Workbook { get; set; }

            public void Execute()
            {
                string testValue = TestUtils.ReadFirstNodeValue(stream, TEST_NODE);
                Workbook.AuxiliaryData.SetData(PlugInUUID.METADATA_APP_INLINE_READER, 0, testValue, true);
                this.stream.Position = 0;
            }

            public void Init(ref MemoryStream stream, Workbook workbook, int? index = null)
            {
                this.stream = stream;
                this.stream.Position = 0;
                this.Workbook = workbook;
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "MetadatCoreReaderPlugIn1", QueueUUID = PlugInUUID.METADATA_CORE_INLINE_READER)]
        public class InlineCoreMetadataReader : IInlinePlugInReader
        {
            private const string TEST_NODE = "category";
            private MemoryStream stream;
            public Workbook Workbook { get; set; }

            public void Execute()
            {
                string testValue = TestUtils.ReadFirstNodeValue(stream, TEST_NODE);
                Workbook.AuxiliaryData.SetData(PlugInUUID.METADATA_CORE_INLINE_READER, 0, testValue, true);
                this.stream.Position = 0;
            }

            public void Init(ref MemoryStream stream, Workbook workbook, int? index = null)
            {
                this.stream = stream;
                this.stream.Position = 0;
                this.Workbook = workbook;
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "SharedStringReaderPlugIn1", QueueUUID = PlugInUUID.SHARED_STRINGS_INLINE_READER)]
        public class InlineSharedStringReader : IInlinePlugInReader
        {
            private const string TEST_NODE = "t";
            private MemoryStream stream;
            public Workbook Workbook { get; set; }

            public void Execute()
            {
                string testValue = TestUtils.ReadFirstNodeValue(stream, TEST_NODE);
                Workbook.AuxiliaryData.SetData(PlugInUUID.SHARED_STRINGS_INLINE_READER, 0, testValue, true);
                this.stream.Position = 0;
            }

            public void Init(ref MemoryStream stream, Workbook workbook, int? index = null)
            {
                this.stream = stream;
                this.stream.Position = 0;
                this.Workbook = workbook;
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "ThemeReaderPlugIn1", QueueUUID = PlugInUUID.THEME_INLINE_READER)]
        public class InlineThemeReader : IInlinePlugInReader
        {
            private const string TEST_NODE = "theme";
            private const string TEST_ATTRIBUTE = "name";
            private MemoryStream stream;
            public Workbook Workbook { get; set; }

            public void Execute()
            {
                string testValue = TestUtils.ReadFirstAttributeValue(stream, TEST_NODE, TEST_ATTRIBUTE);
                Workbook.AuxiliaryData.SetData(PlugInUUID.THEME_INLINE_READER, 0, testValue, true);
                this.stream.Position = 0;
            }

            public void Init(ref MemoryStream stream, Workbook workbook, int? index = null)
            {
                this.stream = stream;
                this.stream.Position = 0;
                this.Workbook = workbook;
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "StyleReaderPlugIn1", QueueUUID = PlugInUUID.STYLE_INLINE_READER)]
        public class InlineStyleReader : IInlinePlugInReader
        {
            private const string TEST_NODE = "name";
            private const string TEST_ATTRIBUTE = "val";
            private MemoryStream stream;
            public Workbook Workbook { get; set; }

            public void Execute()
            {
                List<string> testValues = TestUtils.ReadAllAttributeValues(stream, TEST_NODE, TEST_ATTRIBUTE);
                Workbook.AuxiliaryData.SetData(PlugInUUID.STYLE_INLINE_READER, 1, testValues, true);
                this.stream.Position = 0;
            }

            public void Init(ref MemoryStream stream, Workbook workbook, int? index = null)
            {
                this.stream = stream;
                this.stream.Position = 0;
                this.Workbook = workbook;
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "RelationshipReaderPlugIn1", QueueUUID = PlugInUUID.RELATIONSHIP_INLINE_READER)]
        public class RelationshipReader : IInlinePlugInReader
        {
            private const string TEST_NODE = "Relationship";
            private const string TEST_ATTRIBUTE = "Target";
            private MemoryStream stream;
            public Workbook Workbook { get; set; }

            public void Execute()
            {
                List<string> testValues = TestUtils.ReadAllAttributeValues(stream, TEST_NODE, TEST_ATTRIBUTE);
                Workbook.AuxiliaryData.SetData(PlugInUUID.RELATIONSHIP_INLINE_READER, 1, testValues, true);
                this.stream.Position = 0;
            }

            public void Init(ref MemoryStream stream, Workbook workbook, int? index = null)
            {
                this.stream = stream;
                this.stream.Position = 0;
                this.Workbook = workbook;
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "SharedStringsReaderPlugIn1", QueueUUID = PlugInUUID.SHARED_STRINGS_INLINE_READER)]
        public class SharedStringsReader : IInlinePlugInReader
        {
            private const string TEST_NODE = "t";
            private MemoryStream stream;
            public Workbook Workbook { get; set; }

            public void Execute()
            {
                string testValue = TestUtils.ReadFirstNodeValue(stream, TEST_NODE);
                Workbook.AuxiliaryData.SetData(PlugInUUID.SHARED_STRINGS_INLINE_READER, 0, testValue, true);
                this.stream.Position = 0;
            }

            public void Init(ref MemoryStream stream, Workbook workbook, int? index = null)
            {
                this.stream = stream;
                this.stream.Position = 0;
                this.Workbook = workbook;
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "WorkbookReaderPlugIn1", QueueUUID = PlugInUUID.WORKBOOK_INLINE_READER)]
        public class WorkbookTestReader : IInlinePlugInReader
        {
            private const string TEST_NODE = "sheet";
            private const string TEST_ATTRIBUTE = "name";
            private MemoryStream stream;
            public Workbook Workbook { get; set; }

            public void Execute()
            {
                string testValue = TestUtils.ReadFirstAttributeValue(stream, TEST_NODE, TEST_ATTRIBUTE);
                Workbook.AuxiliaryData.SetData(PlugInUUID.WORKBOOK_INLINE_READER, 0, testValue, true);
                this.stream.Position = 0;
            }

            public void Init(ref MemoryStream stream, Workbook workbook, int? index = null)
            {
                this.stream = stream;
                this.stream.Position = 0;
                this.Workbook = workbook;
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "WorksheetReaderPlugIn1", QueueUUID = PlugInUUID.WORKSHEET_INLINE_READER)]
        public class WorksheetReader : IInlinePlugInReader
        {
            private const string TEST_NODE = "sheetFormatPr";
            private const string TEST_ATTRIBUTE = "defaultRowHeight";
            private MemoryStream stream;
            public Workbook Workbook { get; set; }

            public void Execute()
            {
                string testValue = TestUtils.ReadFirstAttributeValue(stream, TEST_NODE, TEST_ATTRIBUTE);
                Workbook.AuxiliaryData.SetData(PlugInUUID.WORKSHEET_INLINE_READER, 0, testValue, true);
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
