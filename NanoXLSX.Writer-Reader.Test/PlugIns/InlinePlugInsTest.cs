using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using NanoXLSX.Interfaces.Writer;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;
using NanoXLSX.Test.Writer_Reader.Utils;
using NanoXLSX.Utils.Xml;
using Xunit;

namespace NanoXLSX.Test.Writer_Reader.PlugIns
{
    // Ensure that these tests are executed sequentially, since static repository methods may be called 
    [Collection(nameof(SequentialCollection3))]
    public class InlinePlugInsTest : IDisposable
    {
        public void Dispose()
        {
            PlugInLoader.DisposePlugins();
        }

        [Theory(DisplayName = "Test of the plug-in handling for inline plug-ins")]
        [InlineData(typeof(InlineAppMetadataWriter), "docProps/app.xml", "inline_app_metadata")]
        [InlineData(typeof(InlineCoreMetadataWriter), "docProps/core.xml", "inline_core_metadata")]
        [InlineData(typeof(InlineStyleWriter), "xl/styles.xml", "replacing_style")]
        [InlineData(typeof(InlineThemeWriter), "xl/theme/theme1.xml", "replacing_theme")]
        [InlineData(typeof(InlineSharedStringWriter), "xl/sharedStrings.xml", "replacing_shared_strings")]
        [InlineData(typeof(InlineWorksheetWriter), "xl/worksheets/sheet1.xml", "replacing_worksheet")]
        [InlineData(typeof(InlineWorkbookWriter), "xl/workbook.xml", "replacing_workbook")]
        public void MetadataAppWriterTest(Type readerType, string expectedPath, string expectedReferenceValue)
        {
            List<Type> plugins = new List<Type>();
            // Note: These plug-ins may lead to an invalid XLSX file, depending on the RId and metadata of the packed file. It is just to test the plug-in functionality
            plugins.Add(readerType);
            PlugInLoader.InjectPlugins(plugins);
            Workbook wb = new Workbook("sheet1");

            wb.CurrentWorksheet.AddCell(expectedReferenceValue, "A1"); // Write reference value
            using (MemoryStream ms = new MemoryStream())
            {
                wb.SaveAsStream(ms, true);
                ms.Position = 0;
                TestUtils.AssertZipEntry(ms, expectedPath, expectedReferenceValue);
            }

        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "MetadatAppWriterPlugIn1", QueueUUID = PlugInUUID.METADATA_APP_INLINE_WRITER)]
        public class InlineAppMetadataWriter : IInlinePlugInWriter
        {
            private string testValue = "test";
            public Workbook Workbook { get; set; }
            public XmlElement RootElement { get; set; }

            [ExcludeFromCodeCoverage]
            public XmlElement XmlElement
            {
                get
                {
                    return RootElement;
                }
            }

            public void Execute()
            {
                RootElement.AddChildElementWithValue("test", testValue);
            }

            public void Init(ref XmlElement rootElement, Workbook workbook, int? index = null)
            {
                this.Workbook = workbook;
                this.RootElement = rootElement;
                if (Workbook.Worksheets[0].HasCell(0, 0))
                {
                    testValue = Workbook.Worksheets[0].Cells["A1"].Value.ToString();
                }
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "MetadatCoreWriterPlugIn1", QueueUUID = PlugInUUID.METADATA_CORE_INLINE_WRITER)]
        public class InlineCoreMetadataWriter : IInlinePlugInWriter
        {
            private string testValue = "test";
            public Workbook Workbook { get; set; }
            public XmlElement RootElement { get; set; }

            [ExcludeFromCodeCoverage]
            public XmlElement XmlElement
            {
                get
                {
                    return RootElement;
                }
            }

            public void Execute()
            {
                RootElement.AddChildElementWithValue("test", testValue);
            }

            public void Init(ref XmlElement rootElement, Workbook workbook, int? index = null)
            {
                this.Workbook = workbook;
                this.RootElement = rootElement;
                if (Workbook.Worksheets[0].HasCell(0, 0))
                {
                    testValue = Workbook.Worksheets[0].Cells["A1"].Value.ToString();
                }
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "StyleWriterPlugIn1", QueueUUID = PlugInUUID.STYLE_INLINE_WRITER)]
        public class InlineStyleWriter : IInlinePlugInWriter
        {
            private string testValue = "test";
            public Workbook Workbook { get; set; }
            public XmlElement RootElement { get; set; }

            [ExcludeFromCodeCoverage]
            public XmlElement XmlElement
            {
                get
                {
                    return RootElement;
                }
            }

            public void Execute()
            {
                RootElement.AddChildElementWithValue("test", testValue);
            }

            public void Init(ref XmlElement rootElement, Workbook workbook, int? index = null)
            {
                this.Workbook = workbook;
                this.RootElement = rootElement;
                if (Workbook.Worksheets[0].HasCell(0, 0))
                {
                    testValue = Workbook.Worksheets[0].Cells["A1"].Value.ToString();
                }
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "ThemeWriterPlugIn1", QueueUUID = PlugInUUID.THEME_INLINE_WRITER)]
        public class InlineThemeWriter : IInlinePlugInWriter
        {
            private string testValue = "test";
            public Workbook Workbook { get; set; }
            public XmlElement RootElement { get; set; }

            [ExcludeFromCodeCoverage]
            public XmlElement XmlElement
            {
                get
                {
                    return RootElement;
                }
            }

            public void Execute()
            {
                RootElement.AddChildElementWithValue("test", testValue);
            }

            public void Init(ref XmlElement rootElement, Workbook workbook, int? index = null)
            {
                this.Workbook = workbook;
                this.RootElement = rootElement;
                if (Workbook.Worksheets[0].HasCell(0, 0))
                {
                    testValue = Workbook.Worksheets[0].Cells["A1"].Value.ToString();
                }
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "SharedStringeWriterPlugIn1", QueueUUID = PlugInUUID.WORKSHEET_INLINE_WRITER)]
        public class InlineSharedStringWriter : IInlinePlugInWriter
        {
            private string testValue = "test";
            public Workbook Workbook { get; set; }
            public XmlElement RootElement { get; set; }

            [ExcludeFromCodeCoverage]
            public XmlElement XmlElement
            {
                get
                {
                    return RootElement;
                }
            }

            public void Execute()
            {
                RootElement.AddChildElementWithValue("test", testValue);
            }

            public void Init(ref XmlElement rootElement, Workbook workbook, int? index = null)
            {
                this.Workbook = workbook;
                this.RootElement = rootElement;
                if (Workbook.Worksheets[0].HasCell(0, 0))
                {
                    testValue = Workbook.Worksheets[0].Cells["A1"].Value.ToString();
                }
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "WorksheetWriterPlugIn1", QueueUUID = PlugInUUID.WORKSHEET_INLINE_WRITER)]
        public class InlineWorksheetWriter : IInlinePlugInWriter
        {
            private string testValue = "test";
            public Workbook Workbook { get; set; }
            public XmlElement RootElement { get; set; }

            [ExcludeFromCodeCoverage]
            public XmlElement XmlElement
            {
                get
                {
                    return RootElement;
                }
            }

            public void Execute()
            {
                RootElement.AddChildElementWithValue("test", testValue);
            }

            public void Init(ref XmlElement rootElement, Workbook workbook, int? index = null)
            {
                this.Workbook = workbook;
                this.RootElement = rootElement;
                if (Workbook.Worksheets[0].HasCell(0, 0))
                {
                    testValue = Workbook.Worksheets[0].Cells["A1"].Value.ToString();
                }
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "WorkbookWriterPlugIn1", QueueUUID = PlugInUUID.WORKBOOK_INLINE_WRITER)]
        public class InlineWorkbookWriter : IInlinePlugInWriter
        {
            private string testValue = "test";
            public Workbook Workbook { get; set; }
            public XmlElement RootElement { get; set; }

            [ExcludeFromCodeCoverage]
            public XmlElement XmlElement
            {
                get
                {
                    return RootElement;
                }
            }

            public void Execute()
            {
                RootElement.AddChildElementWithValue("test", testValue);
            }

            public void Init(ref XmlElement rootElement, Workbook workbook, int? index = null)
            {
                this.Workbook = workbook;
                this.RootElement = rootElement;
                if (Workbook.Worksheets[0].HasCell(0, 0))
                {
                    testValue = Workbook.Worksheets[0].Cells["A1"].Value.ToString();
                }
            }
        }

    }
}
