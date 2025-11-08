using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using NanoXLSX.Interfaces;
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

    public class ReplacingWriterPlugInsTest : IDisposable
    {
        public void Dispose()
        {
            PlugInLoader.DisposePlugins();
        }

        [Theory(DisplayName = "Test of the plug-in handling for replacing writer plug-ins")]
        [InlineData(typeof(ReplaceAppMetadataWriter), "docProps/app.xml", "replacing_app_metadata")]
        [InlineData(typeof(ReplaceCoreMetadataWriter), "docProps/core.xml", "replacing_core_metadata")]
        [InlineData(typeof(ReplaceStyleWriter), "xl/styles.xml", "replacing_style")]
        [InlineData(typeof(ReplaceThemeWriter), "xl/theme/theme1.xml", "replacing_theme")]
        [InlineData(typeof(ReplaceSharedStringWriter), "xl/sharedStrings.xml", "replacing_shared_strings")]
        [InlineData(typeof(ReplaceWorksheetWriter), "xl/worksheets/sheet1.xml", "replacing_worksheet")]
        [InlineData(typeof(ReplaceWorkbookWriter), "xl/workbook.xml", "replacing_workbook")]
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



        [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.METADATA_APP_WRITER)]
        public class ReplaceAppMetadataWriter : IPlugInWriter
        {
            private string testValue = "test";
            public Workbook Workbook { get; set; }

            public XmlElement XmlElement
            {
                get
                {
                    XmlElement element = XmlElement.CreateElement("test");
                    element.InnerValue = testValue;
                    return element;
                }
            }

            public void Execute()
            {
                // NoOp
            }

            void IPlugInWriter.Init(IBaseWriter baseWriter)
            {
                this.Workbook = baseWriter.Workbook;
                if (Workbook.Worksheets[0].HasCell(0, 0))
                {
                    testValue = Workbook.Worksheets[0].Cells["A1"].Value.ToString();
                }
            }
        }

        [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.METADATA_CORE_WRITER)]
        public class ReplaceCoreMetadataWriter : IPlugInWriter
        {
            private string testValue = "test";
            public Workbook Workbook { get; set; }

            public XmlElement XmlElement
            {
                get
                {
                    XmlElement element = XmlElement.CreateElement("test");
                    element.InnerValue = testValue;
                    return element;
                }
            }

            public void Execute()
            {
                // NoOp
            }

            void IPlugInWriter.Init(IBaseWriter baseWriter)
            {
                this.Workbook = baseWriter.Workbook;
                if (Workbook.Worksheets[0].HasCell(0, 0))
                {
                    testValue = Workbook.Worksheets[0].Cells["A1"].Value.ToString();
                }
            }
        }

        [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.STYLE_WRITER)]
        public class ReplaceStyleWriter : IPlugInWriter
        {
            private string testValue = "test";
            public Workbook Workbook { get; set; }

            public XmlElement XmlElement
            {
                get
                {
                    XmlElement element = XmlElement.CreateElement("test");
                    element.InnerValue = testValue;
                    return element;
                }
            }

            public void Execute()
            {
                // NoOp
            }

            void IPlugInWriter.Init(IBaseWriter baseWriter)
            {
                this.Workbook = baseWriter.Workbook;
                if (Workbook.Worksheets[0].HasCell(0, 0))
                {
                    testValue = Workbook.Worksheets[0].Cells["A1"].Value.ToString();
                }
            }
        }

        [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.THEME_WRITER)]
        public class ReplaceThemeWriter : IPlugInWriter
        {
            private string testValue = "test";
            public Workbook Workbook { get; set; }

            public XmlElement XmlElement
            {
                get
                {
                    XmlElement element = XmlElement.CreateElement("test");
                    element.InnerValue = testValue;
                    return element;
                }
            }

            public void Execute()
            {
                // NoOp
            }

            void IPlugInWriter.Init(IBaseWriter baseWriter)
            {
                this.Workbook = baseWriter.Workbook;
                if (Workbook.Worksheets[0].HasCell(0, 0))
                {
                    testValue = Workbook.Worksheets[0].Cells["A1"].Value.ToString();
                }
            }
        }

        [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.SHARED_STRINGS_WRITER)]
        public class ReplaceSharedStringWriter : ISharedStringWriter
        {
            private string testValue = "test";
            private TestSortedMap testSortedMap = new TestSortedMap();
            public Workbook Workbook { get; set; }

            public XmlElement XmlElement
            {
                get
                {
                    XmlElement element = XmlElement.CreateElement("test");
                    element.InnerValue = testValue;
                    return element;
                }
            }

            public ISortedMap SharedStrings => testSortedMap;

            public int SharedStringsTotalCount { get => testSortedMap.Count; set => _ = value; }

            public void Execute()
            {
                // NoOp
            }

            void IPlugInWriter.Init(IBaseWriter baseWriter)
            {
                this.Workbook = baseWriter.Workbook;
                if (Workbook.Worksheets[0].HasCell(0, 0))
                {
                    testValue = Workbook.Worksheets[0].Cells["A1"].Value.ToString();
                }
            }
        }

        [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.WORKSHEET_WRITER)]
        public class ReplaceWorksheetWriter : IWorksheetWriter
        {
            private string testValue = "test";
            private Worksheet worksheet;
            public Workbook Workbook { get; set; }

            public XmlElement XmlElement
            {
                get
                {
                    XmlElement element = XmlElement.CreateElement("test");
                    element.InnerValue = testValue;
                    return element;
                }
            }

            public Worksheet CurrentWorksheet { get => worksheet; set => worksheet = value; }

            public void Execute()
            {
                // NoOp
            }

            void IPlugInWriter.Init(IBaseWriter baseWriter)
            {
                this.Workbook = baseWriter.Workbook;
                if (Workbook.Worksheets[0].HasCell(0, 0))
                {
                    testValue = Workbook.Worksheets[0].Cells["A1"].Value.ToString();
                }
            }
        }

        [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.WORKBOOK_WRITER)]
        public class ReplaceWorkbookWriter : IPlugInWriter
        {
            private string testValue = "test";
            public Workbook Workbook { get; set; }

            public XmlElement XmlElement
            {
                get
                {
                    XmlElement element = XmlElement.CreateElement("test");
                    element.InnerValue = testValue;
                    return element;
                }
            }

            public void Execute()
            {
                // NoOp
            }

            void IPlugInWriter.Init(IBaseWriter baseWriter)
            {
                this.Workbook = baseWriter.Workbook;
                if (Workbook.Worksheets[0].HasCell(0, 0))
                {
                    testValue = Workbook.Worksheets[0].Cells["A1"].Value.ToString();
                }
            }
        }

        public class TestSortedMap : ISortedMap
        {
            List<IFormattableText> list = new List<IFormattableText>();
            public int Count => list.Count;

            [ExcludeFromCodeCoverage]
            public List<IFormattableText> Keys => list;

            public string Add(IFormattableText text, string referenceIndex)
            {
                // Note: Not yet cleanly implemented
                list.Add(text);
                return text.ToString();
            }
        }

    }
}
