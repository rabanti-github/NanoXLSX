using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using NanoXLSX.Interfaces;
using NanoXLSX.Interfaces.Plugin;
using NanoXLSX.Interfaces.Reader;
using NanoXLSX.Internal;
using NanoXLSX.Internal.Readers;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;
using Xunit;
using static NanoXLSX.Internal.Enums.ReaderPassword;

namespace NanoXLSX.Test.Writer_Reader.PlugIns
{
    // Ensure that these tests are executed sequentially, since static repository methods may be called 
    [Collection(nameof(SequentialCollection2))]
    public class ReplacingReaderPluginsTest : IDisposable
    {
        public void Dispose()
        {
            PlugInLoader.DisposePlugins();
        }

        [Theory(DisplayName = "Test of the plug-in handling for replacing reader plug-ins")]
        [InlineData(typeof(ReplaceMetadataAppReader), PlugInUUID.METADATA_APP_READER, "xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"")]
        [InlineData(typeof(ReplaceMetadataCoreReader), PlugInUUID.METADATA_CORE_READER, "xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\"")]
        [InlineData(typeof(ReplaceThemeReader), PlugInUUID.THEME_READER, "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"")]
        [InlineData(typeof(ReplaceStyleReader), PlugInUUID.STYLE_READER, "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"")]
        [InlineData(typeof(ReplaceSharedStringsReader), PlugInUUID.SHARED_STRINGS_READER, "xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"")]
        [InlineData(typeof(ReplaceRelationshipReader), PlugInUUID.RELATIONSHIP_READER, "xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"")]
        [InlineData(typeof(ReplaceWorksheetReader), PlugInUUID.WORKSHEET_READER, "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"")]
        [InlineData(typeof(ReplaceWorkbookReader), PlugInUUID.WORKBOOK_READER, "xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"")]
        public void ReplacingReaderPluginTest(Type readerType, string readerUuid, string expectedDocumentFragment)
        {
            List<Type> plugins = new List<Type>();
            // Note: These plug-ins may lead to an invalid XLSX document, since not ensured to process the corresponding stream parts as intended. It is just to test the plug-in functionality
            plugins.Add(readerType);
            PlugInLoader.InjectPlugins(plugins);
            Workbook wb = new Workbook("sheet1");
            wb.CurrentWorksheet.AddCell("Test", "A1"); // Write default content
            using (MemoryStream ms = new MemoryStream())
            {
                wb.SaveAsStream(ms, true);
                ms.Position = 0;
                Workbook importedWorkbook = WorkbookReader.Load(ms);
                string testValue = importedWorkbook.AuxiliaryData.GetData<string>(readerUuid, 0);
                Assert.Contains(expectedDocumentFragment, testValue);
            }

        }

        [Theory(DisplayName = "Test of the plug-in handling for replacing password reader plug-ins")]
        [InlineData(typeof(ReplacePasswordReader), PlugInUUID.PASSWORD_READER, "916A")] // Expected hash for "TestPassword"
        public void ReplacingPasswordReaderPluginTest(Type readerType, string readerUuid, string expectedDocumentFragment)
        {
            List<Type> plugins = new List<Type>();
            // Note: These plug-ins may lead to an invalid XLSX document, since not ensured to process the corresponding stream parts as intended. It is just to test the plug-in functionality
            plugins.Add(readerType);
            PlugInLoader.InjectPlugins(plugins);
            Workbook wb = new Workbook("sheet1");
            wb.CurrentWorksheet.AddCell("Test", "A1"); // Write default content
            wb.CurrentWorksheet.SetSheetProtectionPassword("TestPassword"); // should have hash: 916A
            using (MemoryStream ms = new MemoryStream())
            {
                wb.SaveAsStream(ms, true);
                ms.Position = 0;
                Workbook importedWorkbook = WorkbookReader.Load(ms);
                string testValue = importedWorkbook.AuxiliaryData.GetData<string>(readerUuid, 0);
                Assert.Equal(expectedDocumentFragment, wb.CurrentWorksheet.SheetProtectionPassword.PasswordHash);
            }
        }


        [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.METADATA_APP_READER)]
        public class ReplaceMetadataAppReader : MetadataAppReader, IPlugInReader
        {
            private MemoryStream stream;
            private Workbook workbook;
            private IOptions options;

            [ExcludeFromCodeCoverage]
            public Workbook Workbook { get; set; }
            public void Execute()
            {
                this.stream.Position = 0;
                string content = System.Text.Encoding.UTF8.GetString(this.stream.ToArray());
                this.stream.Position = 0;
                this.workbook.AuxiliaryData.SetData(PlugInUUID.METADATA_APP_READER, 0, content, true);
                base.Execute(); // Execute regular reader
            }

            public void Init(MemoryStream stream, Workbook workbook, IOptions readerOptions)
            {
                base.Init(stream, workbook, readerOptions);
                this.stream = stream;
                this.workbook = workbook;
                this.options = readerOptions;
            }
        }

        [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.METADATA_CORE_READER)]
        public class ReplaceMetadataCoreReader : MetadataCoreReader, IPlugInReader
        {
            private MemoryStream stream;
            private Workbook workbook;
            private IOptions options;

            [ExcludeFromCodeCoverage]
            public Workbook Workbook { get; set; }
            public void Execute()
            {
                this.stream.Position = 0;
                string content = System.Text.Encoding.UTF8.GetString(this.stream.ToArray());
                this.stream.Position = 0;
                this.workbook.AuxiliaryData.SetData(PlugInUUID.METADATA_CORE_READER, 0, content, true);
                base.Execute(); // Execute regular reader
            }

            public void Init(MemoryStream stream, Workbook workbook, IOptions readerOptions)
            {
                base.Init(stream, workbook, readerOptions);
                this.stream = stream;
                this.workbook = workbook;
                this.options = readerOptions;
            }
        }

        [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.THEME_READER)]
        public class ReplaceThemeReader : ThemeReader, IPlugInReader
        {
            private MemoryStream stream;
            private Workbook workbook;
            private IOptions options;

            [ExcludeFromCodeCoverage]
            public Workbook Workbook { get; set; }
            public void Execute()
            {
                this.stream.Position = 0;
                string content = System.Text.Encoding.UTF8.GetString(this.stream.ToArray());
                this.stream.Position = 0;
                this.workbook.AuxiliaryData.SetData(PlugInUUID.THEME_READER, 0, content, true);
                base.Execute(); // Execute regular reader
            }

            public void Init(MemoryStream stream, Workbook workbook, IOptions readerOptions)
            {
                base.Init(stream, workbook, readerOptions);
                this.stream = stream;
                this.workbook = workbook;
                this.options = readerOptions;
            }
        }

        [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.STYLE_READER)]
        public class ReplaceStyleReader : StyleReader, IPlugInReader
        {
            private MemoryStream stream;
            private Workbook workbook;
            private IOptions options;

            [ExcludeFromCodeCoverage]
            public Workbook Workbook { get; set; }
            public void Execute()
            {
                this.stream.Position = 0;
                string content = System.Text.Encoding.UTF8.GetString(this.stream.ToArray());
                this.stream.Position = 0;
                this.workbook.AuxiliaryData.SetData(PlugInUUID.STYLE_READER, 0, content, true);
                base.Execute(); // Execute regular reader
            }

            public void Init(MemoryStream stream, Workbook workbook, IOptions readerOptions)
            {
                base.Init(stream, workbook, readerOptions);
                this.stream = stream;
                this.workbook = workbook;
                this.options = readerOptions;
            }
        }

        [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.SHARED_STRINGS_READER)]
        public class ReplaceSharedStringsReader : SharedStringsReader, IPlugInReader
        {
            private MemoryStream stream;
            private Workbook workbook;
            private IOptions options;

            [ExcludeFromCodeCoverage]
            public Workbook Workbook { get; set; }
            public void Execute()
            {
                this.stream.Position = 0;
                string content = System.Text.Encoding.UTF8.GetString(this.stream.ToArray());
                this.stream.Position = 0;
                this.workbook.AuxiliaryData.SetData(PlugInUUID.SHARED_STRINGS_READER, 0, content, true);
                base.Execute(); // Execute regular reader
            }

            public void Init(MemoryStream stream, Workbook workbook, IOptions readerOptions)
            {
                base.Init(stream, workbook, readerOptions);
                this.stream = stream;
                this.workbook = workbook;
                this.options = readerOptions;
            }
        }

        [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.RELATIONSHIP_READER)]
        public class ReplaceRelationshipReader : RelationshipReader, IPlugInReader
        {
            private MemoryStream stream;
            private Workbook workbook;
            private IOptions options;

            [ExcludeFromCodeCoverage]
            public Workbook Workbook { get; set; }
            public void Execute()
            {
                this.stream.Position = 0;
                string content = System.Text.Encoding.UTF8.GetString(this.stream.ToArray());
                this.stream.Position = 0;
                this.workbook.AuxiliaryData.SetData(PlugInUUID.RELATIONSHIP_READER, 0, content, true);
                base.Execute(); // Execute regular reader
            }

            public void Init(MemoryStream stream, Workbook workbook, IOptions readerOptions)
            {
                base.Init(stream, workbook, readerOptions);
                this.stream = stream;
                this.workbook = workbook;
                this.options = readerOptions;
            }
        }

        [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.WORKSHEET_READER)]
        public class ReplaceWorksheetReader : WorksheetReader, IPlugInReader
        {
            private MemoryStream stream;
            private Workbook workbook;
            private IOptions options;

            [ExcludeFromCodeCoverage]
            public Workbook Workbook { get; set; }
            public void Execute()
            {
                this.stream.Position = 0;
                string content = System.Text.Encoding.UTF8.GetString(this.stream.ToArray());
                this.stream.Position = 0;
                this.workbook.AuxiliaryData.SetData(PlugInUUID.WORKSHEET_READER, 0, content, true);
                base.Execute(); // Execute regular reader
            }

            public void Init(MemoryStream stream, Workbook workbook, IOptions readerOptions)
            {
                base.Init(stream, workbook, readerOptions);
                this.stream = stream;
                this.workbook = workbook;
                this.options = readerOptions;
            }
        }

        [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.WORKBOOK_READER)]
        public class ReplaceWorkbookReader : Internal.Readers.WorkbookReader, IPlugInReader
        {
            private MemoryStream stream;
            private Workbook workbook;
            private IOptions options;

            [ExcludeFromCodeCoverage]
            public Workbook Workbook { get; set; }
            public void Execute()
            {
                this.stream.Position = 0;
                string content = System.Text.Encoding.UTF8.GetString(this.stream.ToArray());
                this.stream.Position = 0;
                this.workbook.AuxiliaryData.SetData(PlugInUUID.WORKBOOK_READER, 0, content, true);
                base.Execute(); // Execute regular reader
            }

            public void Init(MemoryStream stream, Workbook workbook, IOptions readerOptions)
            {
                base.Init(stream, workbook, readerOptions);
                this.stream = stream;
                this.workbook = workbook;
                this.options = readerOptions;
            }
        }


        [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.PASSWORD_READER)]
        public class ReplacePasswordReader : IPasswordReader
        {
            public string PasswordHash { get; set; }

            public void ReadXmlAttributes(System.Xml.XmlNode node)
            {
                PasswordHash = ReaderUtils.GetAttribute(node, "password"); // Dummy read to ensure the XML node is processed

               // this.workbook.AuxiliaryData.SetData(PlugInUUID.PASSWORD_READER, 0, content, true);
            }

            public void Init(PasswordType type, ReaderOptions readerOptions)
            {
                // NoOp
            }

            [ExcludeFromCodeCoverage]
            public void SetPassword(string plainText)
            {
                // NoOp
            }

            [ExcludeFromCodeCoverage]
            public void UnsetPassword()
            {
                // NoOp
            }

            public string GetPassword()
            {
                return null;
            }

            public bool PasswordIsSet()
            {
                return true;
            }

            [ExcludeFromCodeCoverage]
            public void CopyFrom(IPassword passwordInstance)
            {
                // NoOp
            }
        }

    }
}
