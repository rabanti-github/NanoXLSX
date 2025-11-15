using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using NanoXLSX.Exceptions;
using NanoXLSX.Extensions;
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
        [InlineData(typeof(ReplaceMetadataAppReader), PlugInUUID.MetadataAppReader, "xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"")]
        [InlineData(typeof(ReplaceMetadataCoreReader), PlugInUUID.MetadataCoreReader, "xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\"")]
        [InlineData(typeof(ReplaceThemeReader), PlugInUUID.ThemeReader, "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"")]
        [InlineData(typeof(ReplaceStyleReader), PlugInUUID.StyleReader, "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"")]
        [InlineData(typeof(ReplaceSharedStringsReader), PlugInUUID.SharedStringsReader, "xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"")]
        [InlineData(typeof(ReplaceRelationshipReader), PlugInUUID.RelationshipReader, "xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"")]
        [InlineData(typeof(ReplaceWorksheetReader), PlugInUUID.WorksheetReader, "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"")]
        [InlineData(typeof(ReplaceWorkbookReader), PlugInUUID.WorkbookReader, "xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"")]
        public void ReplacingReaderPluginTest(Type readerType, string readerUuid, string expectedDocumentFragment)
        {
            List<Type> plugins = new List<Type>
            {
                // Note: These plug-ins may lead to an invalid XLSX document, since not ensured to process the corresponding stream parts as intended. It is just to test the plug-in functionality
                readerType
            };
            PlugInLoader.InjectPlugins(plugins);
            Workbook wb = new Workbook("sheet1");
            wb.CurrentWorksheet.AddCell("Test", "A1"); // Write default content
            using (MemoryStream ms = new MemoryStream())
            {
                wb.SaveAsStream(ms, true);
                ms.Position = 0;
                Workbook importedWorkbook = Extensions.WorkbookReader.Load(ms);
                string testValue = importedWorkbook.AuxiliaryData.GetData<string>(readerUuid, 0);
                Assert.Contains(expectedDocumentFragment, testValue);
            }

        }

        [Theory(DisplayName = "Test of the plug-in handling for replacing password reader plug-ins")]
        [InlineData(typeof(ReplacePasswordReader), PlugInUUID.PasswordReader, "916A", false)] // Expected hash for "TestPassword"
        [InlineData(typeof(ReplacePasswordIncompatibleReader), PlugInUUID.PasswordReader, null, true)]
        public void ReplacingPasswordReaderPluginTest(Type readerType, string readerUuid, string expectedDocumentFragment, bool expectedThrow)
        {
            List<Type> plugins = new List<Type>
            {
                // Note: These plug-ins may lead to an invalid XLSX document, since not ensured to process the corresponding stream parts as intended. It is just to test the plug-in functionality
                readerType
            };
            PlugInLoader.InjectPlugins(plugins);
            Workbook wb = new Workbook("sheet1");
            wb.CurrentWorksheet.AddCell("Test", "A1"); // Write default content
            wb.CurrentWorksheet.SetSheetProtectionPassword("TestPassword"); // should have hash: 916A
            using (MemoryStream ms = new MemoryStream())
            {
                wb.SaveAsStream(ms, true);
                ms.Position = 0;
                if (expectedThrow)
                {
                    Assert.Throws<NotSupportedContentException>(() => Extensions.WorkbookReader.Load(ms));
                    return;
                }
                else
                {
                    Workbook importedWorkbook = Extensions.WorkbookReader.Load(ms);
                    string testValue = importedWorkbook.AuxiliaryData.GetData<string>(readerUuid, 0);
                    Assert.Equal(expectedDocumentFragment, wb.CurrentWorksheet.SheetProtectionPassword.PasswordHash);
                }
            }
        }


        [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.MetadataAppReader)]
        public class ReplaceMetadataAppReader : MetadataAppReader, IPlugInReader
        {
            private MemoryStream stream;
            private Workbook workbook;
            private IOptions options;

            public void Execute()
            {
                this.stream.Position = 0;
                string content = System.Text.Encoding.UTF8.GetString(this.stream.ToArray());
                this.stream.Position = 0;
                this.workbook.AuxiliaryData.SetData(PlugInUUID.MetadataAppReader, 0, content, true);
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

        [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.MetadataCoreReader)]
        public class ReplaceMetadataCoreReader : MetadataCoreReader, IPlugInReader
        {
            private MemoryStream stream;
            private Workbook workbook;
            private IOptions options;

            public void Execute()
            {
                this.stream.Position = 0;
                string content = System.Text.Encoding.UTF8.GetString(this.stream.ToArray());
                this.stream.Position = 0;
                this.workbook.AuxiliaryData.SetData(PlugInUUID.MetadataCoreReader, 0, content, true);
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

        [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.ThemeReader)]
        public class ReplaceThemeReader : ThemeReader, IPlugInReader
        {
            private MemoryStream stream;
            private Workbook workbook;
            private IOptions options;

            public void Execute()
            {
                this.stream.Position = 0;
                string content = System.Text.Encoding.UTF8.GetString(this.stream.ToArray());
                this.stream.Position = 0;
                this.workbook.AuxiliaryData.SetData(PlugInUUID.ThemeReader, 0, content, true);
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

        [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.StyleReader)]
        public class ReplaceStyleReader : StyleReader, IPlugInReader
        {
            private MemoryStream stream;
            private Workbook workbook;
            private IOptions options;

            public void Execute()
            {
                this.stream.Position = 0;
                string content = System.Text.Encoding.UTF8.GetString(this.stream.ToArray());
                this.stream.Position = 0;
                this.workbook.AuxiliaryData.SetData(PlugInUUID.StyleReader, 0, content, true);
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

        [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.SharedStringsReader)]
        public class ReplaceSharedStringsReader : SharedStringsReader, IPlugInReader
        {
            private MemoryStream stream;
            private Workbook workbook;
            private IOptions options;

            public void Execute()
            {
                this.stream.Position = 0;
                string content = System.Text.Encoding.UTF8.GetString(this.stream.ToArray());
                this.stream.Position = 0;
                this.workbook.AuxiliaryData.SetData(PlugInUUID.SharedStringsReader, 0, content, true);
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

        [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.RelationshipReader)]
        public class ReplaceRelationshipReader : RelationshipReader, IPlugInReader
        {
            private MemoryStream stream;
            private Workbook workbook;
            private IOptions options;

            public void Execute()
            {
                this.stream.Position = 0;
                string content = System.Text.Encoding.UTF8.GetString(this.stream.ToArray());
                this.stream.Position = 0;
                this.workbook.AuxiliaryData.SetData(PlugInUUID.RelationshipReader, 0, content, true);
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

        [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.WorksheetReader)]
        public class ReplaceWorksheetReader : WorksheetReader, IPlugInReader
        {
            private MemoryStream stream;
            private Workbook workbook;
            private IOptions options;

            public void Execute()
            {
                this.stream.Position = 0;
                string content = System.Text.Encoding.UTF8.GetString(this.stream.ToArray());
                this.stream.Position = 0;
                this.workbook.AuxiliaryData.SetData(PlugInUUID.WorksheetReader, 0, content, true);
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

        [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.WorkbookReader)]
        public class ReplaceWorkbookReader : Internal.Readers.WorkbookReader, IPlugInReader
        {
            private MemoryStream stream;
            private Workbook workbook;
            private IOptions options;

            public void Execute()
            {
                this.stream.Position = 0;
                string content = System.Text.Encoding.UTF8.GetString(this.stream.ToArray());
                this.stream.Position = 0;
                this.workbook.AuxiliaryData.SetData(PlugInUUID.WorkbookReader, 0, content, true);
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


        [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.PasswordReader)]
        public class ReplacePasswordReader : IPasswordReader
        {
            public string PasswordHash { get; set; }

            public void ReadXmlAttributes(System.Xml.XmlNode node)
            {
                PasswordHash = ReaderUtils.GetAttribute(node, "password"); // Dummy read to ensure the XML node is processed
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

        [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.PasswordReader)]
        public class ReplacePasswordIncompatibleReader : LegacyPasswordReader
        {
            [ExcludeFromCodeCoverage]
            public static new void ReadXmlAttributes(System.Xml.XmlNode node)
            {
                // NoOp
            }

            public override bool ContemporaryAlgorithmDetected
            {
                get { return true; } // Force incompatible algorithm
            }
        }

    }
}
