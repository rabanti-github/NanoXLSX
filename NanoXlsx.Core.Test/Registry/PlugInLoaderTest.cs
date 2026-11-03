using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Text;
using NanoXLSX.Interfaces;
using NanoXLSX.Interfaces.Writer;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;
using NanoXLSX.Test.Core.Utils;
using Xunit;

namespace NanoXLSX.Test.Writer_Reader.PlugInsTest
{
    // Ensure that these tests are executed sequentially (in a own collection), since static repository methods may be called 
    [Collection(nameof(SequentialPlugInCollection))]
    public class PluginLoaderTest : IDisposable
    {
        public void Dispose()
        {
            PlugInLoader.DisposePlugins();
        }


        [Fact(DisplayName = "Test of the plug-in handling initializer (dummy; should not crash or initialize twice)")]
        public void InitializeTest()
        {
            PlugInLoader.DisposePlugins(); // Test on a clean basis
            bool state = PlugInLoader.Initialize();
            Assert.True(state);
            bool state2 = PlugInLoader.Initialize();
            Assert.False(state2);
        }

        [Fact(DisplayName = "Getting the next queue plug-in returns null for an unknown last plug-in UUID")]
        public void GetNextQueuePluginWithUnknownLastUuidTest()
        {
            PlugInLoader.InjectPlugins(new List<Type> { typeof(QueuePlugin) });

            IPlugin plugin = PlugInLoader.GetNextQueuePlugIn<IPlugin>(QueueUuid, "UNKNOWN_PLUGIN", out string currentPluginUuid);

            Assert.Null(plugin);
            Assert.Null(currentPluginUuid);
        }

        [Fact(DisplayName = "Getting the next queue plug-in skips entries that do not implement the requested type")]
        public void GetNextQueuePluginSkipsIncompatibleTypeTest()
        {
            PlugInLoader.InjectPlugins(new List<Type> { typeof(QueuePlugin), typeof(WriterQueuePlugin) });

            IPluginWriter plugin = PlugInLoader.GetNextQueuePlugIn<IPluginWriter>(QueueUuid, null, out string currentPluginUuid);

            Assert.IsType<WriterQueuePlugin>(plugin);
            Assert.Equal(WriterPluginUuid, currentPluginUuid);
        }

        [Fact(DisplayName = "A queue plug-in can be created with a writer context constructor")]
        public void GetNextQueuePluginWithWriterContextTest()
        {
            PlugInLoader.InjectPlugins(new List<Type> { typeof(ContextualQueuePlugin) });
            BaseWriterStub baseWriter = new BaseWriterStub();

            IPlugin plugin = PlugInLoader.GetNextQueuePlugIn<IPlugin>(ContextualQueueUuid, null, out string currentPluginUuid, baseWriter);

            ContextualQueuePlugin contextualPlugin = Assert.IsType<ContextualQueuePlugin>(plugin);
            Assert.Same(baseWriter, contextualPlugin.BaseWriter);
            Assert.Equal(ContextualPluginUuid, currentPluginUuid);
        }

        [Theory(DisplayName = "An assembly file of the runtime or of NanoXLSX is detected as managed assembly")]
        [InlineData(typeof(object))]
        [InlineData(typeof(PlugInLoader))]
        public void IsManagedAssemblyTest(Type typeOfAssembly)
        {
            Assert.True(PlugInLoader.IsManagedAssembly(typeOfAssembly.Assembly.Location));
        }

        [Theory(DisplayName = "A PE image is only detected as managed assembly if it declares a non-empty CLI header data directory")]
        [InlineData(Pe32Magic, 16u, 0x2008u, 0x48u, true)]
        [InlineData(Pe32PlusMagic, 16u, 0x2008u, 0x48u, true)]
        [InlineData(Pe32Magic, 16u, 0u, 0u, false)]
        [InlineData(Pe32PlusMagic, 16u, 0u, 0u, false)]
        [InlineData(Pe32Magic, 16u, 0x2008u, 0u, false)]
        [InlineData(Pe32Magic, 16u, 0u, 0x48u, false)]
        [InlineData(Pe32Magic, 14u, 0x2008u, 0x48u, false)]
        [InlineData(RomImageMagic, 16u, 0x2008u, 0x48u, false)]
        public void IsManagedAssemblyPeImageTest(ushort magic, uint numberOfRvaAndSizes, uint cliHeaderRva, uint cliHeaderSize, bool expectedResult)
        {
            byte[] image = CreatePeImage(magic, numberOfRvaAndSizes, cliHeaderRva, cliHeaderSize);
            Assert.Equal(expectedResult, IsManagedAssemblyFile(image));
        }

        [Fact(DisplayName = "A managed PE image that ends within its headers is not detected as managed assembly")]
        public void IsManagedAssemblyTruncatedImageTest()
        {
            byte[] image = CreatePeImage(Pe32Magic, 16, 0x2008, 0x48);
            Assert.True(IsManagedAssemblyFile(image));
            for (int length = 0; length < image.Length; length++)
            {
                byte[] truncatedImage = new byte[length];
                Array.Copy(image, truncatedImage, length);
                Assert.False(IsManagedAssemblyFile(truncatedImage), $"Image truncated to {length} bytes");
            }
        }

        [Theory(DisplayName = "A file with a DOS signature but an invalid PE header position is not detected as managed assembly")]
        [InlineData(-1)]
        [InlineData(0x40)]
        [InlineData(int.MaxValue)]
        public void IsManagedAssemblyInvalidPeHeaderPositionTest(int peHeaderPosition)
        {
            byte[] image = CreatePeImage(Pe32Magic, 16, 0x2008, 0x48);
            using (MemoryStream stream = new MemoryStream(image))
            using (BinaryWriter writer = new BinaryWriter(stream))
            {
                stream.Position = PeHeaderPointerPosition;
                writer.Write(peHeaderPosition);
            }
            Assert.False(IsManagedAssemblyFile(image));
        }

        [Fact(DisplayName = "A file that is not a PE image is not detected as managed assembly")]
        public void IsManagedAssemblyPlainTextFileTest()
        {
            byte[] content = Encoding.ASCII.GetBytes(new string('x', 1024));
            Assert.False(IsManagedAssemblyFile(content));
        }

        /// <summary>
        /// Creates the headers of a PE image, up to the end of the CLI header data directory. All other header fields are zero
        /// </summary>
        private static byte[] CreatePeImage(ushort magic, uint numberOfRvaAndSizes, uint cliHeaderRva, uint cliHeaderSize)
        {
            const int peHeaderPosition = 0x80;
            int optionalHeaderPosition = peHeaderPosition + 24; // PE signature (4 bytes) and COFF file header (20 bytes)
            int rvaCountPosition = optionalHeaderPosition + (magic == Pe32PlusMagic ? 108 : 92);
            int cliHeaderDirectoryPosition = rvaCountPosition + 4 + (14 * 8); // Data directory 14, each directory is 8 bytes long
            byte[] image = new byte[cliHeaderDirectoryPosition + 8];
            using (MemoryStream stream = new MemoryStream(image))
            using (BinaryWriter writer = new BinaryWriter(stream))
            {
                writer.Write(Encoding.ASCII.GetBytes("MZ"));
                stream.Position = PeHeaderPointerPosition;
                writer.Write(peHeaderPosition);
                stream.Position = peHeaderPosition;
                writer.Write(Encoding.ASCII.GetBytes("PE\0\0"));
                stream.Position = optionalHeaderPosition;
                writer.Write(magic);
                stream.Position = rvaCountPosition;
                writer.Write(numberOfRvaAndSizes);
                stream.Position = cliHeaderDirectoryPosition;
                writer.Write(cliHeaderRva);
                writer.Write(cliHeaderSize);
            }
            return image;
        }

        private static bool IsManagedAssemblyFile(byte[] content)
        {
            string path = Path.GetTempFileName();
            try
            {
                File.WriteAllBytes(path, content);
                return PlugInLoader.IsManagedAssembly(path);
            }
            finally
            {
                File.Delete(path);
            }
        }

        private const string QueueUuid = "TEST_PLUGIN_LOADER_QUEUE";
        private const string WriterPluginUuid = "TEST_WRITER_QUEUE_PLUGIN";
        private const string ContextualQueueUuid = "TEST_CONTEXTUAL_PLUGIN_LOADER_QUEUE";
        private const string ContextualPluginUuid = "TEST_CONTEXTUAL_PLUGIN";
        private const ushort Pe32Magic = 0x10B;
        private const ushort Pe32PlusMagic = 0x20B;
        private const ushort RomImageMagic = 0x107;
        private const int PeHeaderPointerPosition = 0x3C;

        [NanoXlsxQueuePlugIn(PlugInUUID = "TEST_QUEUE_PLUGIN", QueueUUID = QueueUuid)]
        internal class QueuePlugin : IPlugin
        {
            [ExcludeFromCodeCoverage]
            public void Execute()
            {
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = WriterPluginUuid, QueueUUID = QueueUuid, PlugInOrder = 1)]
        internal class WriterQueuePlugin : IPluginWriter
        {
            [ExcludeFromCodeCoverage]
            public Workbook Workbook { get; set; }
            [ExcludeFromCodeCoverage]
            public NanoXLSX.Utils.Xml.XmlElement XmlElement => null;

            [ExcludeFromCodeCoverage]
            public void Init(IBaseWriter baseWriter)
            {
            }

            [ExcludeFromCodeCoverage]
            public void Execute()
            {
            }
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = ContextualPluginUuid, QueueUUID = ContextualQueueUuid)]
        internal class ContextualQueuePlugin : IPlugin
        {
            internal ContextualQueuePlugin(IBaseWriter baseWriter)
            {
                BaseWriter = baseWriter;
            }

            public IBaseWriter BaseWriter { get; }

            [ExcludeFromCodeCoverage]
            public void Execute()
            {
            }
        }

        internal class BaseWriterStub : IBaseWriter
        {
            public Workbook Workbook { get; } = new Workbook();
            [ExcludeFromCodeCoverage]
            public IWriterProcessingData WriterProcessingData { get; set; }
            [ExcludeFromCodeCoverage]
            public ISharedStringWriter SharedStringWriter { get; set; }
            [ExcludeFromCodeCoverage]
            public void MarkFeatureAsPrepared(string featureUuid)
            {
            }
            [ExcludeFromCodeCoverage]
            public bool IsFeaturePrepared(string featureUuid)
            {
                return false;
            }
        }

    }
}
