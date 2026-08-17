using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
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

        private const string QueueUuid = "TEST_PLUGIN_LOADER_QUEUE";
        private const string WriterPluginUuid = "TEST_WRITER_QUEUE_PLUGIN";
        private const string ContextualQueueUuid = "TEST_CONTEXTUAL_PLUGIN_LOADER_QUEUE";
        private const string ContextualPluginUuid = "TEST_CONTEXTUAL_PLUGIN";

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
