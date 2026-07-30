using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using NanoXLSX.Interfaces.Writer;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;
using Xunit;

namespace NanoXLSX.Test.Writer_Reader.PlugIns
{
    // Ensure that these tests are executed sequentially, since static repository methods may be called
    [Collection(nameof(SequentialCollection3))]
    public class InlineWriteProcessorPlugInsTest : IDisposable
    {
        private const string FEATURE_UUID = "D68FD790-6620-44F7-91FA-26E58ED6A7BD";

        public void Dispose()
        {
            CompatibilityInlineProcessor.ExecuteAction = null;
            PlugInLoader.DisposePlugins();
        }

        [Fact(DisplayName = "Test marking a writer feature as prepared from an inline processor")]
        public void MarkFeatureAsPreparedTest()
        {
            bool isPrepared = false;
            CompatibilityInlineProcessor.ExecuteAction = processor =>
            {
                processor.MarkFeatureAsPrepared(FEATURE_UUID);
                isPrepared = processor.IsFeaturePrepared(FEATURE_UUID);
            };

            ExecuteWithInjectedProcessor();

            Assert.True(isPrepared);
        }

        [Fact(DisplayName = "Test checking an unprepared writer feature from an inline processor")]
        public void IsFeaturePreparedTest()
        {
            bool isPrepared = true;
            CompatibilityInlineProcessor.ExecuteAction = processor =>
            {
                isPrepared = processor.IsFeaturePrepared(FEATURE_UUID);
            };

            ExecuteWithInjectedProcessor();

            Assert.False(isPrepared);
        }

        [Theory(DisplayName = "Test validation of writer feature UUIDs from an inline processor")]
        [InlineData(true, null)]
        [InlineData(true, "")]
        [InlineData(true, " ")]
        [InlineData(false, null)]
        [InlineData(false, "")]
        [InlineData(false, " ")]
        public void FeatureUuidValidationTest(bool markFeature, string featureUuid)
        {
            ArgumentException validationException = null;
            CompatibilityInlineProcessor.ExecuteAction = processor =>
            {
                Action validationAction = markFeature
                    ? () => processor.MarkFeatureAsPrepared(featureUuid)
                    : () => processor.IsFeaturePrepared(featureUuid);
                validationException = Assert.Throws<ArgumentException>(validationAction);
            };

            ExecuteWithInjectedProcessor();

            Assert.NotNull(validationException);
        }

        private static void ExecuteWithInjectedProcessor()
        {
            PlugInLoader.InjectPlugins(new List<Type> { typeof(CompatibilityInlineProcessor) });
            Workbook workbook = new Workbook("worksheet1");
            using MemoryStream stream = new MemoryStream();
            workbook.SaveAsStream(stream, true);
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "CompatibilityInlineProcessorTest", QueueUUID = PlugInUUID.CompatibilityInlineProcessor)]
        public class CompatibilityInlineProcessor : IPluginInlineWriteProcessor
        {
            private IWriteContext writeContext;

            public static Action<CompatibilityInlineProcessor> ExecuteAction { get; set; }

            [ExcludeFromCodeCoverage]
            IWriteContext IPluginInlineWriteProcessor.WriteContext
            {
                get { return writeContext; }
                set { writeContext = value; }
            }

            public void Execute()
            {
                ExecuteAction?.Invoke(this);
            }

            void IPluginInlineWriteProcessor.Init(IWriteContext context)
            {
                writeContext = context;
            }

            public void MarkFeatureAsPrepared(string featureUuid)
            {
                writeContext.MarkFeatureAsPrepared(featureUuid);
            }

            public bool IsFeaturePrepared(string featureUuid)
            {
                return writeContext.IsFeaturePrepared(featureUuid);
            }
        }
    }
}
