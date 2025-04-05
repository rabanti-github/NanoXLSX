using System;
using NanoXLSX.Registry;
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


        [Fact(DisplayName = "Test of the plug-in handling initializer (dummy; should not crash or initialite twice)")]
        public void InitializeTest()
        {
            PlugInLoader.DisposePlugins(); // Test on a clean basis
            bool state = PlugInLoader.Initialize();
            Assert.True(state);
            bool state2 = PlugInLoader.Initialize();
            Assert.False(state2);
        }

    }
}
