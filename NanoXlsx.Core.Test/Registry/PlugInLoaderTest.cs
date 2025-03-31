using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NanoXLSX.Interfaces.Writer;
using NanoXLSX.Internal.Structures;
using NanoXLSX.Registry;
using NanoXLSX.Test.Core.Utils;

using NanoXLSX.Utils.Xml;
using Xunit;
using Xunit.Sdk;

namespace NanoXLSX.Test.Writer_Reader.PlugInsTest
{
    // Ensure that these tests are executed sequentially (in a own collection), since static repository methods may be called 
    [Collection(nameof(SequentialPlugInCollection))]
    public class PluginLoaderTest
    {

        [Fact(DisplayName = "Test of the plug-in handling initializer (dummy; should not crash or initialite twice)")]
        public void InitializeTest()
        {
            bool state = PlugInLoader.Initialize();
            Assert.True(state);
            bool state2 = PlugInLoader.Initialize();
            Assert.False(state2);
        }

    }
}
