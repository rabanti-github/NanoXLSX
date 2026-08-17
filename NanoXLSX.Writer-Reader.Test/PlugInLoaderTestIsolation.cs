using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using NanoXLSX.Registry;

namespace NanoXLSX.Test.Writer_Reader
{
    /// <summary>
    /// Prevents attributed test fixtures from being discovered as production plug-ins outside their explicit tests
    /// </summary>
    internal static class PlugInLoaderTestIsolation
    {
        [ModuleInitializer]
        internal static void Initialize()
        {
            Reset();
        }

        internal static void Reset()
        {
            PlugInLoader.DisposePlugins();
            PlugInLoader.InjectPlugins(new List<Type>());
        }
    }
}
