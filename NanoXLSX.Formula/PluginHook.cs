using System;
using System.Collections.Generic;
using System.Text;
using NanoXLSX.Extensions;
using NanoXLSX.Interfaces;
using NanoXLSX.Registry;

namespace NanoXLSX.Extensions.Formula
{
    [NanoXlsxPlugin(PluginUID = "A5AA8E89-3C4E-4ECE-84DB-BC27198A7819")]
    public class PluginHook : IPluginHook
    {
        public PluginHook()
        {

        }

        public void Register()
        {
            Test instance = new Test();
            PackageRegistry.RegisterWriterPlugin(instance.GetClassID(), instance);
        }
    }
}
