using System;
using System.Collections.Generic;
using System.Text;

namespace NanoXLSX.Registry
{
    [AttributeUsage(AttributeTargets.Class)]
    public class NanoXlsxPluginAttribute : Attribute
    {
        /// <summary>
        /// Unique ID if the plug-in
        /// </summary>
        public string PluginUID { get; set; }

        /// <summary>
        /// Order how the annotated plug-in is loaded. Default is zero (order may be vary).
        /// </summary>
        public int PluginOrder { get; set; } = 0;

    }
}
