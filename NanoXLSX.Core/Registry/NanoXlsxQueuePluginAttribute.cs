/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
using System;

namespace NanoXLSX.Registry
{
    /// <summary>
    /// Attribute to declare a class a NanoXLSX plug-in that can be queued (not replacing existing instance with the same UUID)
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public class NanoXlsxQueuePlugInAttribute : Attribute
    {
        /// <summary>
        /// Unique ID if the plug-in
        /// </summary>
        public string PlugInUUID { get; set; }

        /// <summary>
        /// Queue UUID for plug-ins that are not replacing a specific base plug-in, but defined as additional resource, e.g. executed before or after the writer / reader base plug-ins
        /// </summary>
        public string QueueUUID { get; set; } = null;

        /// <summary>
        /// Order how the annotated plug-ins are registered. The higher number will executed after the lower ones in the specified queue. 
        /// Default is zero (order may be vary).
        /// </summary>
        public int PlugInOrder { get; set; } = 0;
    }
}
