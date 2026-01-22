/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;

namespace NanoXLSX.Registry.Attributes
{
    /// <summary>
    /// Attribute to declare a class as general NanoXLSX plug-in
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public class NanoXlsxPlugInAttribute
        : Attribute
    {
        /// <summary>
        /// Unique ID if the plug-in
        /// </summary>
        public string PlugInUUID { get; set; }

        /// <summary>
        /// Order how the annotated plug-ins are registered in case of duplicate UIDs. The higher number will override any lower. 
        /// Default is zero (order may be vary).
        /// </summary>
        public int PlugInOrder { get; set; }
    }
}
