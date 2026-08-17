/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

namespace NanoXLSX.Interfaces.Writer
{
    /// <summary>
    /// Interface, used by classes to write XML content iteratively, using <see cref="CurrentIndex"/>, at the end of the XLSX creation process
    /// </summary>
    internal interface IPluginIndexedWriter : IPluginWriter
    {
        /// <summary>
        /// Current index that should be used in <see cref="IPlugin.Execute"/>, to identify the current action to execute.
        /// The index will automatically be set during the iterative execution of the plug-in, from 0 to the <see cref="MaxIndex"/>
        /// </summary>
        int CurrentIndex { get; set; }

        /// <summary>
        /// Current unique index to determine the package part to write. The index must correlate with a prior executed <see cref="IPluginPackageRegistry"/>.
        /// If null, no package part will be written
        /// </summary>
        string CurrentUniquePackagePartIndex { get; }

        /// <summary>
        /// Maximum inclusive (0-based) index for the iterator. A value of -1 indicates that no iterations are required
        /// </summary>
        int MaxIndex { get; }
    }
}
