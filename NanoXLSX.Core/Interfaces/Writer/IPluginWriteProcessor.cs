/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;

namespace NanoXLSX.Interfaces.Writer
{
    /// <summary>
    /// Interface, used by write processors that do not perform actual XML writing tasks, but performing actions on <see cref="Workbook"/>
    /// </summary>
    internal interface IPluginWriteProcessor : IPlugin
    {

        /// <summary>
        /// Gets or replaces the write context, defined by the constructor
        /// </summary>
        IWriteContext WriteContext { get; set; }

        /// <summary>
        /// Reference to a handler of in-line plugins, to be used for preparing operations in the <see cref="IPlugin.Execute"/> method
        /// </summary>
        Action<IWriteContext, string> InlinePluginHandler { get; set; }

        /// <summary>
        /// Initializing method for the processor
        /// </summary>
        /// <param name="context">Context of the current writing operation.</param>
        /// <param name="inlinePluginHandler">Reference to the a handler action, to be used for preparing operations in processor methods</param>
        void Init(IWriteContext context, Action<IWriteContext, string> inlinePluginHandler);

    }
}
