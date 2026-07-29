/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Interfaces.Writer;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;
using NanoXLSX.Styles;
using System;

namespace NanoXLSX.Internal.Writers
{

    /// <summary>
    /// Class to check and prepare the workbook before writing.
    /// </summary>
    [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.PreparingProcessor)]
    internal class PreparingProcessor : IPluginWriteProcessor
    {
        /// <summary>
        /// Writer context
        /// </summary>
        public IWriteContext WriteContext { get; set; }
        /// <summary>
        /// Reference to the <see cref="WriterPlugInHandler"/>, to be used for post operations in the <see cref="Execute"/> method
        /// </summary>
        public Action<IWriteContext, string> InlinePluginHandler { get; set; }

        /// <summary>
        /// Initialization method (interface implementation)
        /// </summary>
        /// <param name="context">Writer context</param>
        /// <param name="inlinePluginHandler">Action reference</param>
        public void Init(IWriteContext context, Action<IWriteContext, string> inlinePluginHandler)
        {
            this.WriteContext = context;
            this.InlinePluginHandler = inlinePluginHandler;
        }

        /// <summary>
        /// Method to execute the main logic of the plug-in (interface implementation)
        /// </summary>
        public void Execute()
        {
            WriteContext.Workbook.ResolveMergedCells();
            WriteContext.WriterProcessingData.StyleManager = StyleManager.GetManagedStyles(WriteContext.Workbook);

            this.InlinePluginHandler?.Invoke(this.WriteContext, PlugInUUID.PreparingInlineProcessor);
        }

    }
}
