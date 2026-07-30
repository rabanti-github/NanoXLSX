/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Exceptions;
using NanoXLSX.Interfaces.Writer;
using NanoXLSX.Registry;
using System;
using System.Linq;

namespace NanoXLSX.Internal.Writers
{
    /// <summary>
    /// Processor to check possible compatibility issue before writing a workbook
    /// </summary>
    /// \remark <remarks>This processor should not be overwritten by an external plug-in. 
    /// It ensures that the core functionality does not propagate incompatible (to the default writer) features into the writing process.
    /// External plug-ins can inject a <see cref="IPluginInlineWriteProcessor"/> class with the UUID <see cref="PlugInUUID.CompatibilityInlineProcessor"/> to enable a feature.</remarks>
    internal class CompatibilityProcessor : IPluginWriteProcessor
    {
        /// <summary>
        /// Writer Context
        /// </summary>
        public IWriteContext WriteContext { get; set; }
        /// <summary>
        /// Reference to the <see cref="WriterPlugInHandler"/>, to be used for <b>initial</b>initial in the <see cref="Execute"/> method
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
            // Possible injected compatibility checks or changes (e.g. enabling a feature by a plug-in)
            this.InlinePluginHandler?.Invoke(this.WriteContext, PlugInUUID.CompatibilityInlineProcessor); // DO NOT MOVE THIS TO THE END
            // -------------------------------------------
            CheckExternalLinks();

            // TODO add further compatibility checks of the core scope (in fo plug-ins are loaded) here
        }

        private void CheckExternalLinks()
        {

            bool externalLinksexists = WriteContext.Workbook.GetDefinedNames().Any(x => x.HasExternalReferences);
            if (externalLinksexists && !WriteContext.IsFeaturePrepared(PlugInUUID.WriteExternalLinkFeature))
            {
                throw new NotSupportedContentException("The workbook contains external links in the defined names, but no compatible writer plug-in is capable to write such links. " +
                    "Note: Consider adding the package NanoXLSX.Compatibility. ");
            }

        }
    }
}
