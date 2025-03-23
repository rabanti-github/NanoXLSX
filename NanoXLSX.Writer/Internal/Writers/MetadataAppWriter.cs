/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Interfaces.Writer;
using NanoXLSX.Registry;
using NanoXLSX.Utils.Xml;

namespace NanoXLSX.Internal.Writers
{
    /// <summary>
    /// Class to generate the metadata XML file for the app metadata part on an XLSX file.
    /// </summary>
    [NanoXlsxPlugin(PluginUUID = PluginUUID.METADATA_APP_WRITER)]
    internal class MetadataAppWriter : IPluginWriter
    {
        private XmlElement properties;

        /// <summary>
        /// Gets or replaces the workbook instance, defined by the constructor
        /// </summary>
        public Workbook Workbook { get; set; }
        /// <summary>
        /// relative Package path of the content. This value is not maintained in base plug-ins, but only in appending queue plug-ins
        /// </summary>
        public string PackagePath { get; set; }
        /// <summary>
        /// File name of the content. This value is not maintained in base plug-ins, but only in appending queue plug-ins
        /// </summary>
        public string PackageFileName { get; set; }

        /// <summary>
        /// Default constructor - Must be defined for instantiation of the plug-ins
        /// </summary>
        internal MetadataAppWriter()
        {
        }

        /// <summary>
        /// Initialization method (interface implementation)
        /// </summary>
        /// <param name="baseWriter">Base writer instance that holds any information for this writer</param>
        public void Init(IBaseWriter baseWriter)
        {
            this.Workbook = baseWriter.Workbook;
        }

        /// <summary>
        /// Get the XmlElement after <see cref="Execute"/> (interface implementation)
        /// </summary>
        /// <returns>XmlElement instance that was created after the plug-in execution</returns>
        public XmlElement GetElement()
        {
            Execute();
            return properties;
        }

        /// <summary>
        /// Method to execute the main logic of the plug-in (interface implementation)
        /// </summary>
        public void Execute()
        {
            properties = XmlElement.CreateElement("Properties");
            properties.AddDefaultXmlNameSpace("http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");
            properties.AddNameSpaceAttribute("vt", "xmlns", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            Metadata md = Workbook.WorkbookMetadata;
            properties.AddChildElementWithValue("TotalTime", "0");
            properties.AddChildElementWithValue("Application", md.Application);
            properties.AddChildElementWithValue("DocSecurity", "0");
            properties.AddChildElementWithValue("ScaleCrop", "false");
            properties.AddChildElementWithValue("Manager", md.Manager);
            properties.AddChildElementWithValue("Company", md.Company);
            properties.AddChildElementWithValue("LinksUpToDate", "false");
            properties.AddChildElementWithValue("SharedDoc", "false");
            properties.AddChildElementWithValue("HyperlinkBase", md.HyperlinkBase);
            properties.AddChildElementWithValue("HyperlinksChanged", "false");
            properties.AddChildElementWithValue("AppVersion", md.ApplicationVersion);
        }

    }
}
