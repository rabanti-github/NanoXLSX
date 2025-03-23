/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using NanoXLSX.Interfaces.Writer;
using NanoXLSX.Registry;
using NanoXLSX.Utils;
using NanoXLSX.Utils.Xml;

namespace NanoXLSX.Internal.Writers
{
    /// <summary>
    /// Class to generate the metadata XML file for the app metadata part on an XLSX file.
    /// </summary>
    [NanoXlsxPlugin(PluginUUID = PluginUUID.METADATA_CORE_WRITER)]
    internal class MetadataCoreWriter : IPluginWriter 
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
        internal MetadataCoreWriter()
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
            properties = XmlElement.CreateElement("coreProperties", "cp");
            properties.AddNameSpaceAttribute("dc", "xmlns", "http://purl.org/dc/elements/1.1/");
            properties.AddNameSpaceAttribute("cp", "xmlns", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
            properties.AddNameSpaceAttribute("dcterms", "xmlns", "http://purl.org/dc/terms/");
            properties.AddNameSpaceAttribute("xsi", "xmlns", "http://www.w3.org/2001/XMLSchema-instance");
            Metadata md = Workbook.WorkbookMetadata;
            properties.AddChildElementWithValue("title", md.Title, "dc");
            properties.AddChildElementWithValue("subject", md.Subject, "dc");
            properties.AddChildElementWithValue("creator", md.Creator, "dc");
            properties.AddChildElementWithValue("lastModifiedBy", md.Creator, "cp");
            properties.AddChildElementWithValue("keywords", md.Keywords, "cp");
            properties.AddChildElementWithValue("description", md.Description, "dc");
            string time = DateTime.Now.ToString("yyyy-MM-ddThh:mm:ssZ", ParserUtils.INVARIANT_CULTURE);
            XmlElement child1 = properties.AddChildElementWithValue("created", time, "dcterms");
            child1.AddAttribute("type", "dcterms:W3CDTF", "xsi");
            XmlElement child2 = properties.AddChildElementWithValue("modified", time, "dcterms");
            child2.AddAttribute("type", "dcterms:W3CDTF", "xsi");
            properties.AddChildElementWithValue("category", md.Category, "cp");
            properties.AddChildElementWithValue("contentStatus", md.ContentStatus, "cp");
        }


    }
}
