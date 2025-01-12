/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Globalization;
using System.Text;
using NanoXLSX.Interfaces;
using NanoXLSX.Interfaces.Writer;

namespace NanoXLSX.Internal.Writers
{
    internal class MetadataCoreWriter : IPluginWriter
    {
        private static readonly CultureInfo CULTURE = CultureInfo.InvariantCulture;

        public MetadataCoreWriter(XlsxWriter writer)
        {
            this.Workbook = writer.Workbook;
        }

        /// <summary>
        /// Method to create the core-properties (part of meta data) as raw XML string
        /// </summary>
        /// \remark <remarks>This method is virtual. Plug-in packages may override it</remarks>
        /// <returns>Raw XML string</returns>
        public virtual string CreateDocument(string currentDocumen = null)
        {
            PreWrite(Workbook);
            StringBuilder sb = new StringBuilder();
            sb.Append("<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">");
            sb.Append(CreateCorePropertiesString());
            sb.Append("</cp:coreProperties>");
            PostWrite(Workbook);
            if (NextWriter != null)
            {
                NextWriter.Workbook = this.Workbook;
                return NextWriter.CreateDocument(sb.ToString());
            }
            else
            {
                return sb.ToString();
            }
        }

        /// <summary>
        /// Method that is called before the <see cref="CreateDocument()"/> method is executed. 
        /// This virtual method is empty by default and can be overridden by a plug-in package
        /// </summary>
        /// <param name="workbook">Workbook instance that is used in this writer</param>
        public virtual void PreWrite(Workbook workbook)
        {
            // NoOp - replaced by plugin
        }

        /// <summary>
        /// Method that is called after the <see cref="CreateDocument()"/> method is executed. 
        /// This virtual method is empty by default and can be overridden by a plug-in package
        /// </summary>
        /// <param name="workbook">Workbook instance that is used in this writer</param>
        public virtual void PostWrite(Workbook workbook)
        {
            // NoOp - replaced by plugin
        }

        /// <summary>
        /// Gets the unique class ID. This ID is used to identify the class when replacing functionality by extension packages
        /// </summary>
        /// <returns>GUID of the class</returns>
        public string GetClassID()
        {
            return "F7494751-5029-4576-9632-FFC2BA1B3E65";
        }

        /// <summary>
        /// Gets or replaces the workbook instance, defined by the constructor
        /// </summary>
        public Workbook Workbook { get; set; }

        /// <summary>
        /// Gets or sets the next plug-in writer. If not null, the next writer to be applied on the document can be called by this property
        /// </summary>
        public IPluginWriter NextWriter { get; set; } = null;


        /// <summary>
        /// Method to create the XML string for the core-properties document
        /// </summary>
        /// <returns>String with formatted XML data</returns>
        private string CreateCorePropertiesString()
        {
            Metadata md = ((Workbook)Workbook).WorkbookMetadata;
            StringBuilder sb = new StringBuilder();
            XlsxWriter.AppendXmlTag(sb, md.Title, "title", "dc");
            XlsxWriter.AppendXmlTag(sb, md.Subject, "subject", "dc");
            XlsxWriter.AppendXmlTag(sb, md.Creator, "creator", "dc");
            XlsxWriter.AppendXmlTag(sb, md.Creator, "lastModifiedBy", "cp");
            XlsxWriter.AppendXmlTag(sb, md.Keywords, "keywords", "cp");
            XlsxWriter.AppendXmlTag(sb, md.Description, "description", "dc");
            string time = DateTime.Now.ToString("yyyy-MM-ddThh:mm:ssZ", CULTURE);
            sb.Append("<dcterms:created xsi:type=\"dcterms:W3CDTF\">").Append(time).Append("</dcterms:created>");
            sb.Append("<dcterms:modified xsi:type=\"dcterms:W3CDTF\">").Append(time).Append("</dcterms:modified>");

            XlsxWriter.AppendXmlTag(sb, md.Category, "category", "cp");
            XlsxWriter.AppendXmlTag(sb, md.ContentStatus, "contentStatus", "cp");

            return sb.ToString();
        }
    }
}
