/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2024
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Globalization;
using System.Text;
using NanoXLSX.Interfaces.Workbook;
using NanoXLSX.Interfaces.Writer;

namespace NanoXLSX.Internal.Writers
{
    internal class MetadataAppWriter : IPluginWriter
    {
        private static readonly CultureInfo CULTURE = CultureInfo.InvariantCulture;

        private readonly Workbook workbook;

        public MetadataAppWriter(XlsxWriter writer)
        {
            this.workbook = writer.Workbook;
        }

        /// <summary>
        /// Method to create the app-properties (part of meta data) as raw XML string
        /// </summary>
        /// \remark <remarks>This method is virtual. Plug-in packages may override it</remarks>
        /// <returns>Raw XML string</returns>
        public virtual string CreateDocument()
        {
            PreWrite(workbook);
            StringBuilder sb = new StringBuilder();
            sb.Append("<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">");
            sb.Append(CreateAppString());
            sb.Append("</Properties>");
            PostWrite(workbook);
            return sb.ToString();
        }

        /// <summary>
        /// Method that is called before the <see cref="CreateDocument()"/> method is executed. 
        /// This virtual method is empty by default and can be overridden by a plug-in package
        /// </summary>
        /// <param name="workbook">Workbook instance that is used in this writer</param>
        public virtual void PreWrite(IWorkbook workbook)
        {
            // NoOp - replaced by plugin
        }

        /// <summary>
        /// Method that is called after the <see cref="CreateDocument()"/> method is executed. 
        /// This virtual method is empty by default and can be overridden by a plug-in package
        /// </summary>
        /// <param name="workbook">Workbook instance that is used in this writer</param>
        public virtual void PostWrite(IWorkbook workbook)
        {
            // NoOp - replaced by plugin
        }

        /// <summary>
        /// Method to create the XML string for the app-properties document
        /// </summary>
        /// <returns>String with formatted XML data</returns>
        private string CreateAppString()
        {
            Metadata md = workbook.WorkbookMetadata;
            StringBuilder sb = new StringBuilder();
            XlsxWriter.AppendXmlTag(sb, "0", "TotalTime", null);
            XlsxWriter.AppendXmlTag(sb, md.Application, "Application", null);
            XlsxWriter.AppendXmlTag(sb, "0", "DocSecurity", null);
            XlsxWriter.AppendXmlTag(sb, "false", "ScaleCrop", null);
            XlsxWriter.AppendXmlTag(sb, md.Manager, "Manager", null);
            XlsxWriter.AppendXmlTag(sb, md.Company, "Company", null);
            XlsxWriter.AppendXmlTag(sb, "false", "LinksUpToDate", null);
            XlsxWriter.AppendXmlTag(sb, "false", "SharedDoc", null);
            XlsxWriter.AppendXmlTag(sb, md.HyperlinkBase, "HyperlinkBase", null);
            XlsxWriter.AppendXmlTag(sb, "false", "HyperlinksChanged", null);
            XlsxWriter.AppendXmlTag(sb, md.ApplicationVersion, "AppVersion", null);
            return sb.ToString();
        }



        /// <summary>
        /// Method to create the XML string for the core-properties document
        /// </summary>
        /// <returns>String with formatted XML data</returns>
        private string CreateCorePropertiesString()
        {
            Metadata md = workbook.WorkbookMetadata;
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
