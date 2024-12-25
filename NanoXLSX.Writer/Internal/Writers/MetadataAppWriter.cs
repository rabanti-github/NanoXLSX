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
    /// <summary>
    /// Class to generate metadata XML files for the app metadata part on an XLSX file.
    /// </summary>
    internal class MetadataAppWriter : IPluginWriter
    {

        private static readonly CultureInfo CULTURE = CultureInfo.InvariantCulture;

        public MetadataAppWriter(XlsxWriter writer)
        {
            this.Workbook = writer.Workbook;
        }

        /// <summary>
        /// Method to create the app-properties (part of meta data) as raw XML string
        /// </summary>
        /// \remark <remarks>This method is virtual. Plug-in packages may override it</remarks>
        /// <returns>Raw XML string</returns>
        public virtual string CreateDocument()
        {
            PreWrite(Workbook);
            StringBuilder sb = new StringBuilder();
            sb.Append("<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">");
            sb.Append(CreateAppString());
            sb.Append("</Properties>");
            PostWrite(Workbook);
            if (Next != null)
            {
                // TODO this does not work. The string builder is not considered
                Next.Workbook = this.Workbook;
                Next.CreateDocument();
            }
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
        /// Gets the unique class ID. This ID is used to identify the class when replacing functionality by extension packages
        /// </summary>
        /// <returns>GUID of the class</returns>
        public string GetClassID()
        {
            return "A73923A8-1E7E-4673-AD3F-B22DD3153D7B";
        }

        /// <summary>
        /// Gets or replaces the workbook instance, defined by the constructor
        /// </summary>
        public IWorkbook Workbook { get; set; }

        public IPluginWriter Next { get; set; }




        /// <summary>
        /// Method to create the XML string for the app-properties document
        /// </summary>
        /// <returns>String with formatted XML data</returns>
        private string CreateAppString()
        {
            Metadata md = ((Workbook)Workbook).WorkbookMetadata;
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
    }
}
