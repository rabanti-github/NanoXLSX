/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Text;
using NanoXLSX.Interfaces;
using NanoXLSX.Interfaces.Writer;
using NanoXLSX.Exceptions;
using NanoXLSX.Utils;
using NanoXLSX.Themes;

namespace NanoXLSX.Internal.Writers
{
    internal class WorkbookWriter : IPluginWriter
    {
        private Workbook workbookInstance;
        private readonly IPasswordWriter passwordWriter;

        public WorkbookWriter(XlsxWriter writer)
        {
            this.workbookInstance = writer.Workbook;
            IPassword passwordInstance = ((Workbook)workbookInstance).WorkbookProtectionPassword;
            //TODO add plugin hook to overwrite password instance
            this.passwordWriter = new LegacyPasswordWriter(LegacyPasswordWriter.PasswordType.WORKBOOK_PROTECTION, passwordInstance.PasswordHash);
        }

        /// <summary>
        /// Method to create a workbook as raw XML string
        /// </summary>
        /// \remark <remarks>This method is virtual. Plug-in packages may override it</remarks>
        /// <returns>Raw XML string</returns>
        /// <exception cref="RangeException">Throws a RangeException if an address was out of range</exception>
        public virtual string CreateDocument(string currentDocument = null)
        {
            Workbook workbook = (Workbook)workbookInstance;
            PreWrite(workbook);
            StringBuilder sb = new StringBuilder();
            sb.Append("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");

            if (workbook.WorkbookTheme == null || workbook.WorkbookTheme.DefaultTheme)
            {
                sb.Append("<workbookPr defaultThemeVersion=\"");
                sb.Append(Theme.DEFAULT_THEME_VERSION);
                // TODO: add further workbook properties here
                sb.Append("\"/>");
            }
            else
            {
                sb.Append("<workbookPr />");
            }
            if (workbook.SelectedWorksheet > 0 || workbook.Hidden)
            {
                sb.Append("<bookViews><workbookView ");
                if (workbook.Hidden)
                {
                    sb.Append("visibility=\"hidden\"");
                }
                else
                {
                    sb.Append("activeTab=\"").Append(ParserUtils.ToString(workbook.SelectedWorksheet)).Append("\"");
                }
                sb.Append("/></bookViews>");
            }
            CreateWorkbookProtectionString(sb);
            sb.Append("<sheets>");
            if (workbook.Worksheets.Count > 0)
            {
                foreach (Worksheet item in workbook.Worksheets)
                {
                    sb.Append("<sheet r:id=\"rId").Append(item.SheetID.ToString()).Append("\" sheetId=\"").Append(item.SheetID.ToString()).Append("\" name=\"").Append(XmlUtils.EscapeXmlAttributeChars(item.SheetName)).Append("\"");
                    if (item.Hidden)
                    {
                        sb.Append(" state=\"hidden\"");
                    }
                    sb.Append("/>");
                }
            }
            else
            {
                // Fallback on empty workbook
                sb.Append("<sheet r:id=\"rId1\" sheetId=\"1\" name=\"sheet1\"/>");
            }
            sb.Append("</sheets>");
            sb.Append("</workbook>");
            PostWrite(workbook);
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
            return "1F1A8E6D-F373-40DC-9959-46299DA8EAAD";
        }

        /// <summary>
        /// Gets or replaces the workbook instance, defined by the constructor
        /// </summary>
        public Workbook Workbook
        {
            get { return workbookInstance; }
            set { workbookInstance = value; }
        }

        /// <summary>
        /// Gets or sets the next plug-in writer. If not null, the next writer to be applied on the document can be called by this property
        /// </summary>
        public IPluginWriter NextWriter { get; set; } = null;

        /// <summary>
        /// Method to create the (sub) part of the workbook protection within the workbook XML document
        /// </summary>
        /// <param name="sb">reference to the StringBuilder</param>
        private void CreateWorkbookProtectionString(StringBuilder sb)
        {
            Workbook workbook = (Workbook)workbookInstance;
            if (workbook.UseWorkbookProtection)
            {
                sb.Append("<workbookProtection");
                if (workbook.LockWindowsIfProtected)
                {
                    sb.Append(" lockWindows=\"1\"");
                }
                if (workbook.LockStructureIfProtected)
                {
                    sb.Append(" lockStructure=\"1\"");
                }
                if (passwordWriter.PasswordIsSet())
                {
                    sb.Append(passwordWriter.GetXmlAttributes());
                }
                sb.Append("/>");
            }
        }
    }
}
