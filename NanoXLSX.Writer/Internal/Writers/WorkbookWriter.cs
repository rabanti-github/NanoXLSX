/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Interfaces;
using NanoXLSX.Interfaces.Writer;
using NanoXLSX.Utils;
using NanoXLSX.Themes;
using NanoXLSX.Registry;
using static NanoXLSX.Internal.Enums.Password;
using NanoXLSX.Utils.Xml;

namespace NanoXLSX.Internal.Writers
{
    /// <summary>
    /// Class to generate the workbook XML file in a XLSX file.
    /// </summary>
    [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.WORKBOOK_WRITER)]
    internal class WorkbookWriter : IPlugInWriter
    {
        private XmlElement workbook;
        private IPasswordWriter passwordWriter;

        /// <summary>
        /// Gets or replaces the workbook instance, defined by the constructor
        /// </summary>
        public Workbook Workbook { get; set; }

        /// <summary>
        /// Default constructor - Must be defined for instantiation of the plug-ins
        /// </summary>
        internal WorkbookWriter()
        {
        }

        /// <summary>
        /// Initialization method (interface implementation)
        /// </summary>
        /// <param name="baseWriter">Base writer instance that holds any information for this writer</param>
        public void Init(IBaseWriter baseWriter)
        {
            this.Workbook = baseWriter.Workbook;
            IPassword passwordInstance = Workbook.WorkbookProtectionPassword;
            this.passwordWriter = PlugInLoader.GetPlugIn<IPasswordWriter>(PlugInUUID.PASSWORD_WRITER, new LegacyPasswordWriter());
            this.passwordWriter.Init(PasswordType.WORKBOOK_PROTECTION, passwordInstance.PasswordHash);
        }

        /// <summary>
        /// Get the XmlElement after <see cref="Execute"/> (interface implementation)
        /// </summary>
        /// <returns>XmlElement instance that was created after the plug-in execution</returns>
        public XmlElement GetElement()
        {
            Execute();
            return workbook;
        }

        /// <summary>
        /// Method to execute the main logic of the plug-in (interface implementation)
        /// </summary>
        public void Execute()
        {
            Workbook wb = Workbook;
            workbook = XmlElement.CreateElement("workbook");
            workbook.AddDefaultXmlNameSpace("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            workbook.AddNameSpaceAttribute("r", "xmlns", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            XmlElement workbookPr = workbook.AddChildElement("workbookPr");
            if (wb.WorkbookTheme != null)
            {
                workbookPr.AddAttribute("defaultThemeVersion", Theme.DEFAULT_THEME_VERSION);
                // TODO: add further workbook properties here
            }
            if (wb.SelectedWorksheet > 0 || wb.Hidden)
            {
                XmlElement bookViews = workbook.AddChildElement("bookViews");
                XmlElement workbookView = bookViews.AddChildElement("workbookView");
                if (wb.Hidden)
                {
                    workbookView.AddAttribute("visibility", "hidden");
                }
                else
                {
                    workbookView.AddAttribute("activeTab", ParserUtils.ToString(wb.SelectedWorksheet));
                }
            }
            workbook.AddChildElement(GetWorkbookProtectionElement());
            XmlElement sheets = workbook.AddChildElement("sheets");
            if (wb.Worksheets.Count > 0)
            {
                foreach (Worksheet item in wb.Worksheets)
                {
                    XmlElement sheet = sheets.AddChildElementWithAttribute("sheet", "id", "rId" + ParserUtils.ToString(item.SheetID), "", "r");
                    sheet.AddAttribute("sheetId", item.SheetID.ToString());
                    sheet.AddAttribute("name", XmlUtils.SanitizeXmlValue(item.SheetName));
                    if (item.Hidden)
                    {
                        sheet.AddAttribute("state", "hidden");
                    }
                }
            }
            else
            {
                // Fallback on empty workbook
                XmlElement sheet = sheets.AddChildElementWithAttribute("sheet", "id", "rId1", "", "r");
                sheet.AddAttribute("sheetId", "1");
                sheet.AddAttribute("name", "sheet1");
            }
        }

        /// <summary>
        /// Method to get the workbook protection entries in one XmlElement
        /// </summary>
        /// <returns>XmlElement, holding workbook protection information</returns>
        private XmlElement GetWorkbookProtectionElement()
        {
            Workbook workbook = Workbook;
            XmlElement workbookProtection = null;
            if (workbook.UseWorkbookProtection)
            {
                workbookProtection = XmlElement.CreateElement("workbookProtection");
                if (workbook.LockWindowsIfProtected)
                {
                    workbookProtection.AddAttribute("lockWindows", "1");
                }
                if (workbook.LockStructureIfProtected)
                {
                    workbookProtection.AddAttribute("lockStructure", "1");
                }
                if (passwordWriter.PasswordIsSet())
                {
                    workbookProtection.AddAttributes(passwordWriter.GetAttributes());
                }
            }
            return workbookProtection;
        }

    }
}
