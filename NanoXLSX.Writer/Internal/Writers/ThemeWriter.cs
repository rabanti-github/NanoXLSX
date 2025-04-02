﻿/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Interfaces;
using NanoXLSX.Interfaces.Writer;
using NanoXLSX.Registry;
using NanoXLSX.Themes;
using NanoXLSX.Utils.Xml;

namespace NanoXLSX.Internal.Writers
{
    /// <summary>
    /// Class to generate the theme XML file in a XLSX file.
    /// </summary>
    [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.THEME_WRITER)]
    internal class ThemeWriter : IPlugInWriter
    {
        private XmlElement theme;

        /// <summary>
        /// Gets or replaces the workbook instance, defined by the constructor
        /// </summary>
        public Workbook Workbook { get; set; }

        /// <summary>
        /// Default constructor - Must be defined for instantiation of the plug-ins
        /// </summary>
        internal ThemeWriter()
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
            return theme;
        }

        /// <summary>
        /// Method to execute the main logic of the plug-in (interface implementation)
        /// </summary>
        public void Execute()
        {
            Theme workbookTheme = Workbook.WorkbookTheme;
            theme = XmlElement.CreateElement("theme", "a");
            theme.AddNameSpaceAttribute("a", "xmlns", "http://schemas.openxmlformats.org/drawingml/2006/main");
            theme.AddAttribute("name", XmlUtils.SanitizeXmlValue(workbookTheme.Name));
            XmlElement themeElements = theme.AddChildElement("themeElements", "a");
            themeElements.AddChildElement(GetColorSchemeElement(workbookTheme.Colors));

            WriterPlugInHandler.HandleInlineQueuePlugins(ref themeElements, Workbook, PlugInUUID.THEME_INLINE_WRITER);
        }

        /// <summary>
        /// Method to get all XML elements of a color scheme in one top element
        /// </summary>
        /// <param name="scheme">Color scheme instance</param>
        /// <returns>XmlElement, holding color scheme information (XML)</returns>
        private XmlElement GetColorSchemeElement(ColorScheme scheme)
        {
            XmlElement colorScheme = XmlElement.CreateElementWithAttribute("clrScheme", "name", XmlUtils.SanitizeXmlValue(scheme.Name), "a");
            colorScheme.AddChildElement(GetColor("dk1", scheme.Dark1, "a"));
            colorScheme.AddChildElement(GetColor("lt1", scheme.Light1, "a"));
            colorScheme.AddChildElement(GetColor("dk2", scheme.Dark2, "a"));
            colorScheme.AddChildElement(GetColor("lt2", scheme.Light2, "a"));
            colorScheme.AddChildElement(GetColor("accent1", scheme.Accent1, "a"));
            colorScheme.AddChildElement(GetColor("accent2", scheme.Accent2, "a"));
            colorScheme.AddChildElement(GetColor("accent3", scheme.Accent3, "a"));
            colorScheme.AddChildElement(GetColor("accent4", scheme.Accent4, "a"));
            colorScheme.AddChildElement(GetColor("accent5", scheme.Accent5, "a"));
            colorScheme.AddChildElement(GetColor("accent6", scheme.Accent6, "a"));
            colorScheme.AddChildElement(GetColor("hlink", scheme.Hyperlink, "a"));
            colorScheme.AddChildElement(GetColor("folHlink", scheme.FollowedHyperlink, "a"));
            return colorScheme;
        }

        /// <summary>
        /// Method to determine a single color XML element
        /// </summary>
        /// <param name="name">Name of the element</param>
        /// <param name="color">Color instance</param>
        /// <param name="prefix">Element name prefix</param>
        /// <returns>XmlElement, holding color information</returns>
        private XmlElement GetColor(string name, IColor color, string prefix)
        {
            XmlElement colorElement = XmlElement.CreateElement(name, prefix);
            if (color is SystemColor)
            {
                SystemColor sysColor = color as SystemColor;
                XmlElement sysColorElement = colorElement.AddChildElementWithAttribute("sysClr", "val", sysColor.StringValue, "a");
                if (!string.IsNullOrEmpty(sysColor.LastColor))
                {
                    sysColorElement.AddAttribute("lastClr", sysColor.LastColor);
                }
            }
            else if (color is SrgbColor)
            {
                colorElement.AddChildElementWithAttribute("srgbClr", "val", color.StringValue, "a");
            }
            return colorElement;
        }


    }
}
