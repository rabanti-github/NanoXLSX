﻿/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2024
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Text;
using NanoXLSX.Interfaces.Workbook;
using NanoXLSX.Interfaces.Writer;
using NanoXLSX.Shared.Interfaces;
using NanoXLSX.Shared.Utils;
using NanoXLSX.Themes;

namespace NanoXLSX.Internal.Writers
{
    internal class ThemeWriter : IPluginWriter
    {

        private readonly Workbook workbook;

        internal ThemeWriter(XlsxWriter writer)
        {
            this.workbook = writer.Workbook;
        }

        public virtual string CreateDocument()
        {
            PreWrite(workbook);
            Theme theme = workbook.WorkbookTheme;
            StringBuilder sb = new StringBuilder();
            sb.Append("<theme xmlns=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"").Append(XmlUtils.EscapeXmlAttributeChars(theme.Name)).Append("\">");
            sb.Append("<themeElements>");
            CreateColorSchemeString(sb, theme.Colors);
            sb.Append("</themeElements>");
            sb.Append("</theme>");
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

        private void CreateColorSchemeString(StringBuilder sb, ColorScheme scheme)
        {
            sb.Append("<clrScheme name=\"").Append(XmlUtils.EscapeXmlAttributeChars(scheme.Name)).Append("\">");
            ParseColor(sb, "dk1", scheme.Dark1);
            ParseColor(sb, "lt1", scheme.Light1);
            ParseColor(sb, "dk2", scheme.Dark2);
            ParseColor(sb, "lt2", scheme.Light2);
            ParseColor(sb, "accent1", scheme.Accent1);
            ParseColor(sb, "accent2", scheme.Accent2);
            ParseColor(sb, "accent3", scheme.Accent3);
            ParseColor(sb, "accent4", scheme.Accent4);
            ParseColor(sb, "accent5", scheme.Accent5);
            ParseColor(sb, "accent6", scheme.Accent6);
            ParseColor(sb, "hlink", scheme.Hyperlink);
            ParseColor(sb, "folHlink", scheme.FollowedHyperlink);
            sb.Append("</clrScheme>");
        }

        private void ParseColor(StringBuilder sb, string name, IColor color)
        {
            sb.Append("<").Append(name).Append(">");
            if (color is SystemColor)
            {
                SystemColor sysColor = color as SystemColor;
                sb.Append("<sysClr val=\"").Append(sysColor.StringValue).Append("\" ");
                if (!string.IsNullOrEmpty(sysColor.LastColor))
                {
                    sb.Append("lastClr=\"").Append(sysColor.LastColor).Append("\" ");
                }
                sb.Append("/>");
            }
            else if (color is SrgbColor)
            {
                sb.Append("<srgbClr val=\"").Append(color.StringValue).Append("\" />");
            }
            sb.Append("</").Append(name).Append(">");
        }

    }
}
