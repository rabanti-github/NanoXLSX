/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2022
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NanoXLSX.Themes
{
    /// <summary>
    /// Class representing an Office theme
    /// </summary>
    public class Theme
    {
        /// <summary>
        /// Gets or sets the name of the theme
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// Gets or sets the <see cref="ColorScheme"/> of the theme
        /// </summary>
        public ColorScheme Colors { get; set; }
        /// <summary>
        /// Gets or sets the internal ID of the theme. The value is usually identical to <see cref="ThemeRepository.DEFAULT_THEME_ID"/>
        /// </summary>
        public int ID { get; set; }

        /// <summary>
        /// Constructor with parameters
        /// </summary>
        /// <param name="id">Internal ID of the theme</param>
        /// <param name="name">Name of the theme</param>
        public Theme(int id, string name)
        {
            this.ID = id;
            this.Name = name;
        }

        /// <summary>
        /// Gets the default theme if no theme was explicitly defined. This theme will be stored into an XLSX file of not otherwise defined
        /// </summary>
        /// <returns>Theme with default values according to the default theme of Office 2019 (may be deviating)</returns>
        internal static Theme GetDefaultTheme()
        {
            Theme theme = new Theme(ThemeRepository.DEFAULT_THEME_ID, "default");
            ColorScheme colors = new ColorScheme(1);
            colors.Name = "default";
            colors.Dark1 = new SystemColor(SystemColor.Value.WindowText);
            colors.Light1 = new SystemColor(SystemColor.Value.Window, "FFFFFF");
            colors.Dark2 = new SrgbColor("44546A");
            colors.Light2 = new SrgbColor("E7E6E6");
            colors.Accent1 = new SrgbColor("4472C4");
            colors.Accent2 = new SrgbColor("ED7D31");
            colors.Accent3 = new SrgbColor("A5A5A5");
            colors.Accent4 = new SrgbColor("FFC000");
            colors.Accent5 = new SrgbColor("5B9BD5");
            colors.Accent6 = new SrgbColor("70AD47");
            colors.HyperLink = new SrgbColor("0563C1");
            colors.FollowedHyperlink = new SrgbColor("954F72");
            theme.Colors = colors;
            return theme;
        }
    }
}
