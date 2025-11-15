/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Collections.Generic;

namespace NanoXLSX.Themes
{
    /// <summary>
    /// Class representing an Office theme
    /// </summary>
    public class Theme
    {

        #region constants
        /// <summary>
        /// Default theme ID, stated in the workbook document
        /// </summary>
        /// \remark <remarks>According to the official OOXML documentation (part 1, chapter 18.2.28) the version consists of the application version and build where the excel file was created.
        /// The value was extracted from a valid Excel file, created with Excel 2019. However, although '16' can be assumed to be the Version of Excel 2019, 
        /// the build part '6925' cannot be originated, is not reflecting the retrieved application build version, and seems not to be listed publicly
        /// </remarks>
        public const string DEFAULT_THEME_VERSION = "166925";

        #endregion

        #region enums
        /// <summary>
        /// Enum to define the sequence index of color scheme element, used in the implementations of <see cref="Interfaces.IColorScheme"/>
        /// </summary>
        public enum ColorSchemeElement
        {
            /// <summary>Dark 1</summary>
            dark1 = 0,
            /// <summary>Light 1</summary>
            light1 = 1,
            /// <summary>Dark 2</summary>
            dark2 = 2,
            /// <summary>Light 2</summary>
            light2 = 3,
            /// <summary>Accent 1</summary>
            accent1 = 4,
            /// <summary>Accent 2</summary>
            accent2 = 5,
            /// <summary>Accent 3</summary>
            accent3 = 6,
            /// <summary>Accent 4</summary>
            accent4 = 7,
            /// <summary>Accent 5</summary>
            accent5 = 8,
            /// <summary>Accent 6</summary>
            accent6 = 9,
            /// <summary>Hyperlink</summary>
            hyperlink = 10,
            /// <summary>Followed Hyperlink</summary>
            followedHyperlink = 11
        }
        #endregion

        /// <summary>
        /// Gets or sets the name of the theme
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// Gets or sets the <see cref="ColorScheme"/> of the theme
        /// </summary>
        public ColorScheme Colors { get; set; }

        /// <summary>
        /// Gets whether the theme is defined as copy or reference to the application default theme.
        /// </summary>
        /// \remark <remarks>This indication and the default theme (<see cref="Theme.GetDefaultTheme"/>) may still deviate from the actual default theme defined by the handling application (e.g. Excel)- The property is not written into Excel files</remarks>
        public bool DefaultTheme { get; private set; }

        /// <summary>
        /// Constructor with parameters. Using this constructor initialized the <see cref="Colors"/> property with valid default values
        /// </summary>
        /// <param name="name">Name of the theme</param>
        public Theme(string name)
        {
            this.Name = name;
            this.Colors = GetDefaultColorScheme();
        }


        /// <summary>
        /// Gets the default theme if no theme was explicitly defined. This theme will be stored into an XLSX file if not otherwise defined
        /// </summary>
        /// <returns>Theme with default values according to the default theme of Office 2019 (may be deviating)</returns>
        internal static Theme GetDefaultTheme()
        {
            Theme theme = new Theme("default");
            theme.DefaultTheme = true;
            ColorScheme colors = GetDefaultColorScheme();
            theme.Colors = colors;
            return theme;
        }

        /// <summary>
        /// Gets the default color scheme if no scheme was explicitly defined. This theme will be incorporated into the default theme of an XLSX file if not otherwise defined 
        /// </summary>
        /// <returns>Color scheme with default values according to the default color scheme of Office 2019 (may be deviating)</returns>
        internal static ColorScheme GetDefaultColorScheme()
        {
            ColorScheme colors = new ColorScheme();
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
            colors.Hyperlink = new SrgbColor("0563C1");
            colors.FollowedHyperlink = new SrgbColor("954F72");
            return colors;
        }

        /// <summary>
        /// Returns whether two instances are the same
        /// </summary>
        /// <param name="obj">Object to compare</param>
        /// <returns>True if this instance and the other are the same</returns>
        public override bool Equals(object obj)
        {
            return obj is Theme theme &&
                   Name == theme.Name &&
                   EqualityComparer<ColorScheme>.Default.Equals(Colors, theme.Colors);
        }

        /// <summary>
        /// Returns a hash code for this instance.
        /// </summary>
        /// <returns>
        /// A hash code for this instance, suitable to be used in hashing algorithms and data structures like a hash table. 
        /// </returns>
        public override int GetHashCode()
        {
            unchecked
            {
                int hashCode = 1172093127;
                hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(Name);
                hashCode = hashCode * -1521134295 + EqualityComparer<ColorScheme>.Default.GetHashCode(Colors);
                return hashCode;
            }
        }
    }
}
