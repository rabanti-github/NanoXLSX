/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2022
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Collections.Generic;

namespace NanoXLSX.Themes
{
    /// <summary>
    /// Class to manage all themes of a workbook
    /// </summary>
    /// <remarks>Themes are currently not really used in the library when it comes to data referencing. However, IDs of a theme may be referenced in Font style components</remarks>
    public class ThemeRepository
    {
        /// <summary>
        /// ID if the default theme if one or more are defined in a workbook. Use <see cref="UndefinedTheme"/> if no theme is defined
        /// </summary>
        public const int DEFAULT_THEME_ID = 1;

        public static Theme UndefinedTheme { get; } = Theme.GetDefaultTheme();

        public Dictionary<int, Theme> Themes { get; } = new Dictionary<int, Theme>();

        private static ThemeRepository instance;

        /// <summary>
        /// Gets the singleton instance of the repository
        /// </summary>
        public static ThemeRepository Instance
        {
            get
            {
                instance = instance ?? new ThemeRepository();
                return instance;
            }
        }

        public static Theme GetThemeOrDefault()
        {
            if (Instance.Themes.ContainsKey(DEFAULT_THEME_ID))
            {
                return Instance.Themes[DEFAULT_THEME_ID];
            }
            else
            {
                return UndefinedTheme;
            }
        }
    
    }
}
