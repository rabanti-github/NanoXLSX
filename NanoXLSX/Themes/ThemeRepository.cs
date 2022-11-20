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

        /// <summary>
        /// Default theme ID, stated in the workbook document
        /// </summary>
        /// <remarks>According to the official OOXML documentation (part 1, chapter 18.2.28) the version consists of the application version and build where the excel file was created.
        /// The value was extracted from a valid Excel file, created with Excel 2019. However, although '16' can be assumed to be the Version of Excel 2019, 
        /// the build part '6925' cannot be originated, is not reflecting the retrieved application build version, and seems not to be listed publicly
        /// </remarks>
        public const string DEFAULT_THEME_VERSION = "166925";

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

        /// <summary>
        /// Gets the defined theme with the ID of <see cref="DEFAULT_THEME_ID"/> or the <see cref="Theme.GetDefaultTheme"/> if not theme was defined whit that ID
        /// </summary>
        /// <returns></returns>
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
