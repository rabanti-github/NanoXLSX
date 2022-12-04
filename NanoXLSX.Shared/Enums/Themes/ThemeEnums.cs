/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2022
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Shared.Interfaces;

namespace NanoXLS.Shared.Enums.Schemes
{
    /// <summary>
    /// Class providing shared enums used by theme instances, or as references to parts of such theme instances  
    /// </summary>
    public class ThemeEnums
    {
        /// <summary>
        /// Enum to define the sequence index of color scheme element, used in the implementations of <see cref="IColorScheme"/>
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
    }
}
