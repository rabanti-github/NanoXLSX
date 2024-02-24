/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2024
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLS.Shared.Enums.Schemes;
using static NanoXLSX.Shared.Enums.Styles.FontEnums;

namespace NanoXLSX.Shared.Interfaces
{
    /// <summary>
    /// Interface to represent a Font object for styling or formatting
    /// </summary>
    public interface IFont
    {
        bool Bold { get; set; }
        bool Italic { get; set; }
        bool Strike { get; set; }
        UnderlineValue Underline { get; set; }
        CharsetValue Charset { get; set; }
        ThemeEnums.ColorSchemeElement ColorTheme { get; set; }
        string ColorValue { get; set; }
        FontFamilyValue Family { get; set; }
        string Name { get; set; }
        SchemeValue Scheme { get; set; }
        float Size { get; set; }
        VerticalTextAlignValue  VerticalAlign { get; set; }
    }
}
