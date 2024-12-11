/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2024
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Styles;

namespace NanoXLSX.Shared.Interfaces.Styles
{
    /// <summary>
    /// Interface to represent a Fill object for styling or formatting
    /// </summary>
    public interface IFill
    {
        string BackgroundColor { get; set; }

        string ForegroundColor { get; set; }

        int IndexedColor { get; set; }

        PatternValue PatternFill { get; set; }
    }
}
