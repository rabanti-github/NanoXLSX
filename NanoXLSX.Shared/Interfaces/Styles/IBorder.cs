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
    /// Interface to represent a Border object for styling or formatting
    /// </summary>
    public interface IBorder
    {
        string BottomColor { get; set; }

        StyleValue BottomStyle { get; set; }

        string DiagonalColor { get; set; }

        bool DiagonalDown { get; set; }

        bool DiagonalUp { get; set; }

        StyleValue DiagonalStyle { get; set; }

        string LeftColor { get; set; }

        StyleValue LeftStyle { get; set; }

        string RightColor { get; set; }

        StyleValue RightStyle { get; set; }

        string TopColor { get; set; }

        StyleValue TopStyle { get; set; }

    }
}
