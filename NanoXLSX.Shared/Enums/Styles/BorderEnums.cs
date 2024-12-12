/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2024
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Shared.Interfaces.Styles;

namespace NanoXLSX.Styles
{

    /// <summary>
    /// Enum for the border style, used by implementations of the <see cref="IBorder"/> interface
    /// </summary>
    public enum StyleValue
    {
        /// <summary>no border</summary>
        none,
        /// <summary>hair border</summary>
        hair,
        /// <summary>dotted border</summary>
        dotted,
        /// <summary>dashed border with double-dots</summary>
        dashDotDot,
        /// <summary>dash-dotted border</summary>
        dashDot,
        /// <summary>dashed border</summary>
        dashed,
        /// <summary>thin border</summary>
        thin,
        /// <summary>medium-dashed border with double-dots</summary>
        mediumDashDotDot,
        /// <summary>slant dash-dotted border</summary>
        slantDashDot,
        /// <summary>medium dash-dotted border</summary>
        mediumDashDot,
        /// <summary>medium dashed border</summary>
        mediumDashed,
        /// <summary>medium border</summary>
        medium,
        /// <summary>thick border</summary>
        thick,
        /// <summary>double border</summary>
        s_double,
    }
}
