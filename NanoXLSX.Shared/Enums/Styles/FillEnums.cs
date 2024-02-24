/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2024
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Shared.Interfaces;

namespace NanoXLSX.Shared.Enums.Styles
{
    /// <summary>
    /// Class providing shared enums used by implementations of the <see cref="IFill"/> interface // TODO: Define interface
    /// </summary>
    public static class FillEnums
    {
        /// <summary>
        /// Enum for the type of the color
        /// </summary>
        public enum FillType
        {
            /// <summary>Color defines a pattern color </summary>
            patternColor,
            /// <summary>Color defines a solid fill color </summary>
            fillColor,
        }
        /// <summary>
        /// Enum for the pattern values
        /// </summary>
        public enum PatternValue
        {
            /// <summary>
            /// No pattern (default)
            /// </summary>
            /// <remarks>The value none will lead to a invalidation of the foreground or background color values</remarks>
            none,
            /// <summary>Solid fill (for colors)</summary>
            solid,
            /// <summary>Dark gray fill</summary>
            darkGray,
            /// <summary>Medium gray fill</summary>
            mediumGray,
            /// <summary>Light gray fill</summary>
            lightGray,
            /// <summary>6.25% gray fill</summary>
            gray0625,
            /// <summary>12.5% gray fill</summary>
            gray125,
        }
    }
}
