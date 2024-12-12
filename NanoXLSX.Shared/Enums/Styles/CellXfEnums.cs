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
    /// Enum for the horizontal alignment of a cell, used by implementations of the <see cref="ICellXf"/>
    /// </summary>
    public enum HorizontalAlignValue
    {
        /// <summary>Content will be aligned left</summary>
        left,
        /// <summary>Content will be aligned in the center</summary>
        center,
        /// <summary>Content will be aligned right</summary>
        right,
        /// <summary>Content will fill up the cell</summary>
        fill,
        /// <summary>justify alignment</summary>
        justify,
        /// <summary>General alignment</summary>
        general,
        /// <summary>Center continuous alignment</summary>
        centerContinuous,
        /// <summary>Distributed alignment</summary>
        distributed,
        /// <summary>No alignment. The alignment will not be used in a style</summary>
        none,
    }

    /// <summary>
    /// Enum for text break options, used by implementations of the <see cref="ICellXf"/>
    /// </summary>
    public enum TextBreakValue
    {
        /// <summary>Word wrap is active</summary>
        wrapText,
        /// <summary>Text will be resized to fit the cell</summary>
        shrinkToFit,
        /// <summary>Text will overflow in cell</summary>
        none,
    }

    /// <summary>
    /// Enum for the general text alignment direction, used by implementations of the <see cref="ICellXf"/>
    /// </summary>
    public enum TextDirectionValue
    {
        /// <summary>Text direction is horizontal (default)</summary>
        horizontal,
        /// <summary>Text direction is vertical</summary>
        vertical,
    }

    /// <summary>
    /// Enum for the vertical alignment of a cell, used by implementations of the <see cref="ICellXf"/>
    /// </summary>
    public enum VerticalAlignValue
    {
        /// <summary>Content will be aligned on the bottom (default)</summary>
        bottom,
        /// <summary>Content will be aligned on the top</summary>
        top,
        /// <summary>Content will be aligned in the center</summary>
        center,
        /// <summary>justify alignment</summary>
        justify,
        /// <summary>Distributed alignment</summary>
        distributed,
        /// <summary>No alignment. The alignment will not be used in a style</summary>
        none,
    }
}
