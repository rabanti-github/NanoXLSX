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
    /// Interface to represent an XF entry for styling or formatting
    /// </summary>
    public interface ICellXf
    {
        bool ForceApplyAlignment { get; set; }

        bool Hidden { get; set; }

        HorizontalAlignValue HorizontalAlign { get; set; }

        bool Locked { get; set; }

        TextBreakValue Alignment { get; set; }

        TextDirectionValue TextDirection { get; set; }

        int TextRotation { get; set; }

        VerticalAlignValue VerticalAlign { get; set; }

        int Indent { get; set; }
    }
}
