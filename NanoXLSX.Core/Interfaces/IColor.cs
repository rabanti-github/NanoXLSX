/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

namespace NanoXLSX.Interfaces
{
    /// <summary>
    /// Interface to represent non typed color, either defined by the system or the user
    /// </summary>
    public interface IColor
    {
        /// <summary>
        /// Color value. This may be a system defined name or a value like an ARGB value
        /// </summary>
        string StringValue { get; }
    }
}
