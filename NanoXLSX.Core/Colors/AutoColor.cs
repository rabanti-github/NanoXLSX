/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Interfaces;

namespace NanoXLSX.Colors
{
    /// <summary>
    /// Class representing an automatic color. 
    /// </summary>
    /// \remark <remarks>This class does not carry any value. It is only for the purpose of identification.</remarks>
    public class AutoColor : IColor
    {
        /// <summary>
        /// Static instance of the AutoColor class to avoid multiple instances (instances does not deviate)
        /// </summary>
        public static readonly AutoColor Instance = new AutoColor();

        /// <summary>
        /// The string value of an auto color is always null
        /// </summary>
        public string StringValue => null;
    }
}
