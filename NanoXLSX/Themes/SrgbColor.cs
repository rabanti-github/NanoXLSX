/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2023
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Shared.Interfaces;
using NanoXLSX.Shared.Utils;
using System;
using System.Collections.Generic;

namespace NanoXLSX.Themes
{

    /// <summary>
    /// Class representing a generic sRGB color without an apha value
    /// </summary>
    public class SrgbColor : ITypedColor<string>
    {
        private string colorValue;

        /// <summary>
        /// Gets or sets the sRGB value (Hex code of RGB)
        /// </summary>
        public string ColorValue 
        { get => colorValue;
            set 
            {
                Validators.ValidateColor(value, false);
                colorValue = value; 
            }
        }

        /// <summary>
        /// Gets the string value of the color. The value is identical to <see cref="ColorValue"/> and defined as interface implementation of <see cref="ITypedColor{T}"/>
        /// </summary>
        public string StringValue => colorValue;

        /// <summary>
        /// Default constructor
        /// </summary>
        public SrgbColor()
        {
        }

        /// <summary>
        /// Constructor with parameters
        /// </summary>
        /// <param name="rgb">RGB value</param>
        public SrgbColor(string rgb) : this()
        {
            this.ColorValue = rgb;
        }

        /// <summary>
        /// Converts the sRGB value to an ARGB value
        /// </summary>
        /// <returns>ARGB value with 'FF' as alpha component</returns>
        public string ToArgbColor()
        {
            // Is already validated
            return "FF" + colorValue;
        }

        public override bool Equals(object obj)
        {
            return obj is SrgbColor color &&
                   ColorValue == color.ColorValue;
        }

        public override int GetHashCode()
        {
            return 800285905 + EqualityComparer<string>.Default.GetHashCode(ColorValue);
        }
    }
}
