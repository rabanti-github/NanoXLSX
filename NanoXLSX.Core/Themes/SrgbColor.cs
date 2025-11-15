/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Collections.Generic;
using NanoXLSX.Interfaces;
using NanoXLSX.Utils;

namespace NanoXLSX.Themes
{

    /// <summary>
    /// Class representing a generic sRGB color without an alpha value
    /// </summary>
    public class SrgbColor : ITypedColor<string>
    {
        private string colorValue;

        /// <summary>
        /// Gets or sets the sRGB value (Hex code of RGB). If set, the value will be cast to upper case
        /// </summary>
        public string ColorValue
        {
            get => colorValue;
            set
            {
                Validators.ValidateColor(value, false);
                colorValue = ParserUtils.ToUpper(value);
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

        /// <summary>
        /// Determines whether the specified object is equal to the current object
        /// </summary>
        /// <param name="obj">Other object to compare</param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            return obj is SrgbColor color &&
                   ColorValue == color.ColorValue;
        }

        /// <summary>
        /// Gets the hash code of the instance
        /// </summary>
        /// <returns>Hash code</returns>
        public override int GetHashCode()
        {
            return 800285905 + EqualityComparer<string>.Default.GetHashCode(ColorValue);
        }
    }
}
