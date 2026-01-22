/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Collections.Generic;
using NanoXLSX.Interfaces;
using NanoXLSX.Utils;

namespace NanoXLSX.Colors
{

    /// <summary>
    /// Class representing a generic sRGB color (with or without alpha channel)
    /// </summary>
    public class SrgbColor : ITypedColor<string>
    {
        #region constants
        /// <summary>
        /// Default color value (opaque black: #000000)
        /// </summary>
        public const string DefaultSrgbColor = "FF000000";
        #endregion

        private string colorValue;

        /// <summary>
        /// Gets or sets the sRGB value (Hex code of RGB/ARGB). If set, the value will be cast to upper case.
        /// If a 6-character RGB value is provided, 'FF' is automatically prepended as alpha channel.
        /// </summary>
        public string ColorValue
        {
            get => colorValue;
            set
            {
                Validators.ValidateGenericColor(value, false);
                if (value.Length == 6)
                {
                    colorValue = "FF" + ParserUtils.ToUpper(value);
                }
                else
                {
                    colorValue = ParserUtils.ToUpper(value);
                }
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
        /// <param name="rgb">RGB/ARGB value</param>
        public SrgbColor(string rgb) : this()
        {
            ColorValue = rgb;
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
