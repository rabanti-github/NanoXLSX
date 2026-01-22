/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Exceptions;
using NanoXLSX.Interfaces;
using NanoXLSX.Themes;
using NanoXLSX.Utils;

namespace NanoXLSX.Colors
{
    /// <summary>
    /// Class representing a color defined by a theme color scheme element (see <see cref="Theme.ColorSchemeElement"/>)
    /// </summary>
    public class ThemeColor : ITypedColor<Theme.ColorSchemeElement>
    {
        /// <summary>
        /// Gets or sets the color scheme element
        /// </summary>
        public Theme.ColorSchemeElement ColorValue { get; set; }

        /// <summary>
        /// Gets the internal, numeric OOXML string value of the enum, defined in <see cref="ColorValue"/>
        /// </summary>
        public string StringValue
        {
            get
            {
                return ParserUtils.ToString((int)ColorValue);
            }
        }

        /// <summary>
        /// Default constructor
        /// </summary>
        public ThemeColor()
        {
        }

        /// <summary>
        /// Constructor with color scheme element as parameter
        /// </summary>
        /// <param name="color">Color value</param>
        public ThemeColor(Theme.ColorSchemeElement color)
        {

            ColorValue = color;
        }

        /// <summary>
        /// Constructor with index as parameter
        /// </summary>
        /// <param name="index">Theme color index</param>
        /// <exception cref="StyleException">Throws a StyleException if the color scheme element index is out of range</exception>
        public ThemeColor(int index)
        {
            if (index < 0 || index > 11)
            {
                throw new StyleException("Indexed color value must be between 0 and 65.");
            }
            ColorValue = (Theme.ColorSchemeElement)index;
        }


        /// <summary>
        /// Determines whether the specified object is equal to the current system color instance
        /// </summary>
        /// <param name="obj">Other object to compare</param>
        /// <returns>True if both objects are equal</returns>
        public override bool Equals(object obj)
        {
            return obj is ThemeColor color &&
                   ColorValue == color.ColorValue;
        }

        /// <summary>
        /// Gets the hash code of the instance
        /// </summary>
        /// <returns>Hash code</returns>
        public override int GetHashCode()
        {
            return 800285905 + ColorValue.GetHashCode();
        }
    }
}
