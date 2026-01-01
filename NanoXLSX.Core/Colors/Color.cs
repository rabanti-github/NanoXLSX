/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using NanoXLSX.Exceptions;
using NanoXLSX.Interfaces;
using NanoXLSX.Themes;

namespace NanoXLSX.Colors
{
    /// <summary>
    /// Compound class representing a color in various representations (RGB, indexed, theme, system or automatic)
    /// </summary>
    public class Color : IComparable
    {

        #region enums
        /// <summary>
        /// Enum defining the type of color representation
        /// </summary>
        public enum ColorType
        {
            /// <summary>No color defined</summary>
            None,
            /// <summary>Automatic color (determined by application)</summary>
            Auto,
            /// <summary>RGB/ARGB color value</summary>
            Rgb,
            /// <summary>Legacy indexed color (0-56+)</summary>
            Indexed,
            /// <summary>Theme color reference (0-11: dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink)</summary>
            Theme,
            /// <summary>System color (used in themes)</summary>
            System
        }
        #endregion

        #region properties
        /// <summary>
        /// The type of color this value represents
        /// </summary>
        public ColorType Type { get; private set; }

        /// <summary>
        /// Auto attribute - if true, color is automatically determined
        /// </summary>
        public bool? Auto { get; private set; }

        /// <summary>
        /// RGB/ARGB value when Type is Rgb
        /// </summary>
        public SrgbColor RgbColor { get; private set; }

        /// <summary>
        /// Indexed color when Type is Indexed (See <see cref="IndexedColor.Value"/>)
        /// </summary>
        public IndexedColor IndexedColor { get; private set; }

        /// <summary>
        /// Theme-based color when Type is Theme (See <see cref="Theme.ColorSchemeElement"/>)
        /// </summary>
        public ThemeColor ThemeColor { get; private set; }

        /// <summary>
        /// System color when Type is System (See <see cref="SystemColor.Value"/>)
        /// </summary>
        public SystemColor SystemColor { get; private set; }

        /// <summary>
        /// Optional tint value for colors (-1.0 to 1.0)
        /// Positive values lighten, negative values darken
        /// </summary>
        public double? Tint { get; set; }

        /// <summary>
        /// Checks if this Color is defined (not None)
        /// </summary>
        public bool IsDefined => Type != ColorType.None;

        /// <summary>
        /// Gets the color value as IColor interface. If no color was defined (<see cref="ColorType.None"/> or property <see cref="IsDefined"/> is false), null is returned."/>
        /// </summary>
        public IColor Value
        {
            get
            {
                switch (Type)
                {
                    case ColorType.Rgb:
                        return RgbColor;
                    case ColorType.Indexed:
                        return IndexedColor;
                    case ColorType.Theme:
                        return ThemeColor;
                    case ColorType.System:
                        return SystemColor;
                    case ColorType.Auto:
                        return AutoColor.Instance;
                    default:
                        return null;
                }
            }
        }

        #endregion

        #region constructors
        /// <summary>
        /// Private constructor to enforce factory methods
        /// </summary>
        private Color() { }
        #endregion

        #region methods
        /// <summary>
        /// Gets the ARGB string value of the color, if applicable
        /// </summary>
        /// <returns>ARGB value or null, if not applicable</returns>
        /// \remark <remarks>This method only works for colors of Type Rgb or Indexed. For Theme, System or auto colors, the RGB value depends on the actual theme or system settings and cannot be determined here.</remarks>
        public string GetArgbValue()
        {
            if (Type == ColorType.Rgb)
            {
                return RgbColor.ColorValue;
            }
            else if (Type == ColorType.Indexed)
            {
                return IndexedColor.GetArgbValue();
            }
            else
            {
                return null;
            }
        }
        #endregion

        #region factory methods
        /// <summary>
        /// Creates an Color with no color (empty element)
        /// </summary>
        /// <return>Returns a dummy instance of the Color class, where <see cref="ColorType"/> is set to None</return>
        public static Color CreateNone()
        {
            return new Color { Type = ColorType.None };
        }

        /// <summary>
        /// Creates an Color with auto=true
        /// </summary>
        /// <returns>Returns a dummy instance of the Color class, where <see cref="ColorType"/> is set to Auto</returns>
        public static Color CreateAuto()
        {
            return new Color
            {
                Type = ColorType.Auto,
                Auto = true
            };
        }

        /// <summary>
        /// Creates an Color from an RGB/ARGB color
        /// </summary>
        /// <param name="color">Instance of the type <see cref="SrgbColor"/></param>
        /// <returns>Color instance with the value type <see cref="SrgbColor"></see>/></returns>
        public static Color CreateRgb(SrgbColor color)
        {
            // If null is passed, error handling is already covered by the string method
            return new Color
            {
                Type = ColorType.Rgb,
                RgbColor = color
            };
        }

        /// <summary>
        /// Creates an Color from an RGB string (e.g., "FFAABBCC")
        /// </summary>
        /// <param name="rgbValue">RGB or ARGB value as string</param>
        /// <returns>Color instance with the value type <see cref="SrgbColor"></see>/></returns>
        /// <exception cref="StyleException">Throws a StyleException if the passed RGB/ARGB value is invalid</exception>
        public static Color CreateRgb(string rgbValue)
        {
            // Validation is done in SrgbColor class
            return new Color
            {
                Type = ColorType.Rgb,
                RgbColor = new SrgbColor(rgbValue)
            };
        }

        /// <summary>
        /// Creates an Color from an indexed color
        /// </summary>
        /// <param name="color">Instance of the type <see cref="IndexedColor"/></param>
        /// <returns>Color instance with the value type <see cref="IndexedColor"></see>/></returns>
        public static Color CreateIndexed(IndexedColor color)
        {
            if (color == null)
            {
                throw new StyleException("An indexed color cannot be null");
            }
            return new Color
            {
                Type = ColorType.Indexed,
                IndexedColor = color
            };
        }

        /// <summary>
        /// Creates an Color from an indexed color value (see <see cref="IndexedColor.Value"/>)
        /// </summary>
        /// <param name="indexValue">Color index enum value</param>
        /// <returns>Color instance with the value type <see cref="IndexedColor"></see>/></returns>
        public static Color CreateIndexed(IndexedColor.Value indexValue)
        {
            return new Color
            {
                Type = ColorType.Indexed,
                IndexedColor = new IndexedColor(indexValue)
            };
        }

        /// <summary>
        /// Creates an Color from a color index (0 to 65)
        /// </summary>
        /// <param name="index">Color index (0 to 65)</param>
        /// <returns>Color instance with the value type <see cref="IndexedColor"></see>/></returns>
        /// <exception cref="StyleException">Throws a StyleException if the passed index is invalid</exception>
        public static Color CreateIndexed(int index)
        {
            return new Color
            {
                Type = ColorType.Indexed,
                IndexedColor = new IndexedColor(index)
            };
        }

        /// <summary>
        /// Creates an Color from a theme color instance
        /// </summary>
        /// <param name="color">Instance of the type <see cref="ThemeColor"></see></param>
        /// <param name="tint">Optional tint value (from -1 to 1)</param>
        /// <returns>Color instance with the value type <see cref="ThemeColor"></see>/></returns>
        public static Color CreateTheme(ThemeColor color, double? tint = null)
        {
            if (color == null)
            {
                throw new StyleException("A theme color cannot be null");
            }
            return new Color
            {
                Type = ColorType.Theme,
                ThemeColor = color,
                Tint = tint
            };
        }

        /// <summary>
        /// Creates an Color from a theme color scheme element
        /// </summary>
        /// <param name="themeColor">Color scheme element</param>
        /// <param name="tint">Optional tint value (from -1 to 1)</param>
        /// <returns>Color instance with the value type <see cref="ThemeColor"></see>/></returns>
        public static Color CreateTheme(Theme.ColorSchemeElement themeColor, double? tint = null)
        {
            return new Color
            {
                Type = ColorType.Theme,
                ThemeColor = new ThemeColor(themeColor),
                Tint = tint
            };
        }

        /// <summary>
        /// Creates an Color from a system color
        /// </summary>
        /// <param name="color">Instance of the type  <see cref="SystemColor"></see></param>
        /// <returns>Color instance with the value type <see cref="SystemColor"></see></returns>
        public static Color CreateSystem(SystemColor color)
        {
            if (color == null)
            {
                throw new StyleException("A system color cannot be null");
            }
            return new Color
            {
                Type = ColorType.System,
                SystemColor = color
            };
        }

        /// <summary>
        /// Creates an Color from a system color instance
        /// </summary>
        /// <param name="systemColorValue">System color value</param>
        /// <returns>Color instance with the value type <see cref="SystemColor"></see></returns>
        public static Color CreateSystem(SystemColor.Value systemColorValue)
        {
            return new Color
            {
                Type = ColorType.System,
                SystemColor = new SystemColor(systemColorValue)
            };
        }

        #endregion

        /// <summary>
        /// Implicit conversion from (RGB or ARGB) string to Color. This is the most common use case.
        /// </summary>
        /// <param name="rgbValue">RGB or ARGB value</param>
        /// \remark <remarks>The resulting color value will be of the type <see cref="SrgbColor"/>, if valid</remarks>
        public static implicit operator Color(string rgbValue)
        {
            return CreateRgb(rgbValue);
        }

        /// <summary>
        /// Implicit conversion from index number to Color.
        /// </summary>
        /// <param name="colorIndex">Numeric value of the color index (<see cref="IndexedColor.ColorValue"/>)</param>
        /// \remark <remarks>The resulting color value will be of the type <see cref="IndexedColor"/>, if valid</remarks>
        public static implicit operator Color(int colorIndex)
        {
            return CreateIndexed(colorIndex);
        }

        /// <summary>
        /// Implicit conversion from index number to Color.
        /// </summary>
        /// <param name="colorIndex">Index (<see cref="IndexedColor.ColorValue"/>)</param>
        /// \remark <remarks>The resulting color value will be of the type <see cref="IndexedColor"/>, if valid</remarks>
        public static implicit operator Color(IndexedColor.Value colorIndex)
        {
            return CreateIndexed(colorIndex);
        }


        /// <summary>
        /// Determines whether the specified object is equal to the current object
        /// </summary>
        /// <param name="obj">Other object to compare</param>
        /// <returns>True if both objects are equal</returns>
        public override bool Equals(object obj)
        {
            return obj is Color color &&
                   Type == color.Type &&
                   Auto == color.Auto &&
                   EqualityComparer<SrgbColor>.Default.Equals(RgbColor, color.RgbColor) &&
                   EqualityComparer<IndexedColor>.Default.Equals(IndexedColor, color.IndexedColor) &&
                   EqualityComparer<ThemeColor>.Default.Equals(ThemeColor, color.ThemeColor) &&
                   EqualityComparer<SystemColor>.Default.Equals(SystemColor, color.SystemColor) &&
                   Tint == color.Tint &&
                   IsDefined == color.IsDefined;
        }

        /// <summary>
        /// Gets the hash code of the instance
        /// </summary>
        /// <returns>Hash code</returns>
        public override int GetHashCode()
        {
            var hashCode = -1729664991;
            hashCode = hashCode * -1521134295 + Type.GetHashCode();
            hashCode = hashCode * -1521134295 + Auto.GetHashCode();
            hashCode = hashCode * -1521134295 + EqualityComparer<AutoColor>.Default.GetHashCode(AutoColor.Instance);
            hashCode = hashCode * -1521134295 + EqualityComparer<SrgbColor>.Default.GetHashCode(RgbColor);
            hashCode = hashCode * -1521134295 + EqualityComparer<IndexedColor>.Default.GetHashCode(IndexedColor);
            hashCode = hashCode * -1521134295 + EqualityComparer<ThemeColor>.Default.GetHashCode(ThemeColor);
            hashCode = hashCode * -1521134295 + EqualityComparer<SystemColor>.Default.GetHashCode(SystemColor);
            hashCode = hashCode * -1521134295 + Tint.GetHashCode();
            hashCode = hashCode * -1521134295 + IsDefined.GetHashCode();
            return hashCode;
        }

        /// <summary>
        /// Compares two instances for sorting purpose
        /// </summary>
        /// <param name="obj">Object to compare</param>
        /// <returns></returns>
        /// <exception cref="StyleException">Throws a StyleException if the compared object is not from the type Color</exception>
        public int CompareTo(object obj)
        {
            if (obj == null)
            {
                return 1;
            }

            if (!(obj is Color other))
            {
                throw new StyleException("The provided object to compare is not a Color");
            }

            // 1) Compare by color type first
            int typeCompare = Type.CompareTo(other.Type);
            if (typeCompare != 0)
            {
                return typeCompare;
            }

            // 2) Same type -> compare internal representation
            switch (Type)
            {
                case ColorType.None:
                    return 0;
                case ColorType.Auto:
                    return 0;
                case ColorType.Rgb:
                    return string.Compare(
                        RgbColor?.StringValue,
                        other.RgbColor?.StringValue,
                        StringComparison.OrdinalIgnoreCase);
                case ColorType.Indexed:
                    // Numeric comparison of palette index
                    return IndexedColor.ColorValue.CompareTo(other.IndexedColor.ColorValue);
                case ColorType.Theme:
                    {
                        // Numeric comparison of theme slot
                        int themeCompare = ThemeColor.ColorValue.CompareTo(other.ThemeColor.ColorValue);
                        if (themeCompare != 0)
                        {
                            return themeCompare;
                        }
                        // Same theme slot -> compare tint
                        return Nullable.Compare(Tint, other.Tint);
                    }
                case ColorType.System:
                    // Enum-based comparison -> not string-based
                    return SystemColor.ColorValue.CompareTo(other.SystemColor.ColorValue);
                default:
                    // Defensive fallback —> should normally never happen
                    return string.Compare(
                        Value?.StringValue,
                        other.Value?.StringValue,
                        StringComparison.OrdinalIgnoreCase);
            }
        }

    }
}
