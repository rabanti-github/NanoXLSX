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
        /// Indexed color when Type is Indexed
        /// </summary>
        public IndexedColor IndexedColor { get; private set; }

        /// <summary>
        /// Theme-based color when Type is Theme (See <see cref="Theme.ColorSchemeElement"/>)
        /// </summary>
        public ThemeColor ThemeColor { get; private set; }

        /// <summary>
        /// System color when Type is System
        /// </summary>
        public SystemColor SystemColor { get; private set; }

        /// <summary>
        /// Tint value for theme colors (-1.0 to 1.0)
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
        public static Color CreateNone()
        {
            return new Color { Type = ColorType.None };
        }

        /// <summary>
        /// Creates an Color with auto=true
        /// </summary>
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
        public static Color CreateRgb(SrgbColor color)
        {
            if (color == null)
            {
                throw new StyleException("sRGB value cannot be null");
            }

            return new Color
            {
                Type = ColorType.Rgb,
                RgbColor = color
            };
        }

        /// <summary>
        /// Creates an Color from an RGB string (e.g., "FFAABBCC")
        /// </summary>
        public static Color CreateRgb(string rgbValue)
        {
            if (string.IsNullOrEmpty(rgbValue))
            {
                throw new StyleException("RGB value cannot be null or empty");
            }

            return new Color
            {
                Type = ColorType.Rgb,
                RgbColor = new SrgbColor(rgbValue)
            };
        }

        /// <summary>
        /// Creates an Color from an indexed color
        /// </summary>
        public static Color CreateIndexed(IndexedColor color)
        {
            return new Color
            {
                Type = ColorType.Indexed,
                IndexedColor = color
            };
        }

        /// <summary>
        /// Creates an Color from an indexed color value (see <see cref="IndexedColor.Value"/>)
        /// </summary>
        public static Color CreateIndexed(IndexedColor.Value enumValue)
        {
            return new Color
            {
                Type = ColorType.Indexed,
                IndexedColor = new IndexedColor(enumValue)
            };
        }

        /// <summary>
        /// Creates an Color from a color index (0 to 65)
        /// </summary>
        public static Color CreateIndexed(int index)
        {
            return new Color
            {
                Type = ColorType.Indexed,
                IndexedColor = new IndexedColor(index)
            };
        }


        /// <summary>
        /// Creates an Color from a theme color scheme element
        /// </summary>
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

        #endregion

        /// <summary>
        /// Implicit conversion from (RGB or ARGB) string to Color. This is the most common use case.
        /// </summary>
        /// <param name="rgbValue">RGB or ARGB value</param>
        public static implicit operator Color(string rgbValue)
        {
            return CreateRgb(rgbValue);
        }

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
                case ColorType.Auto: // Auto has no value
                    return 0;
                case ColorType.Rgb:
                    return string.Compare(
                        RgbColor?.StringValue,
                        other.RgbColor?.StringValue,
                        StringComparison.OrdinalIgnoreCase);
                case ColorType.Indexed:
                    return string.Compare(
                        IndexedColor?.StringValue,
                        other.IndexedColor?.StringValue,
                        StringComparison.Ordinal);
                case ColorType.Theme:
                    {
                        int themeCompare = string.Compare(
                            ThemeColor?.StringValue,
                            other.ThemeColor?.StringValue,
                            StringComparison.Ordinal);

                        if (themeCompare != 0)
                        {
                            return themeCompare;
                        }
                        // Same theme index -> compare tint
                        return Nullable.Compare(Tint, other.Tint);
                    }
                case ColorType.System:
                    return string.Compare(
                        SystemColor?.StringValue,
                        other.SystemColor?.StringValue,
                        StringComparison.OrdinalIgnoreCase);
                default: // Defensive fallback
                    return string.Compare(
                        Value?.StringValue,
                        other.Value?.StringValue,
                        StringComparison.OrdinalIgnoreCase);
            }
        }

    }
}
