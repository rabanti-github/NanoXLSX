/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using NanoXLSX.Colors;
using NanoXLSX.Interfaces;

namespace NanoXLSX.Styles
{
    /// <summary>
    /// Class representing a Fill (background) entry. The Fill entry is used to define background colors and fill patterns
    /// </summary>
    public class Fill : AbstractStyle
    {
        #region constants
        /// <summary>
        /// Default Color (foreground or background)
        /// </summary>
        public static readonly Color DefaultColor = Color.CreateRgb(SrgbColor.DefaultSrgbColor);
        /// <summary>
        /// Default index color
        /// </summary>
        public static readonly Color DefaultIndexedColor = Color.CreateIndexed(IndexedColor.DefaultIndexedColor);
        /// <summary>
        /// Default pattern
        /// </summary>
        public static readonly PatternValue DefaultPatternFill = PatternValue.None;

        #endregion

        #region enums
        /// <summary>
        /// Enum for the type of the color, used by the <see cref="Fill"/> class
        /// </summary>
        public enum FillType
        {
            /// <summary>Color defines a pattern color </summary>
            PatternColor,
            /// <summary>Color defines a solid fill color </summary>
            FillColor,
        }

        /// <summary>
        /// Enum for the pattern values, used by the <see cref="Fill"/> class
        /// </summary>
        public enum PatternValue
        {
            /// <summary>
            /// No pattern (default)
            /// </summary>
            /// \remark <remarks>The value none will lead to a invalidation of the foreground or background color values</remarks>
            None,
            /// <summary>Solid fill (for colors)</summary>
            Solid,
            /// <summary>Dark gray fill</summary>
            DarkGray,
            /// <summary>Medium gray fill</summary>
            MediumGray,
            /// <summary>Light gray fill</summary>
            LightGray,
            /// <summary>6.25% gray fill</summary>
            Gray0625,
            /// <summary>12.5% gray fill</summary>
            Gray125,
        }
        #endregion

        #region privateFields
        private Color backgroundColor = DefaultColor;
        private Color foregroundColor = DefaultColor;
        #endregion

        #region properties
        /// <summary>
        /// Gets or sets the background color of the fill. The value is expressed as hex string with the format AARRGGBB. AA (Alpha) is usually FF. If set, the value will be cast to upper case
        /// </summary>
        /// \remark <remarks>If a background color is set and the <see cref="PatternFill">PatternFill</see> Property is currently set to <see cref="PatternValue.None">PatternValue.none</see>, 
        /// the PatternFill property will be automatically set to <see cref="PatternValue.Solid">PatternValue.solid</see>, since none invalidates the color values of the foreground or background</remarks>
        [Append]
        public Color BackgroundColor
        {
            get => backgroundColor;
            set
            {
                backgroundColor = value;
                if (PatternFill == PatternValue.None)
                {
                    PatternFill = PatternValue.Solid;
                }
            }
        }
        /// <summary>
        /// Gets or sets the foreground color of the fill. The value is expressed as hex string with the format AARRGGBB. AA (Alpha) is usually FF. If set, the value will be bast to upper case
        /// </summary>
        /// \remark <remarks>If a foreground color is set and the <see cref="PatternFill">PatternFill</see> Property is currently set to <see cref="PatternValue.None">PatternValue.none</see>, 
        /// the PatternFill property will be automatically set to <see cref="PatternValue.Solid">PatternValue.solid</see>, since none invalidates the color values of the foreground or background</remarks>
        [Append]
        public Color ForegroundColor
        {
            get => foregroundColor;
            set
            {
                foregroundColor = value;
                if (PatternFill == PatternValue.None)
                {
                    PatternFill = PatternValue.Solid;
                }
            }
        }
        /// <summary>
        /// Gets or sets the pattern type of the fill (Default is none)
        /// </summary>
        [Append]
        public PatternValue PatternFill { get; set; }
        #endregion

        #region constructors
        /// <summary>
        /// Default constructor
        /// </summary>
        public Fill()
        {
            PatternFill = DefaultPatternFill;
            foregroundColor = DefaultColor;
            backgroundColor = DefaultColor;
        }
        /// <summary>
        /// Constructor with foreground and background color as sRGB values (without alpha)
        /// </summary>
        /// <param name="foreground">Foreground color of the fill</param>
        /// <param name="background">Background color of the fill</param>
        public Fill(string foreground, string background)
        {
            BackgroundColor = Color.CreateRgb(background);
            ForegroundColor = Color.CreateRgb(foreground);
            PatternFill = PatternValue.Solid;
        }

        /// <summary>
        /// Constructor with color value as sRGB  and fill type
        /// </summary>
        /// <param name="value">Color value</param>
        /// <param name="fillType">Fill type (fill or pattern)</param>
        public Fill(string value, FillType fillType)
        {
            if (fillType == FillType.FillColor)
            {
                backgroundColor = DefaultColor;
                ForegroundColor = Color.CreateRgb(value);
            }
            else
            {
                BackgroundColor = Color.CreateRgb(value);
                foregroundColor = DefaultColor;
            }
            PatternFill = PatternValue.Solid;
        }
        #endregion

        #region methods

        /// <summary>
        /// Override toString method
        /// </summary>
        /// <returns>String of a class</returns>
        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("\"Fill\": {\n");
            AddPropertyAsJson(sb, "BackgroundColor", BackgroundColor);
            AddPropertyAsJson(sb, "ForegroundColor", ForegroundColor);
            AddPropertyAsJson(sb, "PatternFill", PatternFill);
            AddPropertyAsJson(sb, "HashCode", this.GetHashCode(), true);
            sb.Append("\n}");
            return sb.ToString();
        }

        /// <summary>
        /// Method to copy the current object to a new one without casting
        /// </summary>
        /// <returns>Copy of the current object without the internal ID</returns>
        public override AbstractStyle Copy()
        {
            Fill copy = new Fill
            {
                BackgroundColor = BackgroundColor,
                ForegroundColor = ForegroundColor,
                PatternFill = PatternFill
            };
            return copy;
        }

        /// <summary>
        /// Returns a hash code for this instance.
        /// </summary>
        /// <returns>
        /// A hash code for this instance, suitable to be used in hashing algorithms and data structures like a hash table. 
        /// </returns>
        public override int GetHashCode()
        {
            unchecked
            {
                int hashCode = -1564173520;
                hashCode = hashCode * -1521134295 + EqualityComparer<Color>.Default.GetHashCode(BackgroundColor);
                hashCode = hashCode * -1521134295 + EqualityComparer<Color>.Default.GetHashCode(ForegroundColor);
                hashCode = hashCode * -1521134295 + PatternFill.GetHashCode();
                return hashCode;
            }
        }

        /// <summary>
        /// Returns whether two instances are the same
        /// </summary>
        /// <param name="obj">Object to compare</param>
        /// <returns>True if this instance and the other are the same</returns>
        public override bool Equals(object obj)
        {
            return obj is Fill fill &&
                   BackgroundColor == fill.BackgroundColor &&
                   ForegroundColor == fill.ForegroundColor &&
                   PatternFill == fill.PatternFill;
        }

        /// <summary>
        /// Method to copy the current object to a new one with casting
        /// </summary>
        /// <returns>Copy of the current object without the internal ID</returns>
        public Fill CopyFill()
        {
            return (Fill)Copy();
        }

        /// <summary>
        /// Sets the color depending on fill type, using a sRGB value (without alpha)
        /// </summary>
        /// <param name="value">color value</param>
        /// <param name="fillType">fill type (fill or pattern)</param>
        public void SetColor(string value, FillType fillType)
        {
            if (fillType == FillType.FillColor)
            {
                backgroundColor = DefaultColor;
                ForegroundColor = Color.CreateRgb(value);
            }
            else
            {
                BackgroundColor = Color.CreateRgb(value);
                foregroundColor = DefaultColor;
            }
            PatternFill = PatternValue.Solid;
        }

        /// <summary>
        /// Sets the color depending on fill type, using a color object of the type <see cref="Color"/>
        /// </summary>
        /// <param name="value">color value (compound object)</param>
        /// <param name="fillType">fill type (fill or pattern)</param>
        public void SetColor(Color value, FillType fillType)
        {
            if (fillType == FillType.FillColor)
            {
                backgroundColor = DefaultColor;
                ForegroundColor = value;
            }
            else
            {
                BackgroundColor = value;
                foregroundColor = DefaultColor;
            }
            PatternFill = PatternValue.Solid;
        }

        /// <summary>
        /// Sets the color depending on fill type, using a color object, deriving from <see cref="IColor"/>
        /// </summary>
        /// <param name="value">color value (component)</param>
        /// <param name="fillType">fill type (fill or pattern)</param>
        public void SetColor(IColor value, FillType fillType)
        {
            SetColor(GetColorByComponent(value), fillType);
        }
        #endregion

        #region staticMethods
        /// <summary>
        /// Implicit operator to create a Fill object from a string (RGB or ARGB) as foreground color with <see cref="FillType.FillColor"/> 
        /// </summary>
        /// <param name="value">RGB/ARGB value</param>
        public static implicit operator Fill(string value)
        {
            return new Fill(value, FillType.FillColor);
        }

        /// <summary>
        /// Implicit operator to create a Fill object from an indexed color value (<see cref="IndexedColor.Value"/>) as foreground color with <see cref="FillType.FillColor"/> 
        /// </summary>
        /// <param name="index">Color index (0 to 65)</param>
        public static implicit operator Fill(IndexedColor.Value index)
        {
            Fill fill = new Fill();
            fill.PatternFill = PatternValue.Solid;
            fill.ForegroundColor = Color.CreateIndexed(index);
            return fill;
        }

        /// <summary>
        /// Implicit operator to create a Fill object from an indexed color index (numeric) as foreground color with <see cref="FillType.FillColor"/> 
        /// </summary>
        /// <param name="index">Color index (0 to 65)</param>
        public static implicit operator Fill(int index)
        {
            Fill fill = new Fill();
            fill.PatternFill = PatternValue.Solid;
            fill.ForegroundColor = Color.CreateIndexed(index);
            return fill;
        }


        /// <summary>
        /// Gets the pattern name from the enum
        /// </summary>
        /// <param name="pattern">Enum to process</param>
        /// <returns>The valid value of the pattern as String</returns>
        internal static string GetPatternName(PatternValue pattern)
        {
            string output;
            switch (pattern)
            {
                case PatternValue.Solid:
                    output = "solid";
                    break;
                case PatternValue.DarkGray:
                    output = "darkGray";
                    break;
                case PatternValue.MediumGray:
                    output = "mediumGray";
                    break;
                case PatternValue.LightGray:
                    output = "lightGray";
                    break;
                case PatternValue.Gray0625:
                    output = "gray0625";
                    break;
                case PatternValue.Gray125:
                    output = "gray125";
                    break;
                default:
                    output = "none";
                    break;
            }
            return output;
        }

        /// <summary>
        /// Converts a string to its corresponding PatternValue enum
        /// </summary>
        internal static PatternValue GetPatternEnum(string name)
        {
            switch (name)
            {
                case "none": return PatternValue.None;
                case "solid": return PatternValue.Solid;
                case "darkGray": return PatternValue.DarkGray;
                case "mediumGray": return PatternValue.MediumGray;
                case "lightGray": return PatternValue.LightGray;
                case "gray0625": return PatternValue.Gray0625;
                case "gray125": return PatternValue.Gray125;
                default:
                    return PatternValue.None;
            }
        }

        /// <summary>
        /// Gets a Color object based on the passed component
        /// </summary>
        /// <param name="component">Color component</param>
        /// <returns>Color instance</returns>
        private static Color GetColorByComponent(IColor component)
        {
            if (component is SrgbColor)
            {
                return Color.CreateRgb((SrgbColor)component);
            }
            else if (component is IndexedColor)
            {
                return Color.CreateIndexed((IndexedColor)component);
            }
            else if (component is ThemeColor)
            {
                return Color.CreateTheme((ThemeColor)component);
            }
            else if (component is SystemColor)
            {
                return Color.CreateSystem((SystemColor)component);
            }
            else // AutoColor
            {
                return Color.CreateAuto();
            }
        }
        #endregion

    }
}
