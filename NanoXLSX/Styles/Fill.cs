﻿/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Exceptions;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

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
        public static readonly string DEFAULT_COLOR = "FF000000";
        /// <summary>
        /// Default index color
        /// </summary>
        public static readonly int DEFAULT_INDEXED_COLOR = 64;
        /// <summary>
        /// Default pattern
        /// </summary>
        public static readonly PatternValue DEFAULT_PATTERN_FILL = PatternValue.none;

        #endregion

        #region enums
        /// <summary>
        /// Enum for the type of the color
        /// </summary>
        public enum FillType
        {
            /// <summary>Color defines a pattern color </summary>
            patternColor,
            /// <summary>Color defines a solid fill color </summary>
            fillColor,
        }
        /// <summary>
        /// Enum for the pattern values
        /// </summary>
        public enum PatternValue
        {
            /// <summary>
            /// No pattern (default)
            /// </summary>
            /// <remarks>The value none will lead to a invalidation of the foreground or background color values</remarks>
            none,
            /// <summary>Solid fill (for colors)</summary>
            solid,
            /// <summary>Dark gray fill</summary>
            darkGray,
            /// <summary>Medium gray fill</summary>
            mediumGray,
            /// <summary>Light gray fill</summary>
            lightGray,
            /// <summary>6.25% gray fill</summary>
            gray0625,
            /// <summary>12.5% gray fill</summary>
            gray125,
        }
        #endregion

        #region privateFields
        private string backgroundColor = DEFAULT_COLOR;
        private string foregroundColor = DEFAULT_COLOR;
        #endregion

        #region properties
        /// <summary>
        /// Gets or sets the background color of the fill. The value is expressed as hex string with the format AARRGGBB. AA (Alpha) is usually FF
        /// </summary>
        /// <remarks>If a background color is set and the <see cref="PatternFill">PatternFill</see> Property is currently set to <see cref="PatternValue.none">PatternValue.none</see>, 
        /// the PatternFill property will be automatically set to <see cref="PatternValue.solid">PatternValue.solid</see>, since none invalidates the color values of the foreground or background</remarks>
        [Append]
        public string BackgroundColor
        {
            get => backgroundColor;
            set
            {
                ValidateColor(value, true);
                backgroundColor = value;
                if (PatternFill == PatternValue.none)
                {
                    PatternFill = PatternValue.solid;
                }
            }
        }
        /// <summary>
        /// Gets or sets the foreground color of the fill. The value is expressed as hex string with the format AARRGGBB. AA (Alpha) is usually FF
        /// </summary>
        /// <remarks>If a foreground color is set and the <see cref="PatternFill">PatternFill</see> Property is currently set to <see cref="PatternValue.none">PatternValue.none</see>, 
        /// the PatternFill property will be automatically set to <see cref="PatternValue.solid">PatternValue.solid</see>, since none invalidates the color values of the foreground or background</remarks>
        [Append]
        public string ForegroundColor
        {
            get => foregroundColor;
            set
            {
                ValidateColor(value, true);
                foregroundColor = value;
                if (PatternFill == PatternValue.none)
                {
                    PatternFill = PatternValue.solid;
                }
            }
        }
        /// <summary>
        /// Gets or sets the indexed color (Default is 64)
        /// </summary>
        [Append]
        public int IndexedColor { get; set; }
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
            IndexedColor = DEFAULT_INDEXED_COLOR;
            PatternFill = DEFAULT_PATTERN_FILL;
            foregroundColor = DEFAULT_COLOR;
            backgroundColor = DEFAULT_COLOR;
        }
        /// <summary>
        /// Constructor with foreground and background color
        /// </summary>
        /// <param name="foreground">Foreground color of the fill</param>
        /// <param name="background">Background color of the fill</param>
        public Fill(string foreground, string background)
        {
            BackgroundColor = background;
            ForegroundColor = foreground;
            IndexedColor = DEFAULT_INDEXED_COLOR;
            PatternFill = PatternValue.solid;
        }

        /// <summary>
        /// Constructor with color value and fill type
        /// </summary>
        /// <param name="value">Color value</param>
        /// <param name="fillType">Fill type (fill or pattern)</param>
        public Fill(string value, FillType fillType)
        {
            if (fillType == FillType.fillColor)
            {
                backgroundColor = DEFAULT_COLOR;
                ForegroundColor = value;
            }
            else
            {
                BackgroundColor = value;
                foregroundColor = DEFAULT_COLOR;
            }
            IndexedColor = DEFAULT_INDEXED_COLOR;
            PatternFill = PatternValue.solid;
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
            AddPropertyAsJson(sb, "IndexedColor", IndexedColor);
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
            Fill copy = new Fill();
            copy.BackgroundColor = BackgroundColor;
            copy.ForegroundColor = ForegroundColor;
            copy.IndexedColor = IndexedColor;
            copy.PatternFill = PatternFill;
            return copy;
        }

        /// <summary>
        /// Returns a hash code for this instance.
        /// </summary>
        /// <returns>
        /// A hash code for this instance, suitable for use in hashing algorithms and data structures like a hash table. 
        /// </returns>
        public override int GetHashCode()
        {
            int hashCode = -1564173520;
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(BackgroundColor);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(ForegroundColor);
            hashCode = hashCode * -1521134295 + IndexedColor.GetHashCode();
            hashCode = hashCode * -1521134295 + PatternFill.GetHashCode();
            return hashCode;
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
                   IndexedColor == fill.IndexedColor &&
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
        /// Sets the color and the depending fill type
        /// </summary>
        /// <param name="value">color value</param>
        /// <param name="filltype">fill type (fill or pattern)</param>
        public void SetColor(string value, FillType filltype)
        {
            if (filltype == FillType.fillColor)
            {
                backgroundColor = DEFAULT_COLOR;
                ForegroundColor = value;
            }
            else
            {
                BackgroundColor = value;
                foregroundColor = DEFAULT_COLOR;
            }
            PatternFill = PatternValue.solid;
        }
        #endregion

        #region staticMethods
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
                case PatternValue.solid:
                    output = "solid";
                    break;
                case PatternValue.darkGray:
                    output = "darkGray";
                    break;
                case PatternValue.mediumGray:
                    output = "mediumGray";
                    break;
                case PatternValue.lightGray:
                    output = "lightGray";
                    break;
                case PatternValue.gray0625:
                    output = "gray0625";
                    break;
                case PatternValue.gray125:
                    output = "gray125";
                    break;
                default:
                    output = "none";
                    break;
            }
            return output;
        }

        /// <summary>
        /// Validates the passed string, whether it is a valid RGB value that can be used for Fills or Fonts
        /// </summary>
        /// <exception cref="StyleException">A StyleException is thrown if an invalid hex value is passed</exception>
        /// <param name="hexCode">Hex string to check</param>
        /// <param name="useAlpha">If true, two additional characters (total 8) are expected as alpha value</param>
        /// <param name="allowEmpty">Optional parameter that allows null or empty as valid values</param>
        public static void ValidateColor(string hexCode, bool useAlpha, bool allowEmpty = false)
        {
            if (string.IsNullOrEmpty(hexCode))
            {
                if (allowEmpty)
                {
                    return;
                }
                throw new StyleException("The color expression was null or empty");
            }

            int length = useAlpha ? 8 : 6;
            if (hexCode.Length != length)
            {
                throw new StyleException("The value '" + hexCode + "' is invalid. A valid value must contain " + length + " hex characters");
            }
            if (!Regex.IsMatch(hexCode, "[a-fA-F0-9]{6,8}"))
            {
                throw new StyleException("The expression '" + hexCode + "' is not a valid hex value");
            }
        }
        #endregion

    }
}
