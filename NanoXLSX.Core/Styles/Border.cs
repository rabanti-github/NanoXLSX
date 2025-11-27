/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Collections.Generic;
using System.Text;
using NanoXLSX.Utils;

namespace NanoXLSX.Styles
{
    /// <summary>
    /// Class representing a Border entry. The Border entry is used to define frames and cell borders
    /// </summary>
    public class Border : AbstractStyle
    {

        #region constants
        /// <summary>
        /// Default border style as constant
        /// </summary>
        public static readonly StyleValue DefaultBorderStyle = StyleValue.None;

        /// <summary>
        /// Default border color as constant
        /// </summary>
        public static readonly string DefaultBorderColor = "";

        #endregion

        #region privateFields
        private string diagonalColor;
        private string leftColor;
        private string rightColor;
        private string topColor;
        private string bottomColor;
        #endregion


        #region enums
        /// <summary>
        /// Enum for the border style, used by the <see cref="Border"/> class
        /// </summary>
        public enum StyleValue
        {
            /// <summary>no border</summary>
            None,
            /// <summary>hair border</summary>
            Hair,
            /// <summary>dotted border</summary>
            Dotted,
            /// <summary>dashed border with double-dots</summary>
            DashDotDot,
            /// <summary>dash-dotted border</summary>
            DashDot,
            /// <summary>dashed border</summary>
            Dashed,
            /// <summary>thin border</summary>
            Thin,
            /// <summary>medium-dashed border with double-dots</summary>
            MediumDashDotDot,
            /// <summary>slant dash-dotted border</summary>
            SlantDashDot,
            /// <summary>medium dash-dotted border</summary>
            MediumDashDot,
            /// <summary>medium dashed border</summary>
            MediumDashed,
            /// <summary>medium border</summary>
            Medium,
            /// <summary>thick border</summary>
            Thick,
            /// <summary>double border</summary>
#pragma warning disable CA1720 // Suppress: Identifiers should not contain types
            Double,
#pragma warning restore CA1720
        }
        #endregion

        #region properties
        /// <summary>
        /// Gets or sets the color code of the bottom border. The value is expressed as hex string with the format AARRGGBB. AA (Alpha) is usually FF. If set, the value will be cast to upper case
        /// </summary>
        [Append]
        public string BottomColor
        {
            get => bottomColor;
            set
            {
                Validators.ValidateColor(value, true, true);
                if (value != null)
                {
                    bottomColor = ParserUtils.ToUpper(value);
                }
                else
                {
                    bottomColor = value;
                }
            }
        }
        /// <summary>
        /// Gets or sets the  style of bottom cell border
        /// </summary>
        [Append]
        public StyleValue BottomStyle { get; set; }
        /// <summary>
        /// Gets or sets the color code of the diagonal lines. The value is expressed as hex string with the format AARRGGBB. AA (Alpha) is usually FF. If set, the value will be cast to upper case
        /// </summary>
        [Append]
        public string DiagonalColor
        {
            get => diagonalColor;
            set
            {
                Validators.ValidateColor(value, true, true);
                if (value != null)
                {
                    diagonalColor = ParserUtils.ToUpper(value);
                }
                else
                {
                    diagonalColor = value;
                }
            }
        }
        /// <summary>
        /// Gets or sets whether the downwards diagonal line is used. If true, the line is used
        /// </summary>
        [Append]
        public bool DiagonalDown { get; set; }
        /// <summary>
        /// Gets or sets whether the upwards diagonal line is used. If true, the line is used
        /// </summary>
        [Append]
        public bool DiagonalUp { get; set; }
        /// <summary>
        /// Gets or sets the style of the diagonal lines
        /// </summary>
        [Append]
        public StyleValue DiagonalStyle { get; set; }
        /// <summary>
        /// Gets or sets the color code of the left border. The value is expressed as hex string with the format AARRGGBB. AA (Alpha) is usually FF. If set, the value will be cast to upper case
        /// </summary>
        [Append]
        public string LeftColor
        {
            get => leftColor;
            set
            {
                Validators.ValidateColor(value, true, true);
                if (value != null)
                {
                    leftColor = ParserUtils.ToUpper(value);
                }
                else
                {
                    leftColor = value;
                }
            }
        }
        /// <summary>
        /// Gets or sets the style of left cell border
        /// </summary>
        [Append]
        public StyleValue LeftStyle { get; set; }
        /// <summary>
        /// Gets or sets the color code of the right border. The value is expressed as hex string with the format AARRGGBB. AA (Alpha) is usually FF. If set, the value will be cast to upper case
        /// </summary>
        [Append]
        public string RightColor
        {
            get => rightColor;
            set
            {
                Validators.ValidateColor(value, true, true);
                if (value != null)
                {
                    rightColor = ParserUtils.ToUpper(value);
                }
                else
                {
                    rightColor = value;
                }
            }
        }
        /// <summary>
        /// Gets or sets the style of right cell border
        /// </summary>
        [Append]
        public StyleValue RightStyle { get; set; }
        /// <summary>
        /// Gets or sets the color code of the top border. The value is expressed as hex string with the format AARRGGBB. AA (Alpha) is usually FF. If set, the value will be cast to upper case
        /// </summary>
        [Append]
        public string TopColor
        {
            get => topColor; set
            {
                Validators.ValidateColor(value, true, true);
                if (value != null)
                {
                    topColor = ParserUtils.ToUpper(value);
                }
                else
                {
                    topColor = value;
                }
            }
        }
        /// <summary>
        /// Gets or sets the style of top cell border
        /// </summary>
        [Append]
        public StyleValue TopStyle { get; set; }
        #endregion

        #region constructors
        /// <summary>
        /// Default constructor
        /// </summary>
        public Border()
        {
            BottomColor = DefaultBorderColor;
            TopColor = DefaultBorderColor;
            LeftColor = DefaultBorderColor;
            RightColor = DefaultBorderColor;
            DiagonalColor = DefaultBorderColor;
            LeftStyle = DefaultBorderStyle;
            RightStyle = DefaultBorderStyle;
            TopStyle = DefaultBorderStyle;
            BottomStyle = DefaultBorderStyle;
            DiagonalStyle = DefaultBorderStyle;
            DiagonalDown = false;
            DiagonalUp = false;
        }
        #endregion

        #region methods
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
                int hashCode = -153001865;
                hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(BottomColor);
                hashCode = hashCode * -1521134295 + BottomStyle.GetHashCode();
                hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(DiagonalColor);
                hashCode = hashCode * -1521134295 + DiagonalDown.GetHashCode();
                hashCode = hashCode * -1521134295 + DiagonalUp.GetHashCode();
                hashCode = hashCode * -1521134295 + DiagonalStyle.GetHashCode();
                hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(LeftColor);
                hashCode = hashCode * -1521134295 + LeftStyle.GetHashCode();
                hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(RightColor);
                hashCode = hashCode * -1521134295 + RightStyle.GetHashCode();
                hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(TopColor);
                hashCode = hashCode * -1521134295 + TopStyle.GetHashCode();
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
            return obj is Border border &&
                   BottomColor == border.BottomColor &&
                   BottomStyle == border.BottomStyle &&
                   DiagonalColor == border.DiagonalColor &&
                   DiagonalDown == border.DiagonalDown &&
                   DiagonalUp == border.DiagonalUp &&
                   DiagonalStyle == border.DiagonalStyle &&
                   LeftColor == border.LeftColor &&
                   LeftStyle == border.LeftStyle &&
                   RightColor == border.RightColor &&
                   RightStyle == border.RightStyle &&
                   TopColor == border.TopColor &&
                   TopStyle == border.TopStyle;
        }

        /// <summary>
        /// Method to copy the current object to a new one without casting
        /// </summary>
        /// <returns>Copy of the current object without the internal ID</returns>
        public override AbstractStyle Copy()
        {
            Border copy = new Border
            {
                BottomColor = BottomColor,
                BottomStyle = BottomStyle,
                DiagonalColor = DiagonalColor,
                DiagonalDown = DiagonalDown,
                DiagonalStyle = DiagonalStyle,
                DiagonalUp = DiagonalUp,
                LeftColor = LeftColor,
                LeftStyle = LeftStyle,
                RightColor = RightColor,
                RightStyle = RightStyle,
                TopColor = TopColor,
                TopStyle = TopStyle
            };
            return copy;
        }

        /// <summary>
        /// Method to copy the current object to a new one with casting
        /// </summary>
        /// <returns>Copy of the current object without the internal ID</returns>
        public Border CopyBorder()
        {
            return (Border)Copy();
        }

        /// <summary>
        /// Override toString method
        /// </summary>
        /// <returns>String of a class</returns>
        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("\"Border\": {\n");
            AddPropertyAsJson(sb, "BottomStyle", BottomStyle);
            AddPropertyAsJson(sb, "DiagonalColor", DiagonalColor);
            AddPropertyAsJson(sb, "DiagonalDown", DiagonalDown);
            AddPropertyAsJson(sb, "DiagonalStyle", DiagonalStyle);
            AddPropertyAsJson(sb, "DiagonalUp", DiagonalUp);
            AddPropertyAsJson(sb, "LeftColor", LeftColor);
            AddPropertyAsJson(sb, "LeftStyle", LeftStyle);
            AddPropertyAsJson(sb, "RightColor", RightColor);
            AddPropertyAsJson(sb, "RightStyle", RightStyle);
            AddPropertyAsJson(sb, "TopColor", TopColor);
            AddPropertyAsJson(sb, "TopStyle", TopStyle);
            AddPropertyAsJson(sb, "HashCode", this.GetHashCode(), true);
            sb.Append("\n}");
            return sb.ToString();
        }

        /// <summary>
        /// Method to determine whether the object has no values but the default values (means: is empty and must not be processed)
        /// </summary>
        /// <returns>True if empty, otherwise false</returns>
        internal bool IsEmpty()
        {
            bool state = true;
            if (BottomColor != DefaultBorderColor)
            { state = false; }
            if (TopColor != DefaultBorderColor)
            { state = false; }
            if (LeftColor != DefaultBorderColor)
            { state = false; }
            if (RightColor != DefaultBorderColor)
            { state = false; }
            if (DiagonalColor != DefaultBorderColor)
            { state = false; }
            if (LeftStyle != DefaultBorderStyle)
            { state = false; }
            if (RightStyle != DefaultBorderStyle)
            { state = false; }
            if (TopStyle != DefaultBorderStyle)
            { state = false; }
            if (BottomStyle != DefaultBorderStyle)
            { state = false; }
            if (DiagonalStyle != DefaultBorderStyle)
            { state = false; }
            if (DiagonalDown)
            { state = false; }
            if (DiagonalUp)
            { state = false; }
            return state;
        }
        #endregion

        #region staticMethods
        /// <summary>
        /// Gets the border style name from the enum
        /// </summary>
        /// <param name="style">Enum to process</param>
        /// <returns>The valid value of the border style as String</returns>
        internal static string GetStyleName(StyleValue style)
        {
            string output = "";
            switch (style)
            {
                case StyleValue.Hair: output = "hair"; break;
                case StyleValue.Dotted: output = "dotted"; break;
                case StyleValue.DashDotDot: output = "dashDotDot"; break;
                case StyleValue.DashDot: output = "dashDot"; break;
                case StyleValue.Dashed: output = "dashed"; break;
                case StyleValue.Thin: output = "thin"; break;
                case StyleValue.MediumDashDotDot: output = "mediumDashDotDot"; break;
                case StyleValue.SlantDashDot: output = "slantDashDot"; break;
                case StyleValue.MediumDashDot: output = "mediumDashDot"; break;
                case StyleValue.MediumDashed: output = "mediumDashed"; break;
                case StyleValue.Medium: output = "medium"; break;
                case StyleValue.Thick: output = "thick"; break;
                case StyleValue.Double: output = "double"; break; // Default / none is already handled (ignored)
            }
            return output;
        }

        /// <summary>
        /// Parses the border style name and returns the matching enum.
        /// </summary>
        /// <param name="styleName">String to parse</param>
        /// <returns>The corresponding StyleValue enum</returns>
        internal static StyleValue GetStyleEnum(string styleName)
        {
            switch (styleName)
            {
                case "hair": return StyleValue.Hair;
                case "dotted": return StyleValue.Dotted;
                case "dashDotDot": return StyleValue.DashDotDot;
                case "dashDot": return StyleValue.DashDot;
                case "dashed": return StyleValue.Dashed;
                case "thin": return StyleValue.Thin;
                case "mediumDashDotDot": return StyleValue.MediumDashDotDot;
                case "slantDashDot": return StyleValue.SlantDashDot;
                case "mediumDashDot": return StyleValue.MediumDashDot;
                case "mediumDashed": return StyleValue.MediumDashed;
                case "medium": return StyleValue.Medium;
                case "thick": return StyleValue.Thick;
                case "double": return StyleValue.Double;

                default:
                    return StyleValue.None; // fallback – change to throw if you'd prefer strict handling
            }
        }
        #endregion

    }
}
