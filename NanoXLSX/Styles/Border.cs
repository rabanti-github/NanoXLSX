/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2020
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Text;

namespace NanoXLSX.Styles
{
    /// <summary>
    /// Class representing a Border entry. The Border entry is used to define frames and cell borders
    /// </summary>
    public class Border : AbstractStyle
    {
        #region enums
        /// <summary>
        /// Enum for the border style
        /// </summary>
        public enum StyleValue
        {
            /// <summary>no border</summary>
            none,
            /// <summary>hair border</summary>
            hair,
            /// <summary>dotted border</summary>
            dotted,
            /// <summary>dashed border with double-dots</summary>
            dashDotDot,
            /// <summary>dash-dotted border</summary>
            dashDot,
            /// <summary>dashed border</summary>
            dashed,
            /// <summary>thin border</summary>
            thin,
            /// <summary>medium-dashed border with double-dots</summary>
            mediumDashDotDot,
            /// <summary>slant dash-dotted border</summary>
            slantDashDot,
            /// <summary>medium dash-dotted border</summary>
            mediumDashDot,
            /// <summary>medium dashed border</summary>
            mediumDashed,
            /// <summary>medium border</summary>
            medium,
            /// <summary>thick border</summary>
            thick,
            /// <summary>double border</summary>
            s_double,
        }
        #endregion

        #region properties
        /// <summary>
        /// Gets or sets the color code of the bottom border. The value is expressed as hex string with the format AARRGGBB. AA (Alpha) is usually FF
        /// </summary>
        public string BottomColor { get; set; }
        /// <summary>
        /// Gets or sets the  style of bottom cell border
        /// </summary>
        public StyleValue BottomStyle { get; set; }
        /// <summary>
        /// Gets or sets the color code of the diagonal lines. The value is expressed as hex string with the format AARRGGBB. AA (Alpha) is usually FF
        /// </summary>
        public string DiagonalColor { get; set; }
        /// <summary>
        /// Gets or sets whether the downwards diagonal line is used. If true, the line is used
        /// </summary>
        public bool DiagonalDown { get; set; }
        /// <summary>
        /// Gets or sets whether the upwards diagonal line is used. If true, the line is used
        /// </summary>
        public bool DiagonalUp { get; set; }
        /// <summary>
        /// Gets or sets the style of the diagonal lines
        /// </summary>
        public StyleValue DiagonalStyle { get; set; }
        /// <summary>
        /// Gets or sets the color code of the left border. The value is expressed as hex string with the format AARRGGBB. AA (Alpha) is usually FF
        /// </summary>
        public string LeftColor { get; set; }
        /// <summary>
        /// Gets or sets the style of left cell border
        /// </summary>
        public StyleValue LeftStyle { get; set; }
        /// <summary>
        /// Gets or sets the color code of the right border. The value is expressed as hex string with the format AARRGGBB. AA (Alpha) is usually FF
        /// </summary>
        public string RightColor { get; set; }
        /// <summary>
        /// Gets or sets the style of right cell border
        /// </summary>
        public StyleValue RightStyle { get; set; }
        /// <summary>
        /// Gets or sets the color code of the top border. The value is expressed as hex string with the format AARRGGBB. AA (Alpha) is usually FF
        /// </summary>
        public string TopColor { get; set; }
        /// <summary>
        /// Gets or sets the style of top cell border
        /// </summary>
        public StyleValue TopStyle { get; set; }
        #endregion

        #region constructors
        /// <summary>
        /// Default constructor
        /// </summary>
        public Border()
        {
            BottomColor = string.Empty;
            TopColor = string.Empty;
            LeftColor = string.Empty;
            RightColor = string.Empty;
            DiagonalColor = string.Empty;
            LeftStyle = StyleValue.none;
            RightStyle = StyleValue.none;
            TopStyle = StyleValue.none;
            BottomStyle = StyleValue.none;
            DiagonalStyle = StyleValue.none;
            DiagonalDown = false;
            DiagonalUp = false;
        }
        #endregion

        #region methods
        /// <summary>
        /// Returns a hash code for this instance.
        /// </summary>
        /// <returns>
        /// A hash code for this instance, suitable for use in hashing algorithms and data structures like a hash table. 
        /// </returns>
        public override int GetHashCode()
        {
            int p = 271;
            int r = 1;
            r *= p + (int)this.BottomStyle;
            r *= p + (int)this.DiagonalStyle;
            r *= p + (int)this.TopStyle;
            r *= p + (int)this.LeftStyle;
            r *= p + (int)this.RightStyle;
            r *= p + this.BottomColor.GetHashCode();
            r *= p + this.DiagonalColor.GetHashCode();
            r *= p + this.TopColor.GetHashCode();
            r *= p + this.LeftColor.GetHashCode();
            r *= p + this.RightColor.GetHashCode();
            r *= p + (this.DiagonalDown ? 0 : 1);
            r *= p + (this.DiagonalUp ? 0 : 1);
            return r;
        }

        /// <summary>
        /// Method to copy the current object to a new one without casting
        /// </summary>
        /// <returns>Copy of the current object without the internal ID</returns>
        public override AbstractStyle Copy()
        {
            Border copy = new Border();
            copy.BottomColor = BottomColor;
            copy.BottomStyle = BottomStyle;
            copy.DiagonalColor = DiagonalColor;
            copy.DiagonalDown = DiagonalDown;
            copy.DiagonalStyle = DiagonalStyle;
            copy.DiagonalUp = DiagonalUp;
            copy.LeftColor = LeftColor;
            copy.LeftStyle = LeftStyle;
            copy.RightColor = RightColor;
            copy.RightStyle = RightStyle;
            copy.TopColor = TopColor;
            copy.TopStyle = TopStyle;
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
            return "Border:" + this.GetHashCode();
        }

        /// <summary>
        /// Method to determine whether the object has no values but the default values (means: is empty and must not be processed)
        /// </summary>
        /// <returns>True if empty, otherwise false</returns>
        public bool IsEmpty()
        {
            bool state = true;
            if (BottomColor != string.Empty) { state = false; }
            if (TopColor != string.Empty) { state = false; }
            if (LeftColor != string.Empty) { state = false; }
            if (RightColor != string.Empty) { state = false; }
            if (DiagonalColor != string.Empty) { state = false; }
            if (LeftStyle != StyleValue.none) { state = false; }
            if (RightStyle != StyleValue.none) { state = false; }
            if (TopStyle != StyleValue.none) { state = false; }
            if (BottomStyle != StyleValue.none) { state = false; }
            if (DiagonalStyle != StyleValue.none) { state = false; }
            if (DiagonalDown != false) { state = false; }
            if (DiagonalUp != false) { state = false; }
            return state;
        }
        #endregion

        #region staticMethods
        /// <summary>
        /// Gets the border style name from the enum
        /// </summary>
        /// <param name="style">Enum to process</param>
        /// <returns>The valid value of the border style as String</returns>
        public static string GetStyleName(StyleValue style)
        {
            string output = "";
            switch (style)
            {
                case StyleValue.none:
                    output = "";
                    break;
                case StyleValue.hair:
                    break;
                case StyleValue.dotted:
                    output = "dotted";
                    break;
                case StyleValue.dashDotDot:
                    output = "dashDotDot";
                    break;
                case StyleValue.dashDot:
                    output = "dashDot";
                    break;
                case StyleValue.dashed:
                    output = "dashed";
                    break;
                case StyleValue.thin:
                    output = "thin";
                    break;
                case StyleValue.mediumDashDotDot:
                    output = "mediumDashDotDot";
                    break;
                case StyleValue.slantDashDot:
                    output = "slantDashDot";
                    break;
                case StyleValue.mediumDashDot:
                    output = "mediumDashDot";
                    break;
                case StyleValue.mediumDashed:
                    output = "mediumDashed";
                    break;
                case StyleValue.medium:
                    output = "medium";
                    break;
                case StyleValue.thick:
                    output = "thick";
                    break;
                case StyleValue.s_double:
                    output = "double";
                    break;
                default:
                    output = "";
                    break;
            }
            return output;
        }
        #endregion

    }
}