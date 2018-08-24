/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2018
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
using System.Text;

namespace NanoXLSX.Style
{
    public partial class Style
    {
        /// <summary>
        /// Class representing a Fill (background) entry. The Fill entry is used to define background colors and fill patterns
        /// </summary>
        public class Fill : AbstractStyle
        {
            #region constants
            /// <summary>
            /// Default Color (foreground or background) as constant
            /// </summary>
            public const string DEFAULTCOLOR = "FF000000";
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
                /// <summary>No pattern (default)</summary>
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

            #region properties
            /// <summary>
            /// Gets or sets the background color of the fill. The value is expressed as hex string with the format AARRGGBB. AA (Alpha) is usually FF
            /// </summary>
            public string BackgroundColor { get; set; }
            /// <summary>
            /// Gets or sets the foreground color of the fill. The value is expressed as hex string with the format AARRGGBB. AA (Alpha) is usually FF
            /// </summary>
            public string ForegroundColor { get; set; }
            /// <summary>
            /// Gets or sets the indexed color (Default is 64)
            /// </summary>
            public int IndexedColor { get; set; }
            /// <summary>
            /// Gets or sets the pattern type of the fill (Default is none)
            /// </summary>
            public PatternValue PatternFill { get; set; }
            #endregion

            #region constructors
            /// <summary>
            /// Default constructor
            /// </summary>
            public Fill()
            {
                IndexedColor = 64;
                PatternFill = PatternValue.none;
                ForegroundColor = DEFAULTCOLOR;
                BackgroundColor = DEFAULTCOLOR;
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
                IndexedColor = 64;
                PatternFill = PatternValue.solid;
            }

            /// <summary>
            /// Constructor with color value and fill type
            /// </summary>
            /// <param name="value">Color value</param>
            /// <param name="filltype">Fill type (fill or pattern)</param>
            public Fill(string value, FillType filltype)
            {
                if (filltype == FillType.fillColor)
                {
                    BackgroundColor = value;
                    ForegroundColor = DEFAULTCOLOR;
                }
                else
                {
                    BackgroundColor = DEFAULTCOLOR;
                    ForegroundColor = value;
                }
                IndexedColor = 64;
                PatternFill = PatternValue.solid;
            }
            #endregion

            #region methods
            /// <summary>
            /// Override method to calculate the hash of this component
            /// </summary>
            /// <returns>Calculated hash as string</returns>
            public override string CalculateHash()
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(StyleManager.FILLPREFIX);
                CastValue(IndexedColor, ref sb, ':');
                CastValue(PatternFill, ref sb, ':');
                CastValue(ForegroundColor, ref sb, ':');
                CastValue(BackgroundColor, ref sb, null);
                return sb.ToString();
            }

            /// <summary>
            /// Override toString method
            /// </summary>
            /// <returns>String of a class</returns>
            public override string ToString()
            {
                return Hash;
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
                    ForegroundColor = value;
                    BackgroundColor = DEFAULTCOLOR;
                }
                else
                {
                    ForegroundColor = DEFAULTCOLOR;
                    BackgroundColor = value;
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
            public static string GetPatternName(PatternValue pattern)
            {
                string output;
                switch (pattern)
                {
                    case PatternValue.none:
                        output = "none";
                        break;
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
            #endregion

        }
    }
}