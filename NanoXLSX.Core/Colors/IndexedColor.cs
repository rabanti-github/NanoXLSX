/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Exceptions;
using NanoXLSX.Interfaces;
using NanoXLSX.Utils;
using static NanoXLSX.Colors.IndexedColor;

namespace NanoXLSX.Colors
{
    /// <summary>
    /// Class representing an indexed color from the legacy OOXML / Excel indexed color palette.
    /// </summary>
    public class IndexedColor : ITypedColor<Value>
    {
        /// <summary>
        /// Legacy OOXML / Excel indexed color palette.
        /// <para>
        /// This palette exists for backward compatibility with older Excel formats.
        /// Indices 0–7 are redundant with 8–15.
        /// </para>
        /// </summary>
        public enum Value : byte
        {
            /// <summary>Black (duplicate of index 8).</summary>
            Black0 = 0,
            /// <summary>White (duplicate of index 9).</summary>
            White1 = 1,
            /// <summary>Red (duplicate of index 10).</summary>
            Red2 = 2,
            /// <summary>Bright green (duplicate of index 11).</summary>
            BrightGreen3 = 3,
            /// <summary>Blue (duplicate of index 12).</summary>
            Blue4 = 4,
            /// <summary>Yellow (duplicate of index 13).</summary>
            Yellow5 = 5,
            /// <summary>Magenta (duplicate of index 14).</summary>
            Magenta6 = 6,
            /// <summary>Cyan (duplicate of index 15).</summary>
            Cyan7 = 7,
            /// <summary>Black (#000000).</summary>
            Black = 8,
            /// <summary>White (#FFFFFF).</summary>
            White = 9,
            /// <summary>Red (#FF0000).</summary>
            Red = 10,
            /// <summary>Bright green (#00FF00).</summary>
            BrightGreen = 11,
            /// <summary>Blue (#0000FF).</summary>
            Blue = 12,
            /// <summary>Yellow (#FFFF00).</summary>
            Yellow = 13,
            /// <summary>Magenta / Fuchsia (#FF00FF).</summary>
            Magenta = 14,
            /// <summary>Cyan / Aqua (#00FFFF).</summary>
            Cyan = 15,
            /// <summary>Dark red / maroon (#800000).</summary>
            DarkRed = 16,
            /// <summary>Dark green (#008000).</summary>
            DarkGreen = 17,
            /// <summary>Dark blue / navy (#000080).</summary>
            DarkBlue = 18,
            /// <summary>Olive (#808000).</summary>
            Olive = 19,
            /// <summary>Purple (#800080).</summary>
            Purple = 20,
            /// <summary>Teal (#008080).</summary>
            Teal = 21,
            /// <summary>Light gray / silver (#C0C0C0).</summary>
            LightGray = 22,
            /// <summary>Medium gray (#808080).</summary>
            Gray = 23,
            /// <summary>Light cornflower blue (#9999FF).</summary>
            LightCornflowerBlue = 24,
            /// <summary>Dark rose (#993366).</summary>
            DarkRose = 25,
            /// <summary>Light yellow (#FFFFCC).</summary>
            LightYellow = 26,
            /// <summary>Light cyan (#CCFFFF).</summary>
            LightCyan = 27,
            /// <summary>Dark purple (#660066).</summary>
            DarkPurple = 28,
            /// <summary>Salmon pink (#FF8080).</summary>
            Salmon = 29,
            /// <summary>Medium blue (#0066CC).</summary>
            MediumBlue = 30,
            /// <summary>Light lavender blue (#CCCCFF).</summary>
            LightLavender = 31,
            /// <summary>Dark navy blue (#000080).</summary>
            Navy = 32,
            /// <summary>Strong magenta (#FF00FF).</summary>
            StrongMagenta = 33,
            /// <summary>Strong yellow (#FFFF00).</summary>
            StrongYellow = 34,
            /// <summary>Strong cyan (#00FFFF).</summary>
            StrongCyan = 35,
            /// <summary>Dark violet (#800080).</summary>
            DarkViolet = 36,
            /// <summary>Dark maroon (#800000).</summary>
            DarkMaroon = 37,
            /// <summary>Dark teal (#008080).</summary>
            DarkTeal = 38,
            /// <summary>Pure blue (#0000FF).</summary>
            PureBlue = 39,
            /// <summary>Sky blue (#00CCFF).</summary>
            SkyBlue = 40,
            /// <summary>Pale cyan (#CCFFFF).</summary>
            PaleCyan = 41,
            /// <summary>Light mint green (#CCFFCC).</summary>
            LightMint = 42,
            /// <summary>Light pastel yellow (#FFFF99).</summary>
            PastelYellow = 43,
            /// <summary>Light sky blue (#99CCFF).</summary>
            LightSkyBlue = 44,
            /// <summary>Rose pink (#FF99CC).</summary>
            Rose = 45,
            /// <summary>Lavender (#CC99FF).</summary>
            Lavender = 46,
            /// <summary>Peach (#FFCC99).</summary>
            Peach = 47,
            /// <summary>Royal blue (#3366FF).</summary>
            RoyalBlue = 48,
            /// <summary>Turquoise (#33CCCC).</summary>
            Turquoise = 49,
            /// <summary>Light olive green (#99CC00).</summary>
            LightOlive = 50,
            /// <summary>Gold (#FFCC00).</summary>
            Gold = 51,
            /// <summary>Orange (#FF9900).</summary>
            Orange = 52,
            /// <summary>Dark orange (#FF6600).</summary>
            DarkOrange = 53,
            /// <summary>Blue gray (#666699).</summary>
            BlueGray = 54,
            /// <summary>Medium gray (#969696).</summary>
            MediumGray = 55,
            /// <summary>Dark slate blue (#003366).</summary>
            DarkSlateBlue = 56,
            /// <summary>Sea green (#339966).</summary>
            SeaGreen = 57,
            /// <summary>Very dark green (#003300).</summary>
            VeryDarkGreen = 58,
            /// <summary>Dark olive (#333300).</summary>
            DarkOlive = 59,
            /// <summary>Brown (#993300).</summary>
            Brown = 60,
            /// <summary>Dark rose (duplicate of index 25).</summary>
            DarkRoseDuplicate = 61,
            /// <summary>Indigo / dark blue-purple (#333399).</summary>
            Indigo = 62,
            /// <summary>Very dark gray (#333333).</summary>
            VeryDarkGray = 63,
            /// <summary>
            /// System foreground color.
            /// <para>
            /// The actual color is determined by the host operating system or theme.
            /// </para>
            /// </summary>
            SystemForeground = 64,
            /// <summary>
            /// System background color.
            /// <para>
            /// The actual color is determined by the host operating system or theme.
            /// </para>
            /// </summary>
            SystemBackground = 65
        }

        /// <summary>
        /// Default indexed color (system foreground color)
        /// </summary>
        public const Value DefaultIndexedColor = Value.SystemForeground;
        /// <summary>
        /// Default ARGB value for system foreground color
        /// </summary>
        public const string DefaultSystemForegroundColorArgb = "FF000000";
        /// <summary>
        /// Default ARGB value for system background color
        /// </summary>
        public const string DefaultSystemBackgroundColorArgb = "FFFFFFFF";


        /// <summary>
        /// Value of the indexed color
        /// </summary>
        public Value ColorValue { get; set; }

        /// <summary>
        /// String representation of the indexed color value
        /// </summary>
        public string StringValue => ParserUtils.ToString((int)ColorValue);

        /// <summary>
        /// Default constructor with default indexed color
        /// </summary>
        public IndexedColor()
        {
            ColorValue = DefaultIndexedColor;
        }

        /// <summary>
        /// Constructor with specified indexed color value
        /// </summary>
        /// <param name="color">Indexed color</param>
        public IndexedColor(Value color)
        {
            ColorValue = color;
        }

        /// <summary>
        /// Constructor with specified indexed color index
        /// </summary>
        /// <param name="colorIndex">Color index</param>
        /// <exception cref="StyleException">Throws a StyleException if the color index is out of range</exception>
        public IndexedColor(int colorIndex)
        {
            if (colorIndex < 0 || colorIndex > 65)
            {
                throw new StyleException("Indexed color value must be between 0 and 65.");
            }
            ColorValue = (Value)colorIndex;
        }

        /// <summary>
        /// Gets the ARGB hex code representation of the indexed color
        /// </summary>
        /// <returns>ARGB value of the current color instance</returns>
        public string GetArgbValue()
        {
            return GetArgbValue(ColorValue);
        }

        /// <summary>
        /// Gets the sRGB color representation of the indexed color
        /// </summary>
        /// <returns>sRGB color instance</returns>
        public SrgbColor GetSrgbColor()
        {
            return new SrgbColor(GetArgbValue());
        }


        /// <summary>
        /// Determines whether the specified object is equal to the current object
        /// </summary>
        /// <param name="obj">Other object to compare</param>
        /// <returns>True if both objects are equal</returns>
        public override bool Equals(object obj)
        {
            return obj is IndexedColor color &&
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

        /// <summary>
        /// Maps the indexed color value to its ARGB hex code representation
        /// </summary>
        /// <param name="indexedValue">Enum value</param>
        /// <returns>ARGB value</returns>
        public static string GetArgbValue(Value indexedValue)
        {
            switch (indexedValue)
            {
                // 0–7 (duplicates of 8–15)
                case Value.Black0:
                case Value.Black:
                    return "FF000000";

                case Value.White1:
                case Value.White:
                    return "FFFFFFFF";

                case Value.Red2:
                case Value.Red:
                    return "FFFF0000";

                case Value.BrightGreen3:
                case Value.BrightGreen:
                    return "FF00FF00";

                case Value.Blue4:
                case Value.Blue:
                case Value.PureBlue:
                    return "FF0000FF";

                case Value.Yellow5:
                case Value.Yellow:
                case Value.StrongYellow:
                    return "FFFFFF00";

                case Value.Magenta6:
                case Value.Magenta:
                case Value.StrongMagenta:
                    return "FFFF00FF";

                case Value.Cyan7:
                case Value.Cyan:
                case Value.StrongCyan:
                    return "FF00FFFF";

                // Extended palette
                case Value.DarkRed:
                case Value.DarkMaroon:
                    return "FF800000";

                case Value.DarkGreen:
                    return "FF008000";

                case Value.DarkBlue:
                case Value.Navy:
                    return "FF000080";

                case Value.Olive:
                    return "FF808000";

                case Value.Purple:
                case Value.DarkViolet:
                    return "FF800080";

                case Value.Teal:
                case Value.DarkTeal:
                    return "FF008080";

                case Value.LightGray:
                    return "FFC0C0C0";

                case Value.Gray:
                    return "FF808080";

                case Value.LightCornflowerBlue:
                    return "FF9999FF";

                case Value.DarkRose:
                case Value.DarkRoseDuplicate:
                    return "FF993366";

                case Value.LightYellow:
                    return "FFFFFFCC";

                case Value.LightCyan:
                case Value.PaleCyan:
                    return "FFCCFFFF";

                case Value.DarkPurple:
                    return "FF660066";

                case Value.Salmon:
                    return "FFFF8080";

                case Value.MediumBlue:
                    return "FF0066CC";

                case Value.LightLavender:
                    return "FFCCCCFF";

                case Value.SkyBlue:
                    return "FF00CCFF";

                case Value.LightMint:
                    return "FFCCFFCC";

                case Value.PastelYellow:
                    return "FFFFFF99";

                case Value.LightSkyBlue:
                    return "FF99CCFF";

                case Value.Rose:
                    return "FFFF99CC";

                case Value.Lavender:
                    return "FFCC99FF";

                case Value.Peach:
                    return "FFFFCC99";

                case Value.RoyalBlue:
                    return "FF3366FF";

                case Value.Turquoise:
                    return "FF33CCCC";

                case Value.LightOlive:
                    return "FF99CC00";

                case Value.Gold:
                    return "FFFFCC00";

                case Value.Orange:
                    return "FFFF9900";

                case Value.DarkOrange:
                    return "FFFF6600";

                case Value.BlueGray:
                    return "FF666699";

                case Value.MediumGray:
                    return "FF969696";

                case Value.DarkSlateBlue:
                    return "FF003366";

                case Value.SeaGreen:
                    return "FF339966";

                case Value.VeryDarkGreen:
                    return "FF003300";

                case Value.DarkOlive:
                    return "FF333300";

                case Value.Brown:
                    return "FF993300";

                case Value.Indigo:
                    return "FF333399";

                case Value.VeryDarkGray:
                    return "FF333333";

                case Value.SystemBackground:
                    // Excel default: white background
                    return DefaultSystemBackgroundColorArgb;

                default:
                    // Excel default: black text
                    return DefaultSystemForegroundColorArgb;
            }
        }

        /// <summary>
        /// Implicit conversion from Value to IndexedColor
        /// </summary>
        /// <param name="value">Enum value that is automatically assigned to the color</param>
        public static implicit operator IndexedColor(Value value)
        {
            return new IndexedColor(value);
        }
    }
}
