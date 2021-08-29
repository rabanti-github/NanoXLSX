/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2021
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Exceptions;

namespace NanoXLSX.Styles
{
    /// <summary>
    /// Class representing a Font entry. The Font entry is used to define text formatting
    /// </summary>
    public class Font : AbstractStyle
    {
        #region constants
        /// <summary>
        /// Default font family as constant
        /// </summary>
        public static readonly string DEFAULT_FONT_NAME = "Calibri";

        /// <summary>
        /// Maximum possible font size
        /// </summary>
        public static readonly float MIN_FONT_SIZE = 1f;

        /// <summary>
        /// Minimum possible font size
        /// </summary>
        public static readonly float MAX_FONT_SIZE = 409f;

        /// <summary>
        /// Default font size
        /// </summary>
        public static readonly float DEFAULT_FONT_SIZE = 11f;
        #endregion

        /// <summary>
        /// Default font family
        /// </summary>
        public static readonly string DEFAULT_FONT_FAMILY = "2";

        /// <summary>
        /// Default font scheme
        /// </summary>
        public static readonly SchemeValue DEFAULT_FONT_SCHEME = SchemeValue.minor;

        /// <summary>
        /// Default vertical alignment
        /// </summary>
        public static readonly VerticalAlignValue DEFAULT_VERTICAL_ALIGN = VerticalAlignValue.none;

        #region enums
        /// <summary>
        /// Enum for the font scheme
        /// </summary>
        public enum SchemeValue
        {
            /// <summary>Font scheme is major</summary>
            major,
            /// <summary>Font scheme is minor (default)</summary>
            minor,
            /// <summary>No Font scheme is used</summary>
            none,
        }
        /// <summary>
        /// Enum for the vertical alignment of the text from base line
        /// </summary>
        public enum VerticalAlignValue
        {
            // baseline, // Maybe not used in Excel
            /// <summary>Text will be rendered as subscript</summary>
            subscript,
            /// <summary>Text will be rendered as superscript</summary>
            superscript,
            /// <summary>Text will be rendered normal</summary>
            none,
        }
        #endregion

        #region privateFields
        private float size;
        private string name = DEFAULT_FONT_NAME;
        private int colorTheme;
        private string colorValue = "";
        #endregion

        #region properties
        /// <summary>
        /// Gets or sets whether the font is bold. If true, the font is declared as bold
        /// </summary>
        [Append]
        public bool Bold { get; set; }
        /// <summary>
        /// Gets or sets the char set of the Font (Default is empty)
        /// </summary>
        [Append]
        public string Charset { get; set; }
        /// <summary>
        /// Gets or sets the font color theme (Default is 1)
        /// </summary>
        /// <exception cref="StyleException">Test of the ConvertArray methodStyleException if the number is below 1</exception>
        [Append]
        public int ColorTheme
        {
            get => colorTheme;
            set
            {
                if (value < 1)
                {
                    throw new StyleException("The color theme number " + value + " is invalid. Should be >0");
                }
                colorTheme = value;
            }
        }
        /// <summary>
        /// Gets or sets the color code of the font color. The value is expressed as hex string with the format AARRGGBB. AA (Alpha) is usually FF.
        /// To omit the color, an empty string can be set. Empty is also default.
        /// </summary>
        /// <exception cref="StyleException">Test of the ConvertArray methodStyleException if the passed ARGB value is not valid</exception>
        [Append]
        public string ColorValue { 
            get => colorValue;
            set 
            {
                Fill.ValidateColor(value, true, true);
                colorValue = value;
            } 
        }
        /// <summary>
        /// Gets or sets whether the font has a double underline. If true, the font is declared with a double underline
        /// </summary>
        [Append]
        public bool DoubleUnderline { get; set; }
        /// <summary>
        ///  Gets or sets the font family (Default is 2)
        /// </summary>
        [Append]
        public string Family { get; set; }
        /// <summary>
        /// Gets whether the font is equal to the default font
        /// </summary>
        [Append(Ignore = true)]
        public bool IsDefaultFont
        {
            get
            {
                Font temp = new Font();
                return Equals(temp);
            }
        }
        /// <summary>
        /// Gets or sets whether the font is italic. If true, the font is declared as italic
        /// </summary>
        [Append]
        public bool Italic { get; set; }

        /// <summary>
        /// Gets or sets the font name (Default is Calibri)
        /// </summary>
        /// <exception cref="StyleException">A StyleException is thrown if the name is null or empty</exception>
        /// <remarks>Note that the font name is not validated whether it is a valid or existing font</remarks>
        [Append]
        public string Name
        {
            get { return name; }
            set
            {
                if (string.IsNullOrEmpty(value))
                {
                    throw new StyleException("The font name was null or empty");
                }
                name = value;
            }
        }
        /// <summary>
        /// Gets or sets the font scheme (Default is minor)
        /// </summary>
        [Append]
        public SchemeValue Scheme { get; set; }
        /// <summary>
        /// Gets or sets the font size. Valid range is from 1 to 409
        /// </summary>
        [Append]
        public float Size
        {
            get { return size; }
            set
            {
                if (value < MIN_FONT_SIZE)
                { size = MIN_FONT_SIZE; }
                else if (value > MAX_FONT_SIZE)
                { size = MAX_FONT_SIZE; }
                else { size = value; }
            }
        }
        /// <summary>
        /// Gets or sets whether the font is struck through. If true, the font is declared as strike-through
        /// </summary>
        [Append]
        public bool Strike { get; set; }
        /// <summary>
        /// Gets or sets whether the font is underlined. If true, the font is declared as underlined
        /// </summary>
        [Append]
        public bool Underline { get; set; }
        /// <summary>
        /// Gets or sets the alignment of the font (Default is none)
        /// </summary>
        [Append]
        public VerticalAlignValue VerticalAlign { get; set; }
        #endregion

        #region constructors
        /// <summary>
        /// Default constructor
        /// </summary>
        public Font()
        {
            size = DEFAULT_FONT_SIZE;
            Name = DEFAULT_FONT_NAME;
            Family = DEFAULT_FONT_FAMILY;
            ColorTheme = 1;
            ColorValue = string.Empty;
            Charset = string.Empty;
            Scheme = DEFAULT_FONT_SCHEME;
            VerticalAlign = DEFAULT_VERTICAL_ALIGN;
        }
        #endregion

        #region methods            
        /// <summary>
        /// Override toString method
        /// </summary>
        /// <returns>String of a class</returns>
        public override string ToString()
        {
            return "Font:" + this.GetHashCode();
        }

        /// <summary>
        /// Method to copy the current object to a new one without casting
        /// </summary>
        /// <returns>Copy of the current object without the internal ID</returns>
        public override AbstractStyle Copy()
        {
            Font copy = new Font();
            copy.Bold = Bold;
            copy.Charset = Charset;
            copy.ColorTheme = ColorTheme;
            copy.ColorValue = ColorValue;
            copy.VerticalAlign = VerticalAlign;
            copy.DoubleUnderline = DoubleUnderline;
            copy.Family = Family;
            copy.Italic = Italic;
            copy.Name = Name;
            copy.Scheme = Scheme;
            copy.Size = Size;
            copy.Strike = Strike;
            copy.Underline = Underline;
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
            const int p = 257;
            int r = 1;
            r *= p + (this.Bold ? 0 : 1);
            r *= p + (this.Italic ? 0 : 2);
            r *= p + (this.Underline ? 0 : 4);
            r *= p + (this.DoubleUnderline ? 0 : 8);
            r *= p + (this.Strike ? 0 : 16);
            r *= p + this.ColorTheme;
            r *= p + this.ColorValue.GetHashCode();
            r *= p + this.Family.GetHashCode();
            r *= p + this.Name.GetHashCode();
            r *= p + this.Scheme.GetHashCode();
            r *= p + this.VerticalAlign.GetHashCode();
            r *= p + this.Charset.GetHashCode();
            r *= p + this.size.GetHashCode();
            return r;
        }

        /// <summary>
        /// Method to copy the current object to a new one with casting
        /// </summary>
        /// <returns>Copy of the current object without the internal ID</returns>
        public Font CopyFont()
        {
            return (Font)Copy();
        }

        #endregion
    }
}
