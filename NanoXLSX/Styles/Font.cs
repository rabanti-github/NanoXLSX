/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2021
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Exceptions;
using System.Collections.Generic;
using System.Text;

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

        /// <summary>
        /// Enum for the style of the underline property of a stylized text
        /// </summary>
        public enum UnderlineValue
        {
            /// <summary>Text contains a single underline</summary>
            u_single,
            /// <summary>Text contains a double underline</summary>
            u_double,
            /// <summary>Text contains a single, accounting underline</summary>
            singleAccounting,
            /// <summary>Text contains a double, accounting underline</summary>
            doubleAccounting,
            /// <summary>Text contains no underline (default)</summary>
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
        /// Gets or sets whether the font is italic. If true, the font is declared as italic
        /// </summary>
        [Append]
        public bool Italic { get; set; }
        /// <summary>
        /// Gets or sets whether the font is struck through. If true, the font is declared as strike-through
        /// </summary>
        [Append]
        public bool Strike { get; set; }
        /// <summary>
        /// Gets or sets the underline style of the font. If set to <a cref="UnderlineValue.none">none</a> no underline will be applied (default)
        /// </summary>
        [Append]
        public UnderlineValue Underline { get; set; } = UnderlineValue.none;

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
            StringBuilder sb = new StringBuilder();
            sb.Append("\"Font\": {\n");
            AddPropertyAsJson(sb, "Bold", Bold);
            AddPropertyAsJson(sb, "Charset", Charset);
            AddPropertyAsJson(sb, "ColorTheme", ColorTheme);
            AddPropertyAsJson(sb, "ColorValue", ColorValue);
            AddPropertyAsJson(sb, "VerticalAlign", VerticalAlign);
            AddPropertyAsJson(sb, "Family", Family);
            AddPropertyAsJson(sb, "Italic", Italic);
            AddPropertyAsJson(sb, "Name", Name);
            AddPropertyAsJson(sb, "Scheme", Scheme);
            AddPropertyAsJson(sb, "Size", Size);
            AddPropertyAsJson(sb, "Strike", Strike);
            AddPropertyAsJson(sb, "Underline", Underline);
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
            Font copy = new Font();
            copy.Bold = Bold;
            copy.Charset = Charset;
            copy.ColorTheme = ColorTheme;
            copy.ColorValue = ColorValue;
            copy.VerticalAlign = VerticalAlign;
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
            int hashCode = -924704582;
            hashCode = hashCode * -1521134295 + size.GetHashCode();
            hashCode = hashCode * -1521134295 + Bold.GetHashCode();
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(Charset);
            hashCode = hashCode * -1521134295 + ColorTheme.GetHashCode();
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(ColorValue);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(Family);
            hashCode = hashCode * -1521134295 + Italic.GetHashCode();
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(Name);
            hashCode = hashCode * -1521134295 + Scheme.GetHashCode();
            hashCode = hashCode * -1521134295 + Strike.GetHashCode();
            hashCode = hashCode * -1521134295 + Underline.GetHashCode();
            hashCode = hashCode * -1521134295 + VerticalAlign.GetHashCode();
            return hashCode;
        }

        /// <summary>
        /// Returns whether two instances are the same
        /// </summary>
        /// <param name="obj">Object to compare</param>
        /// <returns>True if this instance and the other are the same</returns>
        public override bool Equals(object obj)
        {
            return obj is Font font &&
                   size == font.size &&
                   Bold == font.Bold &&
                   Italic == font.Italic &&
                   Strike == font.Strike &&
                   Underline == font.Underline &&
                   Charset == font.Charset &&
                   ColorTheme == font.ColorTheme &&
                   ColorValue == font.ColorValue &&
                   Family == font.Family &&
                   Name == font.Name &&
                   Scheme == font.Scheme &&
                   VerticalAlign == font.VerticalAlign;
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
