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
    /// Class representing a Font entry. The Font entry is used to define text formatting
    /// </summary>
    public class Font : AbstractStyle
    {
        #region constants
        /// <summary>
        /// Default font family as constant
        /// </summary>
        public const string DEFAULTFONT = "Calibri";
        #endregion

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
        private int size;
        #endregion

        #region properties
        /// <summary>
        /// Gets or sets whether the font is bold. If true, the font is declared as bold
        /// </summary>
        public bool Bold { get; set; }
        /// <summary>
        /// Gets or sets the char set of the Font (Default is empty)
        /// </summary>
        public string Charset { get; set; }
        /// <summary>
        /// Gets or sets the font color theme (Default is 1)
        /// </summary>
        public int ColorTheme { get; set; }
        /// <summary>
        /// Gets or sets the font color (default is empty)
        /// </summary>
        public string ColorValue { get; set; }
        /// <summary>
        /// Gets or sets whether the font has a double underline. If true, the font is declared with a double underline
        /// </summary>
        public bool DoubleUnderline { get; set; }
        /// <summary>
        ///  Gets or sets the font family (Default is 2)
        /// </summary>
        public string Family { get; set; }
        /// <summary>
        /// Gets whether the font is equals the default font
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
        public bool Italic { get; set; }
        /// <summary>
        /// Gets or sets the font name (Default is Calibri)
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// Gets or sets the font scheme (Default is minor)
        /// </summary>
        public SchemeValue Scheme { get; set; }
        /// <summary>
        /// Gets or sets the font size. Valid range is from 8 to 75
        /// </summary>
        public int Size
        {
            get { return size; }
            set
            {
                if (value < 8) { size = 8; }
                else if (value > 75) { size = 72; }
                else { size = value; }
            }
        }
        /// <summary>
        /// Gets or sets whether the font is struck through. If true, the font is declared as strike-through
        /// </summary>
        public bool Strike { get; set; }
        /// <summary>
        /// Gets or sets whether the font is underlined. If true, the font is declared as underlined
        /// </summary>
        public bool Underline { get; set; }
        /// <summary>
        /// Gets or sets the alignment of the font (Default is none)
        /// </summary>
        public VerticalAlignValue VerticalAlign { get; set; }
        #endregion

        #region constructors
        /// <summary>
        /// Default constructor
        /// </summary>
        public Font()
        {
            size = 11;
            Name = DEFAULTFONT;
            Family = "2";
            ColorTheme = 1;
            ColorValue = string.Empty;
            Charset = string.Empty;
            Scheme = SchemeValue.minor;
            VerticalAlign = VerticalAlignValue.none;
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
            int p = 257;
            int r = 1;
            r *= p + (this.Bold ? 0 : 1);
            r *= p + (this.Italic ? 0 : 1);
            r *= p + (this.Underline ? 0 : 1);
            r *= p + (this.DoubleUnderline ? 0 : 1);
            r *= p + (this.Strike ? 0 : 1);
            r *= p + this.ColorTheme;
            r *= p + this.ColorValue.GetHashCode();
            r *= p + this.Family.GetHashCode();
            r *= p + this.Name.GetHashCode();
            r *= p + this.Scheme.GetHashCode();
            r *= p + (int)this.VerticalAlign;
            r *= p + this.Charset.GetHashCode();
            r *= p + this.size;
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
