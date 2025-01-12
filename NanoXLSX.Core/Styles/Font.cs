/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Collections.Generic;
using System.Text;
using NanoXLSX.Exceptions;
using NanoXLSX.Styles;
using NanoXLSX.Utils;
using static NanoXLSX.Themes.Theme;

namespace NanoXLSX.Styles
{
    /// <summary>
    /// Class representing a Font entry. The Font entry is used to define text formatting
    /// </summary>
    public class Font : AbstractStyle
    {
        #region constants
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

        /// <summary>
        /// The default font name that is declared as Major Font (See <see cref="SchemeValue"/>)
        /// </summary>
        public static readonly string DEFAULT_MAJOR_FONT = "Calibri Light";
        /// <summary>
        /// The default font name that is declared as Minor Font (See <see cref="SchemeValue"/>)
        /// </summary>
        public static readonly string DEFAULT_MINOR_FONT = "Calibri";

        /// <summary>
        /// Default font family as constant
        /// </summary>
        public static readonly string DEFAULT_FONT_NAME = DEFAULT_MINOR_FONT;

        /// <summary>
        /// Default font family
        /// </summary>
        public static readonly FontFamilyValue DEFAULT_FONT_FAMILY = FontFamilyValue.Swiss;

        /// <summary>
        /// Default font scheme
        /// </summary>
        public static readonly SchemeValue DEFAULT_FONT_SCHEME = SchemeValue.minor;

        /// <summary>
        /// Default vertical alignment
        /// </summary>
        public static readonly VerticalTextAlignValue DEFAULT_VERTICAL_ALIGN = VerticalTextAlignValue.none;
        #endregion

        #region enums
        /// <summary>
        /// Enum for the font scheme, used by implementations of the <see cref="IFont"/>
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
        /// Enum for the vertical alignment of the text from baseline, used by implementations of the <see cref="IFont"/>
        /// </summary>
        public enum VerticalTextAlignValue
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
        /// Enum for the style of the underline property of a stylized text, used by implementations of the <see cref="IFont"/>
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

        /// <summary>
        /// Enum for the charset definitions of a font, used by implementations of the <see cref="IFont"/>
        /// </summary>
        public enum CharsetValue
        {
            /// <summary>
            /// Application-defined (any other value than the defined enum values; can be ignored)
            /// </summary>
            ApplicationDefined = -1,
            /// <summary>
            /// Charset according to iso-8859-1
            /// </summary>
            ANSI = 0,
            /// <summary>
            /// Default charset (not defined more specific)
            /// </summary>
            Default = 1,
            /// <summary>
            /// Symbols from the private Unicode range U+FF00 to U+FFFF, to display special characters in the range of U+0000 to U+00FF
            /// </summary>
            Symbols = 2,
            /// <summary>
            /// Macintosh charset, Standard Roman
            /// </summary>
            Macintosh = 77,
            /// <summary>
            /// Shift JIS charset (shift_jis)
            /// </summary>
            JIS = 128,
            /// <summary>
            /// Hangul charset (ks_c_5601-1987)
            /// </summary>
            Hangul = 129,
            /// <summary>
            /// Johab charset (KSC-5601-1992)
            /// </summary>
            Johab = 130,
            /// <summary>
            /// KBB charset (GB-2312)
            /// </summary>
            GKB = 134,
            /// <summary>
            /// Chinese Big Five charset
            /// </summary>
            Big5 = 136,
            /// <summary>
            /// Greek charset (windows-1253)
            /// </summary>
            Greek = 161,
            /// <summary>
            /// Turkish charset (iso-8859-9)
            /// </summary>
            Turkish = 162,
            /// <summary>
            /// Vietnamese charset (windows-1258)
            /// </summary>
            Vietnamese = 163,
            /// <summary>
            /// Hebrew charset (windows-1255)
            /// </summary>
            Hebrew = 177,
            /// <summary>
            /// Arabic charset (windows-1256)
            /// </summary>
            Arabic = 178,
            /// <summary>
            /// Baltic charset (windows-1257)
            /// </summary>
            Baltic = 186,
            /// <summary>
            /// Russian charset (windows-1251)
            /// </summary>
            Russian = 204,
            /// <summary>
            /// Thai charset (windows-874)
            /// </summary>
            Thai = 222,
            /// <summary>
            /// Eastern Europe charset (windows-1250)
            /// </summary>
            EasternEuropean = 238,
            /// <summary>
            /// OEM characters, not defined by ECMA-376
            /// </summary>
            OEM = 255
        }

        /// <summary>
        /// Enum for the font family, according to the simple type definition of W3C. Used by implementations of the <see cref="IFont"/>
        /// </summary>
        public enum FontFamilyValue
        {
            /// <summary>
            /// The family is not defined or not applicable
            /// </summary>
            NotApplicable = 0,
            /// <summary>
            /// The specified font implements a Roman font
            /// </summary>
            Roman = 1,
            /// <summary>
            /// The specified font implements a Swiss font
            /// </summary>
            Swiss = 2,
            /// <summary>
            /// The specified font implements a Modern font
            /// </summary>
            Modern = 3,
            /// <summary>
            /// The specified font implements a Script font
            /// </summary>
            Script = 4,
            /// <summary>
            /// The specified font implements a Decorative font
            /// </summary>
            Decorative = 5,
            /// <summary>
            /// The specified font implements a not yet defined font archetype (reserved / do not use)
            /// </summary>
            Reserved1 = 6,
            /// <summary>
            /// The specified font implements a not yet defined font archetype (reserved / do not use)
            /// </summary>
            Reserved2 = 7,
            /// <summary>
            /// The specified font implements a not yet defined font archetype (reserved / do not use)
            /// </summary>
            Reserved3 = 8,
            /// <summary>
            /// The specified font implements a not yet defined font archetype (reserved / do not use)
            /// </summary>
            Reserved4 = 9,
            /// <summary>
            /// The specified font implements a not yet defined font archetype (reserved / do not use)
            /// </summary>
            Reserved5 = 10,
            /// <summary>
            /// The specified font implements a not yet defined font archetype (reserved / do not use)
            /// </summary>
            Reserved6 = 11,
            /// <summary>
            /// The specified font implements a not yet defined font archetype (reserved / do not use)
            /// </summary>
            Reserved7 = 12,
            /// <summary>
            /// The specified font implements a not yet defined font archetype (reserved / do not use)
            /// </summary>
            Reserved8 = 13,
            /// <summary>
            /// The specified font implements a not yet defined font archetype (reserved / do not use)
            /// </summary>
            Reserved9 = 14,
        }
        #endregion

        #region privateFields
        private float size;
        private string name = DEFAULT_FONT_NAME;
        //TODO: V3> Refactor to enum according to specs
        //OOXML: Chp.20.1.6.2(p2839ff)
        private string colorValue = "";
        private ColorSchemeElement colorTheme;
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
        /// Gets or sets the underline style of the font. If set to <see cref="UnderlineValue.none">none</a> no underline will be applied (default)
        /// </summary>
        [Append]
        public UnderlineValue Underline { get; set; } = UnderlineValue.none;

        /// <summary>
        /// Gets or sets the char set of the Font
        /// </summary>
        [Append]
        //TODO: v3> Refactor to enum according to specs
        // OOXML: Chp.19.2.1.13
        public CharsetValue Charset { get; set; } = CharsetValue.Default;

        /// <summary>
        /// Gets or sets the font color theme, represented by a color scheme
        /// </summary>
        [Append]
        //TODO: v3> Reference to Theming
        //OOXML: Chp.18.8.3 and 20.1.6.2
        public ColorSchemeElement ColorTheme
        {
            get => colorTheme;
            set
            {
                if (value == null)
                {
                    throw new StyleException("A color theme cannot be null");
                }
                colorTheme = value;
            }
        }
        /// <summary>
        /// Gets or sets the color code of the font color. The value is expressed as hex string with the format AARRGGBB. AA (Alpha) is usually FF.
        /// To omit the color, an empty string can be set. Empty is also default.
        /// </summary>
        /// <exception cref="StyleException">Throws a StyleException if the passed ARGB value is not valid</exception>
        [Append]
        public string ColorValue
        {
            get => colorValue;
            set
            {
                Validators.ValidateColor(value, true, true);
                colorValue = value;
            }
        }
        /// <summary>
        ///  Gets or sets the font family (Default is 2 = Swiss)
        /// </summary>
        [Append]
        //TODO: v3> Refactor to enum according to specs (18.18.94)
        //OOXML: Chp.18.8.18 and 18.18.94
        public FontFamilyValue Family { get; set; }
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
        /// \remark <remarks>Note that the font name is not validated whether it is a valid or existing font. The font name may not exceed more than 31 characters</remarks>
        [Append]
        public string Name //OOXML: Chp.18.8.29
        {
            get { return name; }
            set
            {
                name = value;
                ValidateFontScheme();
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
        public VerticalTextAlignValue VerticalAlign { get; set; }

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
            ColorTheme = ColorSchemeElement.light1;
            ColorValue = string.Empty;
            Scheme = DEFAULT_FONT_SCHEME;
            VerticalAlign = DEFAULT_VERTICAL_ALIGN;
        }
        #endregion

        #region methods  

        /// <summary>
        /// Validates the font name and sets the scheme automatically
        /// </summary>
        private void ValidateFontScheme()
        {
            if ((string.IsNullOrEmpty(name)) && !StyleRepository.Instance.ImportInProgress)
            {
                throw new StyleException("The font name was null or empty");
            }
            if (name.Equals(DEFAULT_MINOR_FONT))
            {
                Scheme = SchemeValue.minor;
            }
            else if (name.Equals(DEFAULT_MAJOR_FONT))
            {
                Scheme = SchemeValue.major;
            }
            else
            {
                Scheme = SchemeValue.none;
            }
        }

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
            hashCode = hashCode * -1521134295 + Charset.GetHashCode();
            hashCode = hashCode * -1521134295 + ColorTheme.GetHashCode();
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(ColorValue);
            hashCode = hashCode * -1521134295 + Family.GetHashCode();
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
