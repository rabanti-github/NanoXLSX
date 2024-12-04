/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2024
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Shared.Interfaces;
using NanoXLSX.Shared.Exceptions;
using NanoXLSX.Shared.Utils;
using System;
using System.Collections.Generic;
using System.Text;
using static NanoXLSX.Shared.Enums.Styles.FontEnums;
using NanoXLSX.Themes;
using NanoXLS.Shared.Enums.Schemes;

namespace NanoXLSX.Styles
{
    /// <summary>
    /// Class representing a Font entry. The Font entry is used to define text formatting
    /// </summary>
    public class Font : AbstractStyle, IFont
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
        /// The default font name that is declared as Major Font (See <see cref="Font.SchemeValue"/>)
        /// </summary>
        public static readonly string DEFAULT_MAJOR_FONT = "Calibri Light";
        /// <summary>
        /// The default font name that is declared as Minor Font (See <see cref="Font.SchemeValue"/>)
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

        #endregion

        #region privateFields
        private float size;
        private string name = DEFAULT_FONT_NAME;
        //TODO: V3> Refactor to enum according to specs
        //OOXML: Chp.20.1.6.2(p2839ff)
        private string colorValue = "";
        private ThemeEnums.ColorSchemeElement colorTheme;
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
        //TODO: v3> Reeference to Theming
        //OOXML: Chp.18.8.3 and 20.1.6.2
        public ThemeEnums.ColorSchemeElement ColorTheme { 
            get => colorTheme; 
            set {
                if (value == null)
                {
                    throw new StyleException("A color theme cannot be null");
                }
                colorTheme = value; 
            } }
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
        /// <remarks>Note that the font name is not validated whether it is a valid or existing font. The font name may not exceed more than 31 characters</remarks>
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
            ColorTheme = ThemeEnums.ColorSchemeElement.light1;
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
