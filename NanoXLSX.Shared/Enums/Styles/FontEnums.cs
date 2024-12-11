/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2024
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Shared.Interfaces.Styles;

namespace NanoXLSX.Styles
{
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
}
