/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

namespace NanoXLSX.Interfaces
{
    /// <summary>
    /// Interface used by text partitions of option classes (e.g. ReaderOptions)
    /// </summary>
    public interface ITextOptions
    {

        /// <summary>
        /// If true, phonetic characters (like ruby characters / Furigana / Zhuyin fuhao) in strings are added in brackets after the transcribed symbols. By default, phonetic characters are removed from strings.
        /// </summary>
        /// \remark <remarks>This option is not applicable to specific rows or a start column (applied globally)</remarks>
        bool EnforcePhoneticCharacterImport { get; set; }

        /// <summary>
        /// If true, empty cells will be interpreted as type of string with an empty value. If false, the type will be Empty and the value null
        /// </summary>
        bool EnforceEmptyValuesAsString { get; set; }
    }
}
