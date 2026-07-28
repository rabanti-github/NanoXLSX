/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;

namespace NanoXLSX.Enums
{
    /// <summary>
    /// Static class that contains shared enums for error cases 
    /// </summary>
    public static class Errors
    {
        /// <summary>
        /// Errors that can occur in formulas / functions
        /// </summary>
        public enum FormulaError
        {
            /// <summary>
            /// Default value if no error has occurred (not an actual error type)
            /// </summary>
            NoError,
            /// <summary>
            /// Value if a not yet defined value has occurred (not an actual error type)
            /// </summary>
            UnknownError,
            /// <summary>
            /// Indicates that two areas are required to intersect, but do not.
            /// </summary>
            Null,
            /// <summary>
            ///  Indicates that any number (including zero) or any error code is divided by zero.
            /// </summary>
            DivisionByZero,
            /// <summary>
            ///  Indicates that an incompatible type argument is passed to a function, or an incompatible type operand is used with an operator. 
            /// </summary>
            Value,
            /// <summary>
            /// Indicate that a cell reference cannot be evaluated. 
            /// </summary>
            Reference,
            /// <summary>
            /// Indicates that what looks like a name is used, but no such name has been defined. 
            /// </summary>
            Name,
            /// <summary>
            /// Indicates that an argument to a function has a compatible type, but has a value that is outside the domain over which that function is defined.
            /// </summary>
            /// \remark <remarks>This is known as a domain error. In contrast to <see cref="FormulaError.Value"/>, a formula or function may look valid, but tries to handle an invalid value of the valid type. Example: <c>ATANH(50)</c></remarks>
            Number,
            /// <summary>
            /// Indicates that a designated value is not available. 
            /// </summary>
            /// \remark <remarks>Can (f.e.) happen if a formula requires two arrays of the same length, but the provided ones have different lengths. Example: <c>SUMX2MY2(A1:A3;B1:B4)</c></remarks>
            NotAvailable,
            /// <summary>
            /// Indicate that a cell reference cannot be evaluated because the value for the cell has not been retrieved or calculated.
            /// </summary>
            /// \remark <remarks>In contrast to <see cref="FormulaError.NotAvailable"/>, the value will eventually be available (e.g. by an external source).</remarks>
            GettingData
        }

        //TODO Expose public?
        /// <summary>
        /// Tries to parse a <see cref="FormulaError"/> from a string
        /// </summary>
        /// <param name="value">String to parse</param>
        /// <param name="error">Parsed error as out parameter (default to <see cref="FormulaError.UnknownError"/>)</param>
        /// <returns>True if the error could be parsed</returns>
        /// \remark <remarks>Errors within formula expressions cannot be parsed with this method. Such expressions have to be tokenized first.</remarks>
        internal static bool TryParseFormulaError(string value, out FormulaError error)
        {
            switch (value)
            {
                case "#NULL!":
                    error = FormulaError.Null;
                    return true;
                case "#DIV/0!":
                    error = FormulaError.DivisionByZero;
                    return true;
                case "#VALUE!":
                    error = FormulaError.Value;
                    return true;
                case "#REF!":
                    error = FormulaError.Reference;
                    return true;
                case "#NAME?":
                    error = FormulaError.Name;
                    return true;
                case "#NUM!":
                    error = FormulaError.Number;
                    return true;
                case "#N/A":
                    error = FormulaError.NotAvailable;
                    return true;
                case "#GETTING_DATA":
                    error = FormulaError.GettingData;
                    return true;
                default:
                    error = FormulaError.UnknownError;
                    return false;
            }
        }

        //TODO Expose public?
        /// <summary>
        /// Returns the OOXML conform error expression as string
        /// </summary>
        /// <param name="error">Enum value</param>
        /// <returns>OOXML internal error expression as string</returns>
        /// <exception cref="FormatException">Thrown if invalid values are passed, like <see cref="FormulaError.NoError"/></exception>
        internal static string FormulaErrorToString(FormulaError error)
        {
            switch (error)
            {
                case FormulaError.Null:
                    return "#NULL!";
                case FormulaError.DivisionByZero:
                    return "#DIV/0!";
                case FormulaError.Value:
                    return "#VALUE!";
                case FormulaError.Reference:
                    return "#REF!";
                case FormulaError.Name:
                    return "#NAME?";
                case FormulaError.Number:
                    return "#NUM!";
                case FormulaError.NotAvailable:
                    return "#N/A";
                case FormulaError.GettingData:
                    return "#GETTING_DATA";
                default:
                    throw new FormatException($"An invalid error type '{error}' was specified");
            }
        }
    }
}
