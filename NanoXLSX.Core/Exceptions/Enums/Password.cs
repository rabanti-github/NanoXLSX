/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

namespace NanoXLSX.Enums
{
    /// <summary>
    /// Static class that contains shared enums for password handling, during read or write operations
    /// </summary>
    public static class Password
    {
        /// <summary>
        /// Target type of the password
        /// </summary>
        public enum PasswordType
        {
            /// <summary>
            /// Password is used to protect a workbook
            /// </summary>
            WorkbookProtection,
            /// <summary>
            /// Password is used to protect a worksheet
            /// </summary>
            WorksheetProtection
        }
    }
}
