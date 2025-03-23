/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

namespace NanoXLSX.Interfaces
{
    /// <summary>
    /// Interface to represent a protection password, either for workbooks or worksheets. The implementations will define the password algorithms
    /// </summary>
    public interface IPassword
    {
        /// <summary>
        /// Gets or sets the password hash
        /// </summary>
        string PasswordHash { get; set; }

        /// <summary>
        /// Sets the plain text password
        /// </summary>
        /// <param name="plainText">Password in plain text</param>
        void SetPassword(string plainText);

        /// <summary>
        /// Unsets a previously defined password
        /// </summary>
        void UnsetPassword();

        /// <summary>
        /// Gets the password as plain text
        /// </summary>
        /// <returns>Password as plain text</returns>
        string GetPassword();

        /// <summary>
        /// Gets whether a password was set or not
        /// </summary>
        /// <returns></returns>
        bool PasswordIsSet();

        /// <summary>
        /// Method to copy a password instance from another one
        /// </summary>
        /// <param name="passwordInstance">Source instance</param>
        void CopyFrom(IPassword passwordInstance);
    }
}
