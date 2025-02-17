/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using NanoXLSX.Interfaces;
using NanoXLSX.Interfaces.Writer;

namespace NanoXLSX.Internal.Writers
{
    /// <summary>
    /// Class representing a writer for legacy passwords
    /// </summary>
    public class LegacyPasswordWriter : IPasswordWriter
    {
        /// <summary>
        /// Target type of the password
        /// </summary>
        public enum PasswordType
        {
            /// <summary>
            /// Password is used to protect a workbook
            /// </summary>
            WORKBOOK_PROTECTION,
            /// <summary>
            /// Password is used to protect a worksheet
            /// </summary>
            WORKSHEET_PROTECTION
        }

        private string passwordHash;

        /// <summary>
        /// Current target type of the password instance
        /// </summary>
        public PasswordType Type { get; private set; }

        /// <summary>
        /// Gets or sets the password hash
        /// </summary>
        public string PasswordHash
        {
            get { return passwordHash; }
            set { passwordHash = value; }
        }

        /// <summary>
        /// Default constructor with parameter
        /// </summary>
        /// <param name="type">Current target type of the password instance</param>
        /// <param name="hash">Hash representation of the password (do not use null)</param>
        public LegacyPasswordWriter(PasswordType type, string hash)
        {
            this.Type = type;
            this.PasswordHash = hash;
        }

        /// <summary>
        /// Not relevant for the writer (inherited from <see cref="IPassword"/>)
        /// </summary>
        /// <param name="passwordInstance">Source instance</param>
        /// <exception cref="NotImplementedException">Throws a NotImplementedException if called in any case</exception>
        public void CopyFrom(IPassword passwordInstance)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Gets the internally used XML attributes that are set when a workbook is saved
        /// </summary>
        /// <returns>XML attributes with a leading space, or an empty string, if no password was set</returns>
        /// \remark <remarks>This method is only internally used on writing workbooks</remarks>
        public new string GetXmlAttributes()
        {
            if (Type == PasswordType.WORKSHEET_PROTECTION)
            {
                return " password =\"" + passwordHash + "\"";
            }
            else
            {
                return " workbookPassword=\"" + passwordHash + "\"";
            }
        }

        /// <summary>
        /// Gets whether a password to write is defined
        /// </summary>
        /// <returns>True if a password is set to be written</returns>
        public bool PasswordIsSet()
        {
            return PasswordHash != null;
        }

        /// <summary>
        /// Not relevant for the writer (inherited from <see cref="IPassword"/>)
        /// </summary>
        /// <exception cref="NotImplementedException">Throws a NotImplementedException if called in any case</exception>
        public string GetPassword()
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Not relevant for the writer (inherited from <see cref="IPassword"/>)
        /// </summary>
        /// <param name="plainText"></param>
        /// <exception cref="NotImplementedException">Throws a NotImplementedException if called in any case</exception>
        public void SetPassword(string plainText)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Not relevant for the writer (inherited from <see cref="IPassword"/>)
        /// </summary>
        /// <exception cref="NotImplementedException">Throws a NotImplementedException if called in any case</exception>
        public void UnsetPassword()
        {
            throw new NotImplementedException();
        }
    }
}
