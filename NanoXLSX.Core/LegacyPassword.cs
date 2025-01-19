/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using NanoXLSX.Interfaces;
using NanoXLSX.Utils;

namespace NanoXLSX
{
    /// <summary>
    /// Class implementing a legacy password, based on the proprietary hashing algorithm of Excel
    /// </summary>
    public class LegacyPassword : IPassword
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

        private SecureString password;

        /// <summary>
        /// Current target type of the password instance
        /// </summary>
        public PasswordType Type { get; set; }

        /// <summary>
        /// Gets or sets the hashed password
        /// </summary>
        /// <returns></returns>
        public string PasswordHash { get; set; }

        /// <summary>
        /// Default constructor with parameter
        /// </summary>
        /// <param name="type">Current target type of the password instance</param>
        public LegacyPassword(PasswordType type)
        {
            this.Type = type;
            this.PasswordHash = null;
        }

        /// <summary>
        /// Gets the pain text password
        /// </summary>
        /// <returns>Plain text password as string, or null, if not defined</returns>
        public string GetPassword()
        {
            if (password != null && password.Length > 0)
            {
                return GetPasswordOfSecureString(password);
            }
            return null;
        }

        /// <summary>
        /// Sets the current password and calculates the hash
        /// </summary>
        /// <param name="plainText">Plain text password. If null or empty, the password will be unset</param>
        public void SetPassword(string plainText)
        {
            if (string.IsNullOrEmpty(plainText))
            {
                UnsetPassword();
                return;
            }
            else
            {
                password = GetSecureString(plainText);
                PasswordHash = GenerateLegacyPasswordHash(plainText);
            }
        }

        /// <summary>
        /// Removes the password form the current instance
        /// </summary>
        public void UnsetPassword()
        {
            if (password != null)
            {
                password.Clear();
            }
            PasswordHash = null;
        }

        /// <summary>
        /// Gets whether a password was set
        /// </summary>
        /// <returns>True if a password was defined, false otherwise</returns>
        public bool PasswordIsSet()
        {
            return !string.IsNullOrEmpty(PasswordHash);
        }

        /// <summary>
        /// Copes all data from another class instance
        /// </summary>
        /// <param name="passwordInstance">Other instance (source)</param>
        public void CopyFrom(IPassword passwordInstance)
        {
            this.PasswordHash = passwordInstance.PasswordHash;
            this.password = GetSecureString(passwordInstance.GetPassword());
            if (this.GetType() == passwordInstance.GetType())
            {
                this.Type = ((LegacyPassword)passwordInstance).Type;
            }
        }

        /// <summary>
        /// Method to generate a legacy (Excel internal) password hash, to protect workbooks or worksheets<br />This method is derived from the c++ implementation by Kohei Yoshida (<a href="http://kohei.us/2008/01/18/excel-sheet-protection-password-hash/">http://kohei.us/2008/01/18/excel-sheet-protection-password-hash/</a>)
        /// </summary>
        /// \remark <remarks>WARNING! Do not use this method to encrypt 'real' passwords or data outside from NanoXLSX. This is only a minor security feature. Use a proper cryptography method instead.</remarks>
        /// <param name="password">Password string in UTF-8 to encrypt. Null or an empty string (even technical valid) are not allwd, since they cannot be inserted in a password field in Excel</param>
        /// <returns>16 bit hash as hex string. If the passed plain text password is null or empty, the returned hash will be empty</returns>
        public static string GenerateLegacyPasswordHash(string password)
        {
            if (string.IsNullOrEmpty(password)) { return string.Empty; }
            int passwordLength = password.Length;
            int passwordHash = 0;
            char character;
            for (int i = passwordLength; i > 0; i--)
            {
                character = password[i - 1];
                passwordHash = ((passwordHash >> 14) & 0x01) | ((passwordHash << 1) & 0x7fff);
                passwordHash ^= character;
            }
            passwordHash = ((passwordHash >> 14) & 0x01) | ((passwordHash << 1) & 0x7fff);
            passwordHash ^= (0x8000 | ('N' << 8) | 'K');
            passwordHash ^= passwordLength;
            return passwordHash.ToString("X");
        }

        /// <summary>
        /// Method to convert a string into a <see cref="SecureString"/>, to keep a plain text password as secure as possible in memory 
        /// </summary>
        /// <param name="plaintextPassword">Pain text</param>
        /// <returns>SecureString instance of the pain text</returns>
        private static SecureString GetSecureString(string plaintextPassword)
        {
            char[] chars;
            if (string.IsNullOrEmpty(plaintextPassword))
            {
                chars = new char[0];
            }
            else
            {
                chars = plaintextPassword.ToCharArray();
            }
            SecureString str = new SecureString();
            foreach (char c in chars)
            {
                str.AppendChar(c);
            }
            return str;
        }

        /// <summary>
        /// Method to retrieve the plain text from a <see cref="SecureString"/>
        /// </summary>
        /// <param name="secureString">SecureString instance. Cannot be null or empty</param>
        /// <returns>Plain text or null, if no SecureString was defined</returns>
        private static string GetPasswordOfSecureString(SecureString secureString)
        {
            IntPtr unmanagedString = IntPtr.Zero;
            try
            {
                unmanagedString = Marshal.SecureStringToGlobalAllocUnicode(secureString);
                return Marshal.PtrToStringUni(unmanagedString);
            }
            finally
            {
                Marshal.ZeroFreeGlobalAllocUnicode(unmanagedString);
            }
        }

        public override bool Equals(object obj)
        {
            LegacyPassword pwd = (LegacyPassword)obj;
            return obj is LegacyPassword password &&
                   Comparators.CompareSecureStrings(this.password, password.password) &&
                   Type == password.Type &&
                   PasswordHash == password.PasswordHash;
        }

        public override int GetHashCode()
        {
            // The actual password is not considered since its hash is sufficient
            var hashCode = 1034998357;
            hashCode = hashCode * -1521134295 + Type.GetHashCode();
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(PasswordHash);
            hashCode = hashCode * -1521134295 + PasswordIsSet().GetHashCode();
            return hashCode;
        }
    }
}
