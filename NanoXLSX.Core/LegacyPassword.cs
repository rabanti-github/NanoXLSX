/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Security;
using NanoXLSX.Interfaces;
using NanoXLSX.Utils;

namespace NanoXLSX
{
    /// <summary>
    /// Legacy Excel protection password with explicit cleanup of owned storage.
    /// Input and returned strings cannot be erased by this instance.
    /// Not thread-safe.
    /// </summary>
    public class LegacyPassword : IPassword, IDisposable
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
            Type = type;
            PasswordHash = null;
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

            string replacementHash = GenerateLegacyPasswordHash(plainText);
            SecureString replacement = GetSecureString(plainText);
            ReplacePassword(replacement, replacementHash);
        }

        /// <summary>
        /// Removes the password form the current instance
        /// </summary>
        public void UnsetPassword()
        {
            ClearPassword();
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
            if (passwordInstance == null)
            {
                throw new ArgumentNullException(nameof(passwordInstance));
            }

            string replacementHash = passwordInstance.PasswordHash;
            bool copyType = GetType() == passwordInstance.GetType();
            PasswordType replacementType = copyType
                ? ((LegacyPassword)passwordInstance).Type
                : Type;
            SecureString replacement;

            if (passwordInstance.GetType() == typeof(LegacyPassword))
            {
                SecureString source = ((LegacyPassword)passwordInstance).password;
                replacement = source == null ? null : source.Copy();
            }
            else
            {
                replacement = GetSecureString(passwordInstance.GetPassword());
            }

            ReplacePassword(replacement, replacementHash);
            Type = replacementType;
        }

        /// <summary>
        /// Clears owned password storage. Repeated disposal is harmless.
        /// If reused afterwards, newly assigned storage requires cleanup again.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Clears owned password storage. Repeated disposal is harmless.
        /// If reused afterwards, newly assigned storage requires cleanup again.
        /// </summary>
        /// <param name="disposing">If true, the password will be unset, otherwise no actions are taken</param>
        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                UnsetPassword();
            }
        }

        /// <summary>
        /// Method to clear the password internally
        /// </summary>
        private void ClearPassword()
        {
            SecureString previous = password;
            password = null;
            if (previous != null)
            {
                previous.Dispose();
            }
        }

        /// <summary>
        /// Method to replace password with its hash internally
        /// </summary>
        /// <param name="replacement">Replacement password</param>
        /// <param name="replacementHash">Associated hash</param>
        private void ReplacePassword(SecureString replacement, string replacementHash)
        {
            ClearPassword();
            password = replacement;
            PasswordHash = replacementHash;
        }

        /// <summary>
        /// Generates the legacy Excel protection hash, not a cryptographic
        /// password hash. Null and empty inputs produce an empty string.
        /// </summary>
        /// \remark <remarks>WARNING! Do not reuse this method to encrypt 'real' passwords or data outside from NanoXLSX. This is only a minor security feature. Use a proper cryptography method instead.</remarks>
        /// <param name="password">Password string in UTF-8 to encrypt. Null or an empty string (even technical valid) are not allowed, since they cannot be inserted in a password field in Excel</param>
        /// <returns>16 bit hash as hex string. If the passed plain text password is null or empty, the returned hash will be empty</returns>
        public static string GenerateLegacyPasswordHash(string password)
        {
            if (string.IsNullOrEmpty(password)) { return string.Empty; }
            int passwordLength = password.Length;
            int passwordHash = 0;
            for (int i = passwordLength; i > 0; i--)
            {
                char character = password[i - 1];
                passwordHash = ((passwordHash >> 14) & 0x01)
                    | ((passwordHash << 1) & 0x7fff);
                passwordHash ^= character;
            }
            passwordHash = ((passwordHash >> 14) & 0x01)
                | ((passwordHash << 1) & 0x7fff);
            passwordHash ^= (0x8000 | ('N' << 8) | 'K');
            passwordHash ^= passwordLength;
            return passwordHash.ToString("X", ParserUtils.InvariantCulture);
        }

        /// <summary>
        /// Method to convert a string into a <see cref="SecureString"/>, to keep a plain text password as secure as possible in memory 
        /// </summary>
        /// <param name="plaintextPassword">Pain text</param>
        /// <returns>SecureString instance of the pain text</returns>
        private static SecureString GetSecureString(string plaintextPassword)
        {
            SecureString result = new SecureString();
            try
            {
                if (!string.IsNullOrEmpty(plaintextPassword))
                {
                    foreach (char character in plaintextPassword)
                    {
                        result.AppendChar(character);
                    }
                }
                return result;
            }
            catch
            {
                result.Dispose();
                throw;
            }
        }

        /// <summary>
        /// Method to get a plain text password from its secure string
        /// </summary>
        /// <param name="secureString">Secure string holding the password</param>
        /// <returns>Plain text password</returns>
        private static string GetPasswordOfSecureString(SecureString secureString)
        {
            IntPtr unmanagedString = IntPtr.Zero;
            try
            {
                unmanagedString = Marshal.SecureStringToGlobalAllocUnicode(secureString);
                // Preserve the snapshot's null-terminated string conversion behavior.
                return Marshal.PtrToStringUni(unmanagedString);
            }
            finally
            {
                if (unmanagedString != IntPtr.Zero)
                {
                    Marshal.ZeroFreeGlobalAllocUnicode(unmanagedString);
                }
            }
        }

        /// <summary>
        /// Returns whether two instances are the same
        /// </summary>
        /// <param name="obj">Object to compare</param>
        /// <returns>True if this instance and the other are the same</returns>
        public override bool Equals(object obj)
        {
            return obj is LegacyPassword other
                && Comparators.CompareSecureStrings(password, other.password)
                && Type == other.Type
                && PasswordHash == other.PasswordHash;
        }

        /// <summary>
        /// Gets the hash code of the password instance. Note that this is not the actual password hash
        /// </summary>
        /// <returns>Hash code of the password instance</returns>
        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = 1034998357;
                hashCode = hashCode * -1521134295 + Type.GetHashCode();
                hashCode = hashCode * -1521134295
                    + EqualityComparer<string>.Default.GetHashCode(PasswordHash);
                hashCode = hashCode * -1521134295 + PasswordIsSet().GetHashCode();
                return hashCode;
            }
        }
    }
}
