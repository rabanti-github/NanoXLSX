/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;

namespace NanoXLSX
{
    public class PasswordProtection
    {
        public enum Algorithm
        {
            /// <summary>
            /// The old, proprietary password algorithm, by Excel
            /// </summary>
            /// \remark <remarks>This algorithm is rather weak, and should only be used for compatibility reasons</remarks>
            LEGACY,
            /// <summary>
            /// MD2 algorithm, defined by RFC 1319
            /// </summary>
            /// \remark <remarks>This algorithm is not recommended for new values, due to known flaws (breaks)</remarks>
            MD2,
            /// <summary>
            /// MD4 algorithm, defined by RFC 1320
            /// </summary>
            /// \remark <remarks>This algorithm is not recommended for new values, due to known flaws (breaks)</remarks>
            MD4,
            /// <summary>
            /// MD5 algorithm, defined by RFC 1321
            /// </summary>
            /// \remark <remarks>This algorithm is not recommended for new values, due to known flaws (breaks)</remarks>
            MD5,
            /// <summary>
            /// RIPEMD-128 algorithm, defined by ISO/IEC 10118-3:2004
            /// </summary>
            /// \remark <remarks>This algorithm is not recommended for new values, due to known flaws (breaks)</remarks>
            RIPEMD_128,
            /// <summary>
            /// RIPEMD-160 algorithm, defined by ISO/IEC 10118-3:2004
            /// </summary>
            RIPEMD_160,
            /// <summary>
            /// SHA-1 algorithm, defined by ISO/IEC 10118-3:2004
            /// </summary>
            SHA_1,
            /// <summary>
            /// SHA-256 algorithm, defined by ISO/IEC 10118-3:2004
            /// </summary>
            SHA_256,
            /// <summary>
            /// SHA-384 algorithm, defined by ISO/IEC 10118-3:2004
            /// </summary>
            SHA_384,
            /// <summary>
            /// SHA-512 algorithm, defined by ISO/IEC 10118-3:2004
            /// </summary>
            SHA_512,
            /// <summary>
            /// WHIRLPOOL algorithm, defined by ISO/IEC 10118-3:2004
            /// </summary>
            WHIRLPOOL
        }

        /// <summary>
        /// The used algorithm of the password
        /// </summary>
        public Algorithm PasswordAlgorithm { get; set; }
        /// <summary>
        /// The used salt value of the password. A salt value is an arbitrary character sequence that is prepended to the plain text password of the user, right before the password hash is calculated. 
        /// It prevents so called dictionary (or rainbow table) attacks by obfuscating the hash value of the isolated password.
        /// In contrast to the plain text password, the salt must be disclosed to be able to calculate a comparison hash. Furthermore, a salt should never be reused when defining a new password. 
        /// </summary>
        /// \remark <remarks>The LEGACY algorithm doesn't uses a salt value</remarks>
        public string Salt { get; set; }

        /// <summary>
        /// Number of iterations that a password algorithm is using, to calculate the password hash. The larger the number, the longer it takes to calculate a hash, what increases the security but decreases the application performance during the calculation.
        /// </summary>
        /// \remark <remarks>The LEGACY algorithm doesn't uses a spin count</remarks>
        public int SpinCount { get; set; }

        public string GeneratePasswordHash(string plainTextPassword)
        {
            using (var rfc2898 = new Rfc2898DeriveBytes(plainTextPassword, Encoding.UTF8.GetBytes(Salt), SpinCount))
            {
                return null;
               // return Convert.ToBase64String(rfc2898.GetBytes(HashSize));
            }
        }

  

        /// <summary>
        /// Method to create a random string with the given length, to be used as salt
        /// </summary>
        /// <param name="lenght">Number of characters</param>
        /// <returns></returns>
        public static string GenerateSalt(int lenght)
        {
            byte[] saltBytes = new byte[lenght];
            using (RandomNumberGenerator generator = RandomNumberGenerator.Create())
            {
                generator.GetBytes(saltBytes);
            }
            return Convert.ToBase64String(saltBytes);
        }
    }
}
