using System;
using System.Runtime.InteropServices;
using System.Security;
using System.Text.RegularExpressions;
using NanoXLSX.Exceptions;


namespace NanoXLSX.Utils
{
    /// <summary>
    /// Class providing general validator methods 
    /// </summary>
    public static class Validators
    {
        /// <summary>
        /// Validates the passed string, whether it is a valid RGB or ARGB value that can be used for Fills or Fonts
        /// </summary>
        /// <exception cref="StyleException">A StyleException is thrown if an invalid hex value is passed</exception>
        /// <param name="hexCode">Hex string to check</param>
        /// <param name="useAlpha">If true, two additional characters (total 8) are expected as alpha value</param>
        /// <param name="allowEmpty">Optional parameter that allows null or empty as valid values</param>
        public static void ValidateColor(string hexCode, bool useAlpha, bool allowEmpty = false)
        {
            if (string.IsNullOrEmpty(hexCode))
            {
                if (allowEmpty)
                {
                    return;
                }
                throw new StyleException("The color expression was null or empty");
            }

            int length = useAlpha ? 8 : 6;
            if (hexCode.Length != length)
            {
                throw new StyleException("The value '" + hexCode + "' is invalid. A valid value must contain " + length + " hex characters");
            }
            if (!Regex.IsMatch(hexCode, "[a-fA-F0-9]{6,8}"))
            {
                throw new StyleException("The expression '" + hexCode + "' is not a valid hex value");
            }
        }

        /// <summary>
        /// Compares whether the content of two  <see cref="SecureString">SecureString</see> instances are equal. The comparison method tries to handle the operation as secure as possible
        /// </summary>
        /// <param name="value1">SecureString instance one</param>
        /// <param name="value2">SecureString instance two</param>
        /// <returns>True, if the content of the two instance is equal, otherwise false</returns>
        public static bool CompareSecureStrings(SecureString value1, SecureString value2)
        {
            if ((value1 == null && value2 != null)||(value1 != null && value2 == null))
            {
                return false;
            }
            IntPtr unmanagedString1 = IntPtr.Zero;
            IntPtr unmanagedString2 = IntPtr.Zero;
            try
            {
                unmanagedString1 = Marshal.SecureStringToBSTR(value1);
                unmanagedString2 = Marshal.SecureStringToBSTR(value2);
                int length1 = Marshal.ReadInt32(unmanagedString1, -4);
                int length2 = Marshal.ReadInt32(unmanagedString2, -4);
                if (length1 == length2)
                {
                    for (int i = 0; i < length1; ++i)
                    {
                        byte byte1 = Marshal.ReadByte(unmanagedString1, i);
                        byte byte2 = Marshal.ReadByte(unmanagedString2, i);
                        if (byte1 != byte2) return false;
                    }
                }
                else
                {
                    return false;
                }
                return true;
            }
            finally
            {
                // Cleanup
                if (unmanagedString2 != IntPtr.Zero) Marshal.ZeroFreeBSTR(unmanagedString2);
                if (unmanagedString1 != IntPtr.Zero) Marshal.ZeroFreeBSTR(unmanagedString1);
            }
        }
    }
}
