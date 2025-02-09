using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;

namespace NanoXLSX.Utils
{
    /// <summary>
    /// Class providing general comparator methods 
    /// </summary>
    public static class Comparators
    {

        /// <summary>
        /// Threshold, using when floats are compared
        /// </summary>
        private const float FLOAT_THRESHOLD = 0.0001f;

        /// <summary>
        /// Compares whether the content of two  <see cref="SecureString">SecureString</see> instances are equal. The comparison method tries to handle the operation as secure as possible
        /// </summary>
        /// <param name="value1">SecureString instance one</param>
        /// <param name="value2">SecureString instance two</param>
        /// <returns>True, if the content of the two instance is equal, otherwise false</returns>
        public static bool CompareSecureStrings(SecureString value1, SecureString value2)
        {
            bool v1Empty = false;
            bool v2Empty = false;
            if (value1 == null || value1.Length == 0)
            {
                v1Empty = true;
            }
            if (value2 == null || value2.Length == 0)
            {
                v2Empty = true;
            }
            if (v1Empty && !v2Empty || !v1Empty && v2Empty)
            {
                return false;
            }
            if (v1Empty && v2Empty)
            {
                return true;
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

        /// <summary>
        /// Compares two dimensions (e.g. column width, column height).
        /// </summary>
        /// <param name="dimension1">Nullable dimension 1</param>
        /// <param name="dimension2">Nullable dimension 2</param>
        /// <returns>1, if dimension1 is greater than dimension2, -1 if dimension1 is smaller than dimension2, 0 if both values are equal</returns>
        /// \remark <remarks>If dimension1 is null, -1 will be returned, if dimension2 is null, 1 will be returned. For the equality, a threshold value will be used</remarks>
        public static int CompareDimensions(float? dimension1, float? dimension2)
        {
            if (dimension1 == null)
            {
                dimension1 = float.MinValue;
            }
            if (dimension2 == null)
            {
                dimension2 = float.MinValue;
            }
            if (Math.Abs(dimension1.Value - dimension2.Value) < FLOAT_THRESHOLD)
            {
                return 0;
            }
            else if (dimension1 > dimension2)
            {
                return 1;
            }
            else 
            {
                return -1;
            }
        }
    }
}
