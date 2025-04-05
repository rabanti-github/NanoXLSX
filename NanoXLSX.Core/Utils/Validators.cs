﻿using System.Text.RegularExpressions;
using NanoXLSX.Exceptions;


namespace NanoXLSX.Utils
{
    /// <summary>
    /// Class providing general validator methods 
    /// </summary>
    public static class Validators
    {

        /// <summary>
        /// Threshold, using when floats are compared
        /// </summary>
        private const float FLOAT_THRESHOLD = 0.0001f;

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
    }
}
