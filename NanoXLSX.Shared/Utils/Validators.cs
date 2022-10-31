using NanoXLSX.Shared.Exceptions;
using System.Text.RegularExpressions;


namespace NanoXLSX.Shared.Utils
{
    public static class Validators
    {
        /// <summary>
        /// Validates the passed string, whether it is a valid RGB value that can be used for Fills or Fonts
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

            int length;
            length = useAlpha ? 8 : 6;
            if (hexCode.Length != length)
            {
                throw new StyleException("The value '" + hexCode + "' is invalid. A valid value must contain six hex characters");
            }
            if (!Regex.IsMatch(hexCode, "[a-fA-F0-9]{6,8}"))
            {
                throw new StyleException("The expression '" + hexCode + "' is not a valid hex value");
            }
        }
    }
}
