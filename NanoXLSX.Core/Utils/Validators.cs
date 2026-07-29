using System;
using System.Text.RegularExpressions;
using NanoXLSX.Exceptions;
using FormatException = NanoXLSX.Exceptions.FormatException;


namespace NanoXLSX.Utils
{
    /// <summary>
    /// Class providing general validator methods 
    /// </summary>
    public static class Validators
    {

        /// <summary>
        /// Validates the passed string, whether it is a valid RGB or ARGB value that can be used for Fills, Fonts or other styling components.
        /// The method automatically tries to validate for ARGB (8 characters) first, then for RGB (6 characters).
        /// </summary>
        /// <param name="hexCode">Hex string to check</param>
        /// <param name="allowEmpty">Optional parameter that allows null or empty as valid values</param>
        public static void ValidateGenericColor(string hexCode, bool allowEmpty = false)
        {
            string argbMessage = ValidateColorInternal(hexCode, true, allowEmpty);
            string rgbMessage = null;
            if (argbMessage != null)
            {
                rgbMessage = ValidateColorInternal(hexCode, false, allowEmpty);
                if (rgbMessage != null)
                {
                    throw new StyleException(argbMessage);
                }
            }
        }

        /// <summary>
        /// Validates the passed string, whether it is a valid RGB or ARGB value that can be used for Fills, Fonts or other styling components
        /// </summary>
        /// <exception cref="StyleException">A StyleException is thrown if an invalid hex value is passed</exception>
        /// <param name="hexCode">Hex string to check</param>
        /// <param name="useAlpha">If true, two additional characters (total 8) are expected as alpha value</param>
        /// <param name="allowEmpty">Optional parameter that allows null or empty as valid values</param>
        public static void ValidateColor(string hexCode, bool useAlpha, bool allowEmpty = false)
        {
            string message = ValidateColorInternal(hexCode, useAlpha, allowEmpty);
            if (message != null)
            {
                throw new StyleException(message);
            }
        }

        /// <summary>
        /// Validates the passed string, whether it is a valid single cell address or cell range. The address or range can contain modifier characters (<see cref="Cell.AddressType"/>) 
        /// </summary>
        /// <param name="expression">The address expression to validate</param>
        /// <param name="scope">Optional parameter to validate for a specific addres scope (Any, SingleAddress, Range). Default is: Any</param>
        /// <exception cref="FormatException">A format exception is thrown if the passed address is not a valid cell address or range</exception>
        /// /Remark <remarks>If <see cref="Cell.AddressScope"/> of the scope parameter is set to <see cref="Cell.AddressScope.Invalid"/>, it inverts the validation, so that a valid cell OR range will throw an exception</remarks>
        internal static void ValidateCellAddressExpression(string expression, Cell.AddressScope scope = Cell.AddressScope.Any)
        {
            bool isCellAddress = false;
            bool isRange = false;
            Exception lastException = null;
            try
            {
                Cell.ResolveCellCoordinate(expression);
                isCellAddress = true;
            }
            catch (Exception ex)
            {
                if (scope == Cell.AddressScope.SingleAddress)
                {
                    throw new FormatException(ex.Message, ex); // No further checks necessary
                }
                lastException = ex;
            }
            try
            {
                Cell.ResolveCellRange(expression);
                isRange = true;
            }
            catch (Exception ex)
            {
                if (scope == Cell.AddressScope.Range)
                {
                    throw new FormatException(ex.Message, ex); // No further checks necessary
                }
                lastException = ex;
            }
            if (scope == Cell.AddressScope.Any && !isCellAddress && !isRange)
            {
                throw new FormatException(lastException.Message, lastException); // Not a cell or range
            }
            else if (scope == Cell.AddressScope.Invalid && (isCellAddress || isRange))
            {
                throw new FormatException("The passed expression is valid cell address or range, but the validation was explicitly inverted");
            }
        }

        /// <summary>
        /// Validates the passed string, whether it is a valid RGB or ARGB value that can be used for Fills, Fonts or other styling components.
        /// </summary>
        /// <param name="hexCode">Hex string to check</param>
        /// <param name="useAlpha">If true, two additional characters (total 8) are expected as alpha value</param>
        /// <param name="allowEmpty">Optional parameter that allows null or empty as valid values</param>
        /// <returns>Null, if valid, otherwise, the specific exception message is returned</returns>
        private static string ValidateColorInternal(string hexCode, bool useAlpha, bool allowEmpty)
        {
            if (string.IsNullOrEmpty(hexCode))
            {
                if (allowEmpty)
                {
                    return null;
                }
                return "The color expression cannot be null or empty";
            }

            int length = useAlpha ? 8 : 6;
            if (hexCode.Length != length)
            {
                return "The value '" + hexCode + "' is invalid. A valid value must contain " + length + " hex characters";
            }
            if (!Regex.IsMatch(hexCode, "[a-fA-F0-9]{6,8}"))
            {
                return "The expression '" + hexCode + "' is not a valid hex value";
            }
            return null;
        }
    }
}
