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
        /// <param name="scope">Optional parameter to validate for a specific address scope (Any, SingleAddress, Range). Default is: Any</param>
        /// <exception cref="FormatException">A format exception is thrown if the passed address is not a valid cell address or range</exception>
        /// \remark <remarks>If <paramref name="scope"/> is <see cref="Cell.AddressScope.Range"/>, an explicit range expression is required; a single address is rejected even though <see cref="Cell.ResolveCellRange(string)"/> can represent it as a one-cell range. If the scope is <see cref="Cell.AddressScope.Invalid"/>, the validation is inverted, so that a valid cell or range will throw an exception.</remarks>
        public static void ValidateCellAddressExpression(string expression, Cell.AddressScope scope = Cell.AddressScope.Any)
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
            if (scope == Cell.AddressScope.Range && isCellAddress && isRange)
            {
                System.FormatException innerException = new System.FormatException("The expression (" + expression + ") is a single cell address, but a cell range was expected");
                throw new FormatException(innerException.Message, innerException);
            }
            else if (scope == Cell.AddressScope.Any && !isCellAddress && !isRange)
            {
                throw new FormatException(lastException.Message, lastException); // Not a cell or range
            }
            else if (scope == Cell.AddressScope.Invalid && (isCellAddress || isRange))
            {
                throw new FormatException("The passed expression is valid cell address or range, but the validation was explicitly inverted");
            }
        }

        /// <summary>
        /// Validates the passed string, whether it is an expression that can be used as worksheet name
        /// </summary>
        /// <param name="name">Name to validate</param>
        /// <exception cref="NanoXLSX.Exceptions.FormatException">Throws a FormatException if the worksheet name is too long (max. 31) or contains illegal characters [  ]  * ? / \</exception>
        public static void ValidateWorksheetName(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                throw new FormatException("the worksheet name must be between 1 and " + Worksheet.MaxWorksheetNameLength + " characters");
            }
            if (name.Length > Worksheet.MaxWorksheetNameLength)
            {
                throw new FormatException("the worksheet name must be between 1 and " + Worksheet.MaxWorksheetNameLength + " characters");
            }
            Regex regex = new Regex(@"[\[\]\*\?/\\]");
            Match match = regex.Match(name);
            if (match.Captures.Count > 0)
            {
                throw new FormatException(@"the worksheet name must not contain the characters [  ]  * ? / \ ");
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
