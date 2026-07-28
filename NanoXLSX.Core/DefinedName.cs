/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using NanoXLSX.Enums;
using NanoXLSX.Exceptions;
using NanoXLSX.Utils;
using static NanoXLSX.Enums.Errors;
using FormatException = NanoXLSX.Exceptions.FormatException;

namespace NanoXLSX
{
    /// <summary>
    /// Class representing a defined name within a workbook. A defined name is a descriptive text that
    /// represents a cell, a range of cells, a formula, or a constant value, and can be referenced from
    /// formulas in worksheets (e.g. <c>=SUM(MyRange)</c>).
    /// </summary>
    /// \remark <remarks>
    /// A defined name has a workbook scope by default (<see cref="LocalSheet"/> is null). When a
    /// <see cref="Worksheet"/> is supplied as <see cref="LocalSheet"/>, the defined name is scoped to
    /// that worksheet (corresponding to the <c>localSheetId</c> attribute in the OOXML representation).
    /// The <see cref="TextValue"/> is stored verbatim — NanoXLSX does not parse or evaluate it.
    /// </remarks>
    public sealed class DefinedName : IEquatable<DefinedName>, IComparable<DefinedName>
    {
        #region enums
        /// <summary>
        /// Enum to specify the type of the defined name
        /// </summary>
        public enum NameType
        {
            /// <summary>Defined name is a single cell </summary>
            Cell,
            /// <summary>Defined name is a cell range </summary>
            Range,
            /// <summary>Defined name is a formula </summary>
            Formula,
            /// <summary>Defined name is a constant value </summary>
            Constant

        }

        #endregion

        #region constants

        private static readonly Regex EXT_WORKSHEET_REFERENE_REGEX = new Regex(
        @"^\[[0-9]+\].+", RegexOptions.Compiled | RegexOptions.CultureInvariant);

        private static readonly Regex EXT_REFERENCE_REGEX = new Regex(
        @"\[[0-9]+\]", RegexOptions.CultureInvariant);

        private const int MAX_NAME_LENGTH = 255;

        /// <summary>
        /// Disallowed names for defined names (ignore case)
        /// </summary>
        private static readonly HashSet<string> DISALLOWED_NAMES = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
        "C",
        "R"
            };

        /// <summary>
        /// Allowed special characters at the start of a defined name
        /// </summary>
        private static readonly char[] ALLOWED_NAME_START_CHARS = { '\\', '_' };
        /// <summary>
        /// Allowed special characters after the first character of a defined name.
        /// \remark <remarks>'\' is accepted because Excel allows it, although it is not documented in the official naming rules.</remarks>
        /// </summary>
        private static readonly char[] ALLOWED_NAME_CHARS = { '_', '.', '\\' };

        #endregion

        #region properties

        /// <summary>
        /// Type of the defined name
        /// </summary>
        public NameType Type {get;}

        /// <summary>
        /// Gets the name of the defined name as it appears in the workbook (e.g. <c>MyRange</c>).
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Gets the target worksheet in case of Cell or Range values. For other types, like formulas or constants, the target worksheet is null 
        /// </summary>
        public Worksheet TargetWorksheet { get; }

        /// <summary>
        /// Gets the textual reference of the defined name. This is stored verbatim and may be a cell address
        /// (e.g. <c>$A$1</c>), a range (e.g. <c>$A$1:$A$10</c>), a formula (e.g. <c>SUM(Sheet1!$A$1:$A$10)</c>),
        /// or a constant value.
        /// </summary>
        /// \remark <remarks>Do not add the target worksheet name (e.g. <c>Sheet1</c>) in front of the Reference in case of cells or ranges. The worksheet is automatically added by the defined <see cref="TargetWorksheet"/></remarks>
        public string TextValue { get; private set; }

        /// <summary>
        /// Gets the raw reference of the defined Name. The value will be transformed in its appropriate text value (<see cref="TextValue"/>).
        /// If the object is not supported (integer, float, date / date and time, boolean or string), <see cref="Object.ToString"/> will be used to determine the text value
        /// </summary>
        public object Value { get; private set; }

        /// <summary>
        /// Gets the worksheet that scopes (constraint) this defined name. If null, the defined name has workbook scope and
        /// is visible from any worksheet. If non-null, the defined name is scoped to the referenced worksheet
        /// (mapped to the <c>localSheetId</c> attribute on save).
        /// </summary>
        public Worksheet LocalSheet { get; }

        /// <summary>
        /// Gets the optional comment associated with the defined name. Maps to the <c>comment</c> attribute in OOXML.
        /// May be null when no comment was set.
        /// </summary>
        public string Comment { get; }

        /// <summary>
        /// Gets a possible error of the whole value in a defined name. Default is <see cref="FormulaError.NoError"/>.
        /// </summary>
        /// \remark <remarks>Errors within formula expressions are not set to an error in this property. The default will be <see cref="FormulaError.NoError"/>.</remarks>
        public FormulaError Error { get; private set; }

        /// <summary>
        /// Gets whether the value contains a reference or multiple references to an external source (e.g. an external workbook)
        /// </summary>
        public bool HasExternalReferences { get; private set; }

        #endregion

        #region constructors

        /// <summary>
        /// Constructs a new defined name.
        /// </summary>
        /// <param name="workbook">Workbook reference (to check for duplicate names)</param>
        /// <param name="type">Type of the defined name</param>
        /// <param name="name">Name of the defined name. Must be non-empty, must not start with a digit, and must not match a cell reference in the range A1 - XFD1048576.</param>
        /// <param name="reference">Reference text (cell, range, formula, or constant). Must be non-empty.</param>
        /// <param name="worksheet">Target worksheet in case of cells or cell ranges</param>
        /// <param name="localSheet">Optional worksheet that scopes the defined name. Pass null for workbook scope.</param>
        /// <param name="comment">Optional comment.</param>
        /// <exception cref="FormatException">Thrown if <paramref name="workbook"/> is null.</exception>
        /// <exception cref="FormatException">Thrown if <paramref name="name"/> is null, empty, only whitespaces, starts with a digit or contains illegal characters.</exception>
        /// <exception cref="FormatException">Thrown if <paramref name="reference"/> is null or resolved to an empty string</exception>
        /// <exception cref="WorksheetException">Thrown if <paramref name="reference"/> already exists or matches a cell reference in the same scope</exception>
        /// \remark <remarks>Use <see cref="Workbook.AddDefinedNameCell(string, Worksheet, string, Worksheet, string)"/>, <see cref="Workbook.AddDefinedNameFormula(string, string, Worksheet, string)"/>, <see cref="Workbook.AddDefinedNameConstant(string, object, Worksheet, string)"/>, <see cref="Workbook.AddDefinedNameFormula(string, string, Worksheet, string)"/> (or available overloaded methods) to conveniently add defined names</remarks>
        internal DefinedName(Workbook workbook, NameType type, string name, object reference, Worksheet worksheet, Worksheet localSheet = null, string comment = null)
        {
            if (workbook == null)
            {
                throw new FormatException("To set a defined name, a workbook must be provided.");
            }
            ValidateName(workbook, name, localSheet);
            if (reference == null || string.IsNullOrEmpty(reference.ToString()))
            {
                throw new FormatException("The reference of a defined name must not be null or empty.");
            }
            this.Type = type;
            this.Name = name;
            this.Value = reference;
            this.TargetWorksheet = worksheet;
            this.LocalSheet = localSheet;
            this.Comment = comment;
            CastValue(workbook);
        }
        #endregion

        #region methods

        /// <summary>
        /// Validates that the supplied name is non-empty, does not start with a digit, an invalid name or character, and does not match
        /// a cell reference in the range A1 - XFD1048576.
        /// </summary>
        /// <param name="name">Name to validate.</param>
        /// <param name="workbook">Workbook to check for duplicate names</param>
        /// <param name="localSheet">Local worksheet reference. Can be null if workbook scope</param>
        /// <exception cref="FormatException">Thrown if validation fails.</exception>
        private static void ValidateName(Workbook workbook, string name, Worksheet localSheet)
        {
            if (string.IsNullOrEmpty(name))
            {
                throw new FormatException("The name of a defined name must not be null or empty.");
            }
            if (name.Length > MAX_NAME_LENGTH)
            {
                throw new FormatException($"A defined name must not exceed {MAX_NAME_LENGTH} characters.");
            }

            char firstChar = name[0];

            if (!char.IsLetter(firstChar)
                && !ALLOWED_NAME_START_CHARS.Contains(firstChar))
            {
                throw new FormatException($"The name of a defined name must start with a letter, underscore, or backslash. Provided: '{name}'");
            }
            if (DISALLOWED_NAMES.Contains(name))
            {
                throw new FormatException($"'{name}' cannot be used as a defined name.");
            }
            for (int i = 1; i < name.Length; i++)
            {
                char character = name[i];

                if (!char.IsLetterOrDigit(character) && !ALLOWED_NAME_CHARS.Contains(character))
                {
                    throw new FormatException($"The character '{character}' at position {i} is not valid in the defined name '{name}'.");
                }
            }
            if (workbook.FindDefinedNameIndex(name, localSheet) >= 0)
            {
                string scope = localSheet == null ? "workbook" : "worksheet '" + localSheet.SheetName + "'";
                throw new WorksheetException("A defined name with the name '" + name + "' already exists in the " + scope + " scope.");
            }
            try
            {
                Validators.ValidateCellAddressExpression(name, Cell.AddressScope.SingleAddress);
            }
            catch 
            {
                // Not a valid cell address; therefore it may be used as a defined name.
                return;
            }
            throw new FormatException($"The defined name '{name}' must not be a valid cell address.");
        }

        /// <summary>
        /// Casts <see cref="Value"/> to a valid string for <see cref="TextValue"/>
        /// </summary>
        /// <param name="workbook"></param>
        /// <exception cref="FormatException">Thrown if a expected address or range expression is invalid</exception>
        private void CastValue(Workbook workbook)
        {
            if (this.Value == null)
            {
                throw new FormatException("The value of a defined name cannot be null or empty");
            }
            switch (this.Type)
            {
                // The object type is assumed to be validated prior
                case NameType.Cell:
                    string address = this.Value as string;
                    Validators.ValidateCellAddressExpression(address, Cell.AddressScope.SingleAddress); // throw if not an address
                    Address fixedAddress = new Address(address, Cell.AddressType.FixedRowAndColumn);
                    this.TextValue = fixedAddress.ToString();
                    this.Value = fixedAddress; // Reformat passed object
                    break;
                case NameType.Range:
                    string range = this.Value as string;
                    Validators.ValidateCellAddressExpression(range, Cell.AddressScope.Range); // throw if not valid range
                    Range tempRange = new Range(range);
                    Range fixedRange = new Range(new Address(tempRange.StartAddress.Row, tempRange.StartAddress.Column, Cell.AddressType.FixedRowAndColumn), new Address(tempRange.EndAddress.Row, tempRange.EndAddress.Column, Cell.AddressType.FixedRowAndColumn));
                    this.TextValue = fixedRange.ToString();
                    break;
                case NameType.Formula:
                    this.TextValue = this.Value.ToString(); // No formula validation yet
                    break;
                default: // constant
                    this.TextValue = ParserUtils.ToCachedValueString(this.Value, false);
                    break;
            }
        }

        /// <summary>
        /// Resolve a defined name from its string reference 
        /// </summary>
        /// <param name="name">Name (cannot be null)</param>
        /// <param name="reference">String reference (cannot be null)</param>
        /// <param name="workbook">Workbook reference (for cells, ranges and worksheet resolution)</param>
        /// <param name="localSheet">Local sheet (can be null)</param>
        /// <param name="comment"> Comment (can be null)</param>
        /// <returns>Resolved defined name object</returns>
        internal static DefinedName ResolveDefinedName(string name, string reference, Workbook workbook, Worksheet localSheet, string comment)
        {
            string worksheetName;
            NameType type;
            FormulaError formulaError;
            object value = GetParsedObject(reference, out type, out worksheetName, out formulaError);
            bool containsExternalLink = ContainsExternalLink(worksheetName, type, value);
            Worksheet worksheet = null;
            if (worksheetName != null && !containsExternalLink)
            {
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    if (string.Equals(worksheetName, ws.SheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        worksheet = ws;
                        break;
                    }
                }
            }
            DefinedName definedName = new DefinedName(workbook, type, name, value, worksheet, localSheet, comment);
            definedName.Error = formulaError;
            definedName.HasExternalReferences = containsExternalLink;
            return definedName;
        }

        private static bool ContainsExternalLink(string worksheet, NameType type, object value)
        {
            switch (type)
            {
                case NameType.Formula:
                    string formula = value as string;
                    return EXT_REFERENCE_REGEX.IsMatch(formula);
                    case NameType.Range:
                    case NameType.Cell:
                     if (worksheet != null)
                    {
                        return EXT_WORKSHEET_REFERENE_REGEX.IsMatch(worksheet);
                    }
                    break;
                    default: // constant
                    break; // NoOp
            }
            return false;
        }

        private static object GetParsedObject(string reference, out NameType type, out string worksheet, out FormulaError error)
        {
            error = FormulaError.NoError;
            worksheet = null;
            if (ParserUtils.TryParseFormulaStringConstant(reference, out string stringValue))
            {
                type = NameType.Constant; // Formula string is interpreted as constant in this case
                return stringValue;
            }
            if (ParserUtils.TryParseBool(reference, out bool boolValue))
            {
                type = NameType.Constant;
                return boolValue;
            }
            if (ParserUtils.TryParseInt(reference, out int intValue))
            {
                type = NameType.Constant;
                return intValue;
            }
            if (ParserUtils.TryParseDouble(reference, out double doubleValue))
            {
                type = NameType.Constant;
                return doubleValue;
            }
            string worksheetName;
            string addressExpression;
            if (ParserUtils.TryParseWorksheetQualifiedReference(reference, out worksheetName, out addressExpression))
            {
                worksheet = worksheetName;
                try
                {
                    Address addressValue = new Address(addressExpression);
                    type = NameType.Cell;
                    return addressValue;
                }
                catch
                {
                    // NoOp
                }
                try
                {
                    Range range = new Range(addressExpression);
                    type = NameType.Range;
                    return range;
                }
                catch
                {
                    // NoOp
                }
            }
            FormulaError referenceError;
            if (Errors.TryParseFormulaError(reference, out referenceError))
            {
                error = referenceError;
            }
            type = NameType.Formula;
            return reference;
        }

        /// <summary>
        /// Determines whether the specified <see cref="DefinedName"/> instance is equal to the current instance.
        /// Two instances are considered equal when their <see cref="Name"/>, <see cref="TextValue"/>,
        /// <see cref="Comment"/>, and <see cref="LocalSheet"/> (compared by reference) match.
        /// </summary>
        /// <param name="other">Other defined name instance, or null.</param>
        /// <returns>True if equal, otherwise false.</returns>
        public bool Equals(DefinedName other)
        {
            if (other is null)
            {
                return false;
            }
            if (ReferenceEquals(this, other))
            {
                return true;
            }
            return string.Equals(Name, other.Name, StringComparison.Ordinal)
                && Enum.Equals(Type, other.Type)
                && string.Equals(TextValue, other.TextValue, StringComparison.Ordinal) // object implicit compared by string
                && string.Equals(Comment, other.Comment, StringComparison.Ordinal)
                && ReferenceEquals(TargetWorksheet, other.TargetWorksheet)
                && ReferenceEquals(LocalSheet, other.LocalSheet);
        }

        /// <summary>
        /// Determines whether the specified object is equal to the current instance.
        /// </summary>
        /// <param name="obj">Other object, or null.</param>
        /// <returns>True if the object is a <see cref="DefinedName"/> equal to this instance.</returns>
        public override bool Equals(object obj)
        {
            return Equals(obj as DefinedName);
        }

        /// <summary>
        /// Returns a hash code consistent with <see cref="Equals(DefinedName)"/>.
        /// </summary>
        /// <returns>Hash code derived from the public members of this instance.</returns>
        public override int GetHashCode()
        {
            unchecked
            {
                int hash = 17;
                hash = (hash * 31) + (Name != null ? Name.GetHashCode() : 0);
                hash = (hash * 31) + Type.GetHashCode();
                hash = (hash * 31) + (TextValue != null ? TextValue.GetHashCode() : 0); // Object implicit covered by string
                hash = (hash * 31) + (Comment != null ? Comment.GetHashCode() : 0);
                hash = (hash * 31) + (TargetWorksheet != null ? System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(TargetWorksheet) : 0);
                hash = (hash * 31) + (LocalSheet != null ? System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(LocalSheet) : 0);
                return hash;
            }
        }

        /// <summary>
        /// Compares this instance with another <see cref="DefinedName"/> for ordering. The order is
        /// determined by <see cref="Name"/> (ordinal), then by scope (workbook scope sorts before any
        /// worksheet-scoped name; for worksheet-scoped names by <see cref="Worksheet.SheetID"/>), then
        /// by <see cref="TextValue"/>, then by <see cref="Comment"/>.
        /// </summary>
        /// <param name="other">Other defined name, or null. A null comparand sorts after this instance.</param>
        /// <returns>Negative, zero, or positive integer following the standard <see cref="IComparable{T}"/> contract.</returns>
        public int CompareTo(DefinedName other)
        {
            if (other is null)
            {
                return 1;
            }
            int cmp = string.CompareOrdinal(Name, other.Name);
            if (cmp != 0)
            {
                return cmp;
            }
            cmp = Type.CompareTo(other.Type);
            if (cmp != 0)
            {
                return cmp;
            }
            cmp = CompareScope(LocalSheet, other.LocalSheet);
            if (cmp != 0)
            {
                return cmp;
            }
            cmp = CompareScope(TargetWorksheet, other.TargetWorksheet);
            if (cmp != 0)
            {
                return cmp;
            }
            cmp = string.CompareOrdinal(TextValue, other.TextValue);
            if (cmp != 0)
            {
                return cmp;
            }
            return string.CompareOrdinal(Comment, other.Comment);
        }

        /// <summary>
        /// Returns a textual representation of the defined name (intended for debugging).
        /// </summary>
        /// <returns>A short string with name, scope and reference.</returns>
        public override string ToString()
        {
            string scope = LocalSheet == null ? "workbook" : "sheet:" + LocalSheet.SheetName;
            return "DefinedName{name=" + Name + ", scope=" + scope + ", ref=" + TextValue + "}";
        }

        /// <summary>
        /// Compares two scope worksheets for ordering. Workbook scope (null) sorts before any worksheet
        /// scope; two worksheet scopes are ordered by their <see cref="Worksheet.SheetID"/>.
        /// </summary>
        /// <param name="left">Left scope, or null for workbook scope.</param>
        /// <param name="right">Right scope, or null for workbook scope.</param>
        /// <returns>Negative, zero, or positive integer following the standard ordering contract.</returns>
        private static int CompareScope(Worksheet left, Worksheet right)
        {
            if (ReferenceEquals(left, right))
            {
                return 0;
            }
            if (left == null)
            {
                return -1;
            }
            if (right == null)
            {
                return 1;
            }
            return left.SheetID.CompareTo(right.SheetID);
        }
        #endregion
    }
}
