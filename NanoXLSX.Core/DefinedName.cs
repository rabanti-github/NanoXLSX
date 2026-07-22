/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Text.RegularExpressions;
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
    /// The <see cref="Reference"/> is stored verbatim — NanoXLSX does not parse or evaluate it.
    /// </remarks>
    public sealed class DefinedName : IEquatable<DefinedName>, IComparable<DefinedName>
    {
        /// <summary>
        /// Regular expression matching a cell reference in the range A1 - XFD1048576.
        /// </summary>
        /// \remark <remarks>According to the OOXML specification (§18.2.5), a defined name in this range is considered an error.</remarks>
        private static readonly Regex CellReferenceRegex = new Regex(
            "^[A-Za-z]{1,3}[0-9]+$", RegexOptions.Compiled);

        #region properties
        /// <summary>
        /// Gets the name of the defined name as it appears in the workbook (e.g. <c>MyRange</c>).
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Gets the textual reference of the defined name. This is stored verbatim and may be a cell address
        /// (e.g. <c>Sheet1!$A$1</c>), a range (e.g. <c>Sheet1!$A$1:$A$10</c>), a formula (e.g. <c>SUM(Sheet1!$A$1:$A$10)</c>),
        /// or a constant value.
        /// </summary>
        public string Reference { get; }

        /// <summary>
        /// Gets the worksheet that scopes this defined name. If null, the defined name has workbook scope and
        /// is visible from any worksheet. If non-null, the defined name is scoped to the referenced worksheet
        /// (mapped to the <c>localSheetId</c> attribute on save).
        /// </summary>
        public Worksheet LocalSheet { get; }

        /// <summary>
        /// Gets the optional comment associated with the defined name. Maps to the <c>comment</c> attribute in OOXML.
        /// May be null when no comment was set.
        /// </summary>
        public string Comment { get; }
        #endregion

        #region constructors
        /// <summary>
        /// Constructs a new defined name.
        /// </summary>
        /// <param name="name">Name of the defined name. Must be non-empty, must not start with a digit, and must not match a cell reference in the range A1 - XFD1048576.</param>
        /// <param name="reference">Reference text (cell, range, formula, or constant). Must be non-empty.</param>
        /// <param name="localSheet">Optional worksheet that scopes the defined name. Pass null for workbook scope.</param>
        /// <param name="comment">Optional comment.</param>
        /// <exception cref="FormatException">Thrown if <paramref name="name"/> is null, empty, starts with a digit, or matches a cell reference.</exception>
        /// <exception cref="FormatException">Thrown if <paramref name="reference"/> is null or empty.</exception>
        public DefinedName(string name, string reference, Worksheet localSheet = null, string comment = null)
        {
            ValidateName(name);
            if (string.IsNullOrEmpty(reference))
            {
                throw new FormatException("The reference of a defined name must not be null or empty.");
            }
            this.Name = name;
            this.Reference = reference;
            this.LocalSheet = localSheet;
            this.Comment = comment;
        }
        #endregion

        #region methods
        /// <summary>
        /// Determines whether the specified <see cref="DefinedName"/> instance is equal to the current instance.
        /// Two instances are considered equal when their <see cref="Name"/>, <see cref="Reference"/>,
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
                && string.Equals(Reference, other.Reference, StringComparison.Ordinal)
                && string.Equals(Comment, other.Comment, StringComparison.Ordinal)
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
                hash = (hash * 31) + (Reference != null ? Reference.GetHashCode() : 0);
                hash = (hash * 31) + (Comment != null ? Comment.GetHashCode() : 0);
                hash = (hash * 31) + (LocalSheet != null ? System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(LocalSheet) : 0);
                return hash;
            }
        }

        /// <summary>
        /// Compares this instance with another <see cref="DefinedName"/> for ordering. The order is
        /// determined by <see cref="Name"/> (ordinal), then by scope (workbook scope sorts before any
        /// worksheet-scoped name; for worksheet-scoped names by <see cref="Worksheet.SheetID"/>), then
        /// by <see cref="Reference"/>, then by <see cref="Comment"/>.
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
            cmp = CompareScope(LocalSheet, other.LocalSheet);
            if (cmp != 0)
            {
                return cmp;
            }
            cmp = string.CompareOrdinal(Reference, other.Reference);
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
            return "DefinedName{name=" + Name + ", scope=" + scope + ", ref=" + Reference + "}";
        }

        /// <summary>
        /// Validates that the supplied name is non-empty, does not start with a digit, and does not match
        /// a cell reference in the range A1 - XFD1048576.
        /// </summary>
        /// <param name="name">Name to validate.</param>
        /// <exception cref="FormatException">Thrown if validation fails.</exception>
        private static void ValidateName(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                throw new FormatException("The name of a defined name must not be null or empty.");
            }
            char first = name[0];
            if (first >= '0' && first <= '9')
            {
                throw new FormatException("The name of a defined name must not start with a digit. Provided: '" + name + "'");
            }
            if (CellReferenceRegex.IsMatch(name) && IsCellReferenceInRange(name))
            {
                throw new FormatException("The name of a defined name must not match a cell reference in the range A1 - XFD1048576. Provided: '" + name + "'");
            }
        }

        /// <summary>
        /// Determines whether a candidate name parses as a cell reference within the legal range A1 - XFD1048576.
        /// </summary>
        /// <param name="name">Candidate name (already known to match <see cref="CellReferenceRegex"/>).</param>
        /// <returns>True if the candidate is a cell reference in range, otherwise false.</returns>
        private static bool IsCellReferenceInRange(string name)
        {
            int letterCount = 0;
            while (letterCount < name.Length && IsLetter(name[letterCount]))
            {
                letterCount++;
            }
            string letters = name.Substring(0, letterCount).ToUpperInvariant();
            string digits = name.Substring(letterCount);
            int column = 0;
            foreach (char c in letters)
            {
                column = (column * 26) + (c - 'A' + 1);
            }
            // A=1 ... XFD=16384
            if (column < 1 || column > 16384)
            {
                return false;
            }
            if (!int.TryParse(digits, out int row))
            {
                return false;
            }
            return row >= 1 && row <= 1048576;
        }

        /// <summary>
        /// Indicates whether a character is an ASCII letter (A-Z or a-z).
        /// </summary>
        /// <param name="c">Character to test.</param>
        /// <returns>True if the character is a letter, otherwise false.</returns>
        private static bool IsLetter(char c)
        {
            return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z');
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
