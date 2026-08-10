/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using NanoXLSX.Enums;
using NanoXLSX.Internal;
using NanoXLSX.Utils;

namespace NanoXLSX
{
    /// <summary>
    /// Class representing a formula in a cell, its data, respectively
    /// </summary>
    public class FormulaData : IEquatable<FormulaData>, IComparable<FormulaData>
    {
        private string expression;

        #region enums
        /// <summary>
        /// Enum to define the specific type of a formal if the Cell has the type <see cref="Cell.CellType.Formula"/>
        /// </summary>
        public enum FormulaType
        {
            /// <summary>
            /// Cell contains a regular formula (e.g "A1+A2")
            /// </summary>
            Normal,
            /// <summary>
            /// Cell contains a formula that is part of an array
            /// </summary>
            Array,
            /// <summary>
            /// Cell contains a shared formula, pointing to another formula that is identical
            /// </summary>
            Shared,
            /// <summary>
            /// Cell contains a formula that is applied across a range of one or more cells
            /// </summary>
            DataTable
        }
        #endregion

        private bool hasExternalReferences;
        private DefinedName definedNameReference;

        #region properties

        /// <summary>
        /// Gets the formula expression as string. This value is currently identical with <see cref="Cell.Value"/> if <see cref="Cell.DataType"/> is set to <see cref="Cell.CellType.Formula"/>.
        /// </summary>
        public string Expression
        {
            get { return expression; }
            internal set
            {
                expression = value;
                HasExternalReferences = ParserUtils.ContainsExternalReference(value);
            }
        }

        /// <summary>
        /// Gets whether the formula expression contains a reference to an external workbook.
        /// </summary>
        public bool HasExternalReferences
        {
            get { return hasExternalReferences; }
            internal set
            {
                hasExternalReferences = value;
                Features.SetFormulaFeatures(definedNameReference == null ? false : true, hasExternalReferences);
            }
        }

        /// <summary>
        /// Type of the formula. Default is <see cref="FormulaType.Normal"/>
        /// </summary>
        public FormulaType Type { get; internal set; }

        /// <summary>
        /// Gets the range associated with an array, shared, or data-table formula. Can be a range or address (string representation). Default is null, if no reference was defined.
        /// </summary>
        public string FormulaRange { get; internal set; }

        /// <summary>
        /// Resolved defined name, if the complete formula expression is a direct reference to exactly one defined name.
        /// </summary>
        /// \remark <remarks>Defined names within a formula expression are currently not resolved and will not be set in this property, even if only one defined name occurs in the expression.</remarks>
        public DefinedName DefinedNameReference
        {
            get { return definedNameReference; }
            internal set
            {
                definedNameReference = value;
                Features.SetFormulaFeatures(definedNameReference == null ? false : true, hasExternalReferences);
            }
        }
        /// <summary>
        /// Gets the cached value of the formula
        /// </summary>
        /// \remark <remarks>This value can be supplied through the constructor or set when a Workbook is loaded. It is not evaluated when a new formula was defined by <see cref="Worksheet.AddCellFormula(string, int, int)"/> or its overload methods</remarks>
        public object CachedValue { get; internal set; }

        /// <summary>
        /// Gets the data type of <see cref="CachedValue"/>. The default is <see cref="Cell.CellType.Default"/>
        /// if no cached value or no supported cached value type is available.
        /// </summary>
        public Cell.CellType CachedValueType { get; internal set; }

        /// <summary>
        /// Gets the address of the formula's master cell. This is mainly used in case of <see cref="FormulaType.Array"/>.
        /// </summary>
        public string MasterCellAddress { get; internal set; }

        /// <summary>
        /// Internal feature set for cascading feature detection (consider in  <see cref="Copy"/> but not in Equals, GetHashCode etc.)
        /// </summary>
        internal FeatureSet Features { get; private set; } = FeatureSet.CreateFormula();

        #endregion

        #region constructors

        /// <summary>
        /// Default constructor
        /// </summary>
        public FormulaData()
        {
            this.Type = FormulaType.Normal;
            this.CachedValueType = Cell.CellType.Default;
        }

        /// <summary>
        /// Constructor with formula expression and optional cached value to create a formula of the common type <see cref="FormulaType.Normal"/>
        /// </summary>
        /// <param name="expression">Formula expression (without leading equal sign)</param>
        /// <param name="cachedValue">Optional cached value. Default is null</param>
        /// \remark <remarks>A basic validity checks (not full parsing) will perform on the expression, e.g. existence of an external link in the formula</remarks>
        public FormulaData(string expression, object cachedValue = null) : this()
        {
            Expression = expression;
            CachedValue = cachedValue;
            CachedValueType = ResolveCachedValueType(cachedValue);
        }

        #endregion
        #region methods

        /// <summary>
        /// Resolves the cell type of a cached formula value without evaluating the formula.
        /// </summary>
        /// <param name="cachedValue">Cached formula value, or null if unavailable.</param>
        /// <returns>Resolved cached value type.</returns>
        internal static Cell.CellType ResolveCachedValueType(object cachedValue)
        {
            if (cachedValue == null)
            {
                return Cell.CellType.Default;
            }
            if (cachedValue is bool)
            {
                return Cell.CellType.Bool;
            }
            if (cachedValue is byte || cachedValue is sbyte || cachedValue is decimal || cachedValue is double
                || cachedValue is float || cachedValue is int || cachedValue is uint || cachedValue is long
                || cachedValue is ulong || cachedValue is short || cachedValue is ushort)
            {
                return Cell.CellType.Number;
            }
            if (cachedValue is DateTime)
            {
                return Cell.CellType.Date;
            }
            if (cachedValue is TimeSpan)
            {
                return Cell.CellType.Time;
            }
            if (cachedValue is Errors.FormulaError)
            {
                return Cell.CellType.Error;
            }
            return Cell.CellType.String;
        }

        /// <summary>
        /// Copies the current object into a new one (without copying <see cref="DefinedNameReference"/>)
        /// </summary>
        /// <returns>Copy of the current instance</returns>
        /// \remark <remarks>This copy method omits deep-copying <see cref="DefinedNameReference"/> by design. If a full copy is intended, this instance variable must be handled separately.</remarks>
        internal FormulaData Copy()
        {
            FormulaData data = new FormulaData();
            data.Expression = this.Expression;
            data.Type = this.Type;
            data.FormulaRange = this.FormulaRange;
            data.CachedValue = this.CachedValue;
            data.CachedValueType = this.CachedValueType;
            data.MasterCellAddress = this.MasterCellAddress;
            data.Features = this.Features.Copy(); // New feature set
            data.DefinedNameReference = this.DefinedNameReference; // object reference
            return data;
        }

        /// <summary>
        /// Compares this instance with another <see cref="FormulaData"/> instance.
        /// </summary>
        /// <param name="other">Other formula data instance, or null.</param>
        /// <returns>Negative, zero, or positive integer following the standard comparison contract.</returns>
        public int CompareTo(FormulaData other)
        {
            if (other is null)
            {
                return 1;
            }
            int cmp = string.CompareOrdinal(Expression, other.Expression);
            if (cmp != 0)
            {
                return cmp;
            }
            cmp = Type.CompareTo(other.Type);
            if (cmp != 0)
            {
                return cmp;
            }
            cmp = string.CompareOrdinal(FormulaRange, other.FormulaRange);
            if (cmp != 0)
            {
                return cmp;
            }
            cmp = Comparer<DefinedName>.Default.Compare(DefinedNameReference, other.DefinedNameReference);
            if (cmp != 0)
            {
                return cmp;
            }
            cmp = CachedValueType.CompareTo(other.CachedValueType);
            if (cmp != 0)
            {
                return cmp;
            }
            cmp = Comparer<object>.Default.Compare(CachedValue, other.CachedValue);
            if (cmp != 0)
            {
                return cmp;
            }
            return string.CompareOrdinal(MasterCellAddress, other.MasterCellAddress);
        }

        /// <summary>
        /// Determines whether the specified <see cref="FormulaData"/> instance is equal to this instance.
        /// </summary>
        /// <param name="other">Other formula data instance, or null.</param>
        /// <returns>True if equal, otherwise false.</returns>
        public bool Equals(FormulaData other)
        {
            if (other is null)
            {
                return false;
            }
            if (ReferenceEquals(this, other))
            {
                return true;
            }
            return string.Equals(Expression, other.Expression, StringComparison.Ordinal)
                && Type == other.Type
                && string.Equals(FormulaRange, other.FormulaRange, StringComparison.Ordinal)
                && EqualityComparer<DefinedName>.Default.Equals(DefinedNameReference, other.DefinedNameReference)
                && EqualityComparer<object>.Default.Equals(CachedValue, other.CachedValue)
                && CachedValueType == other.CachedValueType
                && string.Equals(MasterCellAddress, other.MasterCellAddress, StringComparison.Ordinal);
        }

        /// <summary>
        /// Determines whether the specified object is equal to this instance.
        /// </summary>
        /// <param name="obj">Other object, or null.</param>
        /// <returns>True if equal, otherwise false.</returns>
        public override bool Equals(object obj)
        {
            return Equals(obj as FormulaData);
        }

        /// <summary>
        /// Returns a hash code consistent with <see cref="Equals(FormulaData)"/>.
        /// </summary>
        /// <returns>Hash code derived from this instance's properties.</returns>
        public override int GetHashCode()
        {
            unchecked
            {
                int hash = 17;
                hash = (hash * 31) + (Expression != null ? Expression.GetHashCode() : 0);
                hash = (hash * 31) + Type.GetHashCode();
                hash = (hash * 31) + (FormulaRange != null ? FormulaRange.GetHashCode() : 0);
                hash = (hash * 31) + (DefinedNameReference != null ? DefinedNameReference.GetHashCode() : 0);
                hash = (hash * 31) + (CachedValue != null ? CachedValue.GetHashCode() : 0);
                hash = (hash * 31) + CachedValueType.GetHashCode();
                hash = (hash * 31) + (MasterCellAddress != null ? MasterCellAddress.GetHashCode() : 0);
                return hash;
            }
        }
        #endregion
    }
}
