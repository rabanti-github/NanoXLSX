/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2018
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Globalization;
using FormatException = NanoXLSX.Exception.FormatException;

namespace NanoXLSX
{
    public partial class Cell
    {
        /// <summary>
        /// Class for handling of basic Excel formulas
        /// </summary>
        public static class BasicFormulas
        {
            /// <summary>
            /// Returns a cell with a average formula
            /// </summary>
            /// <param name="range">Cell range to apply the average operation to</param>
            /// <returns>Prepared Cell object, ready to be added to a worksheet</returns>
            public static Cell Average(Range range)
            { return Average(null, range); }

            /// <summary>
            /// Returns a cell with a average formula
            /// </summary>
            /// <param name="target">Target worksheet of the average operation. Can be null if on the same worksheet</param>
            /// <param name="range">Cell range to apply the average operation to</param>
            /// <returns>Prepared Cell object, ready to be added to a worksheet</returns>
            public static Cell Average(Worksheet target, Range range)
            { return GetBasicFormula(target, range, "AVERAGE", null); }

            /// <summary>
            /// Returns a cell with a ceil formula
            /// </summary>
            /// <param name="address">Address to apply the ceil operation to</param>
            /// <param name="decimals">Number of decimals (digits)</param>
            /// <returns>Prepared Cell object, ready to be added to a worksheet</returns>
            public static Cell Ceil(Address address, int decimals)
            { return Ceil(null, address, decimals); }

            /// <summary>
            /// Returns a cell with a ceil formula
            /// </summary>
            /// <param name="target">Target worksheet of the ceil operation. Can be null if on the same worksheet</param>
            /// <param name="address">Address to apply the ceil operation to</param>
            /// <param name="decimals">Number of decimals (digits)</param>
            /// <returns>Prepared Cell object, ready to be added to a worksheet</returns>
            public static Cell Ceil(Worksheet target, Address address, int decimals)
            { return GetBasicFormula(target, new Range(address, address), "ROUNDUP", decimals.ToString()); }

            /// <summary>
            /// Returns a cell with a floor formula
            /// </summary>
            /// <param name="address">Address to apply the floor operation to</param>
            /// <param name="decimals">Number of decimals (digits)</param>
            /// <returns>Prepared Cell object, ready to be added to a worksheet</returns>
            public static Cell Floor(Address address, int decimals)
            { return Floor(null, address, decimals); }

            /// <summary>
            /// Returns a cell with a floor formula
            /// </summary>
            /// <param name="target">Target worksheet of the floor operation. Can be null if on the same worksheet</param>
            /// <param name="address">Address to apply the floor operation to</param>
            /// <param name="decimals">Number of decimals (digits)</param>
            /// <returns>Prepared Cell object, ready to be added to a worksheet</returns>
            public static Cell Floor(Worksheet target, Address address, int decimals)
            { return GetBasicFormula(target, new Range(address, address), "ROUNDDOWN", decimals.ToString()); }

            /// <summary>
            /// Returns a cell with a max formula
            /// </summary>
            /// <param name="range">Cell range to apply the max operation to</param>
            /// <returns>Prepared Cell object, ready to be added to a worksheet</returns>
            public static Cell Max(Range range)
            { return Max(null, range); }

            /// <summary>
            /// Returns a cell with a max formula
            /// </summary>
            /// <param name="target">Target worksheet of the max operation. Can be null if on the same worksheet</param>
            /// <param name="range">Cell range to apply the max operation to</param>
            /// <returns>Prepared Cell object, ready to be added to a worksheet</returns>
            public static Cell Max(Worksheet target, Range range)
            { return GetBasicFormula(target, range, "MAX", null); }

            /// <summary>
            /// Returns a cell with a median formula
            /// </summary>
            /// <param name="range">Cell range to apply the median operation to</param>
            /// <returns>Prepared Cell object, ready to be added to a worksheet</returns>
            public static Cell Median(Range range)
            { return Median(null, range); }

            /// <summary>
            /// Returns a cell with a median formula
            /// </summary>
            /// <param name="target">Target worksheet of the median operation. Can be null if on the same worksheet</param>
            /// <param name="range">Cell range to apply the median operation to</param>
            /// <returns>Prepared Cell object, ready to be added to a worksheet</returns>
            public static Cell Median(Worksheet target, Range range)
            { return GetBasicFormula(target, range, "MEDIAN", null); }

            /// <summary>
            /// Returns a cell with a min formula
            /// </summary>
            /// <param name="range">Cell range to apply the min operation to</param>
            /// <returns>Prepared Cell object, ready to be added to a worksheet</returns>
            public static Cell Min(Range range)
            { return Min(null, range); }

            /// <summary>
            /// Returns a cell with a min formula
            /// </summary>
            /// <param name="target">Target worksheet of the min operation. Can be null if on the same worksheet</param>
            /// <param name="range">Cell range to apply the median operation to</param>
            /// <returns>Prepared Cell object, ready to be added to a worksheet</returns>
            public static Cell Min(Worksheet target, Range range)
            { return GetBasicFormula(target, range, "MIN", null); }

            /// <summary>
            /// Returns a cell with a round formula
            /// </summary>
            /// <param name="address">Address to apply the round operation to</param>
            /// <param name="decimals">Number of decimals (digits)</param>
            /// <returns>Prepared Cell object, ready to be added to a worksheet</returns>
            public static Cell Round(Address address, int decimals)
            { return Round(null, address, decimals); }

            /// <summary>
            /// Returns a cell with a round formula
            /// </summary>
            /// <param name="target">Target worksheet of the round operation. Can be null if on the same worksheet</param>
            /// <param name="address">Address to apply the round operation to</param>
            /// <param name="decimals">Number of decimals (digits)</param>
            /// <returns>Prepared Cell object, ready to be added to a worksheet</returns>
            public static Cell Round(Worksheet target, Address address, int decimals)
            { return GetBasicFormula(target, new Range(address, address), "ROUND", decimals.ToString()); }

            /// <summary>
            /// Returns a cell with a sum formula
            /// </summary>
            /// <param name="range">Cell range to get a sum of</param>
            /// <returns>Prepared Cell object, ready to be added to a worksheet</returns>
            public static Cell Sum(Range range)
            { return Sum(null, range); }

            /// <summary>
            /// Returns a cell with a sum formula
            /// </summary>
            /// <param name="target">Target worksheet of the sum operation. Can be null if on the same worksheet</param>
            /// <param name="range">Cell range to get a sum of</param>
            /// <returns>Prepared Cell object, ready to be added to a worksheet</returns>
            public static Cell Sum(Worksheet target, Range range)
            { return GetBasicFormula(target, range, "SUM", null); }


            /// <summary>
            /// Function to generate a Vlookup as Excel function
            /// </summary>
            /// <param name="number">Numeric value for the lookup. Valid types are int, long, float and double</param>
            /// <param name="range">Matrix of the lookup</param>
            /// <param name="columnIndex">Column index of the target column (1 based)</param>
            /// <param name="exactMatch">If true, an exact match is applied to the lookup</param>
            /// <returns>Prepared Cell object, ready to be added to a worksheet</returns>
            public static Cell VLookup(object number, Range range, int columnIndex, bool exactMatch)
            { return VLookup(number, null, range, columnIndex, exactMatch); }

            /// <summary>
            /// Function to generate a Vlookup as Excel function
            /// </summary>
            /// <param name="number">Numeric value for the lookup. Valid types are int, long, float and double</param>
            /// <param name="rangeTarget">Target worksheet of the matrix. Can be null if on the same worksheet</param>
            /// <param name="range">Matrix of the lookup</param>
            /// <param name="columnIndex">Column index of the target column (1 based)</param>
            /// <param name="exactMatch">If true, an exact match is applied to the lookup</param>
            /// <returns>Prepared Cell object, ready to be added to a worksheet</returns>
            public static Cell VLookup(object number, Worksheet rangeTarget, Range range, int columnIndex, bool exactMatch)
            { return GetVLookup(null, new Address(), number, rangeTarget, range, columnIndex, exactMatch, true); }

            /// <summary>
            /// Function to generate a Vlookup as Excel function
            /// </summary>
            /// <param name="address">Query address of a cell as string as source of the lookup</param>
            /// <param name="range">Matrix of the lookup</param>
            /// <param name="columnIndex">Column index of the target column (1 based)</param>
            /// <param name="exactMatch">If true, an exact match is applied to the lookup</param>
            /// <returns>Prepared Cell object, ready to be added to a worksheet</returns>
            public static Cell VLookup(Address address, Range range, int columnIndex, bool exactMatch)
            { return VLookup(null, address, null, range, columnIndex, exactMatch); }

            /// <summary>
            /// Function to generate a Vlookup as Excel function
            /// </summary>
            /// <param name="queryTarget">Target worksheet of the query argument. Can be null if on the same worksheet</param>
            /// <param name="address">Query address of a cell as string as source of the lookup</param>
            /// <param name="rangeTarget">Target worksheet of the matrix. Can be null if on the same worksheet</param>
            /// <param name="range">Matrix of the lookup</param>
            /// <param name="columnIndex">Column index of the target column (1 based)</param>
            /// <param name="exactMatch">If true, an exact match is applied to the lookup</param>
            /// <returns>Prepared Cell object, ready to be added to a worksheet</returns>
            public static Cell VLookup(Worksheet queryTarget, Address address, Worksheet rangeTarget, Range range, int columnIndex, bool exactMatch)
            {
                return GetVLookup(queryTarget, address, 0, rangeTarget, range, columnIndex, exactMatch, false);
            }

            /// <summary>
            /// Function to generate a Vlookup as Excel function
            /// </summary>
            /// <param name="queryTarget">Target worksheet of the query argument. Can be null if on the same worksheet</param>
            /// <param name="address">In case of a reference lookup, query address of a cell as string</param>
            /// <param name="number">In case of a numeric lookup, number for the lookup</param>
            /// <param name="rangeTarget">Target worksheet of the matrix. Can be null if on the same worksheet</param>
            /// <param name="range">Matrix of the lookup</param>
            /// <param name="columnIndex">Column index of the target column (1 based)</param>
            /// <param name="exactMatch">If true, an exact match is applied to the lookup</param>
            /// <param name="numericLookup">If true, the lookup is a numeric lookup, otherwise a reference lookup</param>
            /// <returns>Prepared Cell object, ready to be added to a worksheet</returns>
            private static Cell GetVLookup(Worksheet queryTarget, Address address, object number, Worksheet rangeTarget, Range range, int columnIndex, bool exactMatch, bool numericLookup)
            {
                CultureInfo culture = CultureInfo.InvariantCulture;
                string arg1, arg2, arg3, arg4;
                if (numericLookup == true)
                {
                    Type t = number.GetType();
                    if (t == typeof(byte))          { arg1 = ((byte)number).ToString("G", culture); }
                    else if (t == typeof(sbyte))    { arg1 = ((sbyte)number).ToString("G", culture); }
                    else if (t == typeof(decimal))  { arg1 = ((decimal)number).ToString("G", culture); }
                    else if (t == typeof(double))   { arg1 = ((double)number).ToString("G", culture); }
                    else if (t == typeof(float))    { arg1 = ((float)number).ToString("G", culture); }
                    else if (t == typeof(int))      { arg1 = ((int)number).ToString("G", culture); }
                    else if (t == typeof(long))     { arg1 = ((long)number).ToString("G", culture); }
                    else if (t == typeof(ulong))    { arg1 = ((ulong)number).ToString("G", culture); }
                    else if (t == typeof(short))    { arg1 = ((short)number).ToString("G", culture); }
                    else if (t == typeof(ushort))   { arg1 = ((ushort)number).ToString("G", culture); }
                    else
                    {
                        throw new FormatException("InvalidLookupType", "The lookup variable can only be a cell address or a numeric value. The value '" + number + "' is invalid.");
                    }
                }
                else
                {
                    if (queryTarget != null) { arg1 = queryTarget.SheetName + "!" + address.ToString(); }
                    else { arg1 = address.ToString(); }
                }
                if (rangeTarget != null) { arg2 = rangeTarget.SheetName + "!" + range.ToString(); }
                else { arg2 = range.ToString(); }
                arg3 = columnIndex.ToString("G", culture);
                if (exactMatch == true) { arg4 = "TRUE"; }
                else { arg4 = "FALSE"; }
                return new Cell("VLOOKUP(" + arg1 + "," + arg2 + "," + arg3 + "," + arg4 + ")", CellType.FORMULA);
            }


            /// <summary>
            /// Function to generate a basic Excel function with one cell range as parameter and an optional post argument
            /// </summary>
            /// <param name="target">Target worksheet of the cell reference. Can be null if on the same worksheet</param>
            /// <param name="range">Main argument as cell range. If applied on one cell, the start and end address are identical</param>
            /// <param name="functionName">Internal Excel function name</param>
            /// <param name="postArg">Optional argument</param>
            /// <returns>Prepared Cell object, ready to be added to a worksheet</returns>
            private static Cell GetBasicFormula(Worksheet target, Range range, string functionName, string postArg)
            {
                string arg1, arg2, prefix;
                if (postArg == null) { arg2 = ""; }
                else { arg2 = "," + postArg; }
                if (target != null) { prefix = target.SheetName + "!"; }
                else { prefix = ""; }
                if (range.StartAddress.Equals(range.EndAddress)) { arg1 = prefix + range.StartAddress.ToString(); }
                else { arg1 = prefix + range.ToString(); }
                return new Cell(functionName + "(" + arg1 + arg2 + ")", CellType.FORMULA);
            }
        }
    }
}