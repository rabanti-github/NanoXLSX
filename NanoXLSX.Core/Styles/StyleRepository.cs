/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace NanoXLSX.Styles
{
    /// <summary>
    /// Class to manage all styles at runtime, before writing XLSX files. The main purpose is deduplication and decoupling of styles from workbooks at runtime
    /// </summary>
    /// \remark <remarks>Be careful when changing style data in this class. It may lead to inconsistencies</remarks>
    public class StyleRepository
    {
        private readonly object lockObject = new object();
        private readonly Dictionary<int, Style> styles;
        private static readonly StyleRepository instance = new StyleRepository();

        /// <summary>
        /// Gets the singleton instance of the repository
        /// </summary>
        public static StyleRepository Instance
        {
            get
            {
                return instance;
            }
        }

        /// <summary>
        /// If true certain exceptions will be suppressed and transformations on styles are performed when a worksheet is loaded
        /// </summary>
        internal bool ImportInProgress { get; set; }

        /// <summary>
        /// Gets a snapshot of the currently managed styles of the repository
        /// </summary>
        /// <deprecated>Please use <see cref="ManagedStyles"/> instead</deprecated>
        [Obsolete("Will be removed in the next major version. Use ManagedStyles instead.")]
        public Dictionary<int, Style> Styles
        {
            get
            {
                lock (lockObject)
                {
                    return new Dictionary<int, Style>(styles);
                }
            }
        }

        /// <summary>
        /// Gets a read-only snapshot of the currently managed styles of the repository
        /// </summary>
        public IReadOnlyDictionary<int, Style> ManagedStyles
        {
            get
            {
                lock (lockObject)
                {
                    return new ReadOnlyDictionary<int, Style>(new Dictionary<int, Style>(styles));
                }
            }
        }

        /// <summary>
        /// Private constructor. The class is not intended to instantiate outside the singleton
        /// </summary>
        private StyleRepository()
        {
            styles = new Dictionary<int, Style>();
        }

        /// <summary>
        /// Adds a style to the repository and returns the actual reference
        /// </summary>
        /// <param name="style">Style to add</param>
        /// <returns>Reference from the repository. If the style to add already existed, the existing object is returned, otherwise the newly added one</returns>
        public Style AddStyle(Style style)
        {
            lock (lockObject)
            {
                if (style == null)
                {
                    return null;
                }
                int hashCode = style.GetHashCode();
                if (!styles.TryGetValue(hashCode, out var value))
                {
                    value = style;
                    styles.Add(hashCode, value);
                }
                return value;
            }
        }

        /// <summary>
        /// Empties the static repository
        /// </summary>
        /// \remark <remarks>Do not use this maintenance method while writing data on a worksheet or workbook. It will lead to invalid style data or even exceptions.<br />
        /// Only use this method after all worksheets in all workbooks are disposed.It may free memory then.</remarks>
        public void FlushStyles()
        {
            lock (lockObject)
            {
                styles.Clear();
            }
        }

    }
}
