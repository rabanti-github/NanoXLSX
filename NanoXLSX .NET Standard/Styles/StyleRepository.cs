using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2022
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

namespace NanoXLSX.Styles
{
    /// <summary>
    /// Class to manage all styles at runtime, before writing XLSX files. The main purpose is deduplication and decouplig of styles from workbooks at runtime
    /// </summary>
    public class StyleRepository
    {
        private static StyleRepository instance;

        /// <summary>
        /// Singleton instance of the repository
        /// </summary>
        public static StyleRepository Instance
        {
            get
            {
                instance = instance ?? new StyleRepository();
                return instance;
            }
        }

        private Dictionary<int, Style> styles;

        /// <summary>
        /// Gets the currently managed styles of the repository
        /// </summary>
        public Dictionary<int, Style> Styles { get => styles; }

        /// <summary>
        /// Private constructor. The class is not intended to instantiate outside of the singleton
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
            if (style == null)
            {
                return null;
            }
            int hashCode = style.GetHashCode();
            if (!styles.ContainsKey(hashCode))
            {
                styles.Add(hashCode, style);
            }
            return styles[hashCode];
        }

        /// <summary>
        /// Empties the static repository
        /// </summary>
        /// <remarks>Do not use this maintenance method while writing data on a worksheet or workbook. It will lead to invalid style data or even exceptions.<br/>
        /// Only use this method after all worksheets in all workbooks are disposed. It may free memory then.
        /// </remarks>
        public void FlushStyles()
        {
            styles.Clear();
        }


    }
}
