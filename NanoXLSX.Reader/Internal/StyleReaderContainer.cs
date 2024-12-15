/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2024
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using NanoXLSX.Shared.Exceptions;
using NanoXLSX.Styles;

namespace NanoXLSX.Internal
{
    /// <summary>
    /// Class representing a collection of pre-processed styles and their components. This class is internally used and should not be used otherwise.
    /// </summary>
    public class StyleReaderContainer
    {

        #region privateFields

        private List<CellXf> cellXfs = new List<CellXf>();
        private List<NumberFormat> numberFormats = new List<NumberFormat>();
        private List<Style> styles = new List<Style>();
        private List<Border> borders = new List<Border>();
        private List<Fill> fills = new List<Fill>();
        private List<Font> fonts = new List<Font>();
        private List<string> mruColors = new List<string>();

        #endregion

        #region properties

        /// <summary>
        /// Gets the number of resolved styles
        /// </summary>
        public int StyleCount
        {
            get { return styles.Count; }
        }

        #endregion

        #region functions

        /// <summary>
        /// Adds a style component and determines the appropriate type of it automatically
        /// </summary>
        /// <param name="component">Style component to add to the collections</param>

        public void AddStyleComponent(AbstractStyle component)
        {
            Type t = component.GetType();
            if (t == typeof(CellXf))
            {
                cellXfs.Add(component as CellXf);
            }
            else if (t == typeof(NumberFormat))
            {
                numberFormats.Add(component as NumberFormat);
            }
            else if (t == typeof(Style))
            {
                styles.Add(component as Style);
            }
            else if (t == typeof(Border))
            {
                borders.Add(component as Border);
            }
            else if (t == typeof(Fill))
            {
                fills.Add(component as Fill);
            }
            else if (t == typeof(Font))
            {
                fonts.Add(component as Font);
            }
        }

        /// <summary>
        /// Returns a whole style by its index
        /// </summary>
        /// <param name="index">Index of the style</param>
        /// <returns>Style object or null if the component could not be retrieved</returns>
        public Style GetStyle(string index)
        {
            int number;
            if (int.TryParse(index, out number))
            {
                return GetComponent(typeof(Style), number) as Style;
            }
            return null;
        }

        /// <summary>
        /// Returns a whole style by its index. It also returns information about the type of the style
        /// </summary>
        /// <param name="index">Index of the style</param>
        /// <param name="isDateStyle">Out parameter that indicates whether the style represents a date</param>
        /// <param name="isTimeStyle">Out parameter that indicates whether the style represents a time</param>
        /// <returns>Style object or null if the component could not be retrieved</returns>
        public Style GetStyle(int index, out bool isDateStyle, out bool isTimeStyle)
        {
            Style style = GetComponent(typeof(Style), index) as Style;
            isDateStyle = false;
            isTimeStyle = false;
            if (style != null)
            {
                isDateStyle = NumberFormat.IsDateFormat(style.CurrentNumberFormat.Number);
                isTimeStyle = NumberFormat.IsTimeFormat(style.CurrentNumberFormat.Number);
            }
            return style;
        }

        /// <summary>
        /// Returns a number format component by its index
        /// </summary>
        /// <param name="index">Internal index of the style component</param>
        /// <returns>Style component or null if the component could not be retrieved</returns>
        public NumberFormat GetNumberFormat(int index)
        {
            return GetComponent(typeof(NumberFormat), index) as NumberFormat;
        }

        /// <summary>
        /// Returns a border component by its index
        /// </summary>
        /// <param name="index">Internal index of the style component</param>
        /// <returns>Style component or null if the component could not be retrieved</returns>
        public Border GetBorder(int index)
        {
            return GetComponent(typeof(Border), index) as Border;
        }

        /// <summary>
        /// Returns a fill component by its index
        /// </summary>
        /// <param name="index">Internal index of the style component</param>
        /// <returns>Style component or null if the component could not be retrieved</returns>
        public Fill GetFill(int index)
        {
            return GetComponent(typeof(Fill), index) as Fill;
        }

        /// <summary>
        /// Returns a font component by its index
        /// </summary>
        /// <param name="index">Internal index of the style component</param>
        /// <returns>Style component or null if the component could not be retrieved</returns>
        public Font GetFont(int index)
        {
            return GetComponent(typeof(Font), index) as Font;
        }

        /// <summary>
        /// Gets the next internal id of a style
        /// </summary>
        /// <returns>Next id of styles (collected in this class)</returns>
        public int GetNextStyleId()
        {
            return styles.Count;
        }

        /// <summary>
        /// Gets the next internal id of a cell XF component
        /// </summary>
        /// <returns>Next id of the component type (collected in this class)</returns>
        public int GetNextCellXFId()
        {
            return cellXfs.Count;
        }

        /// <summary>
        /// Gets the next internal id of a border component
        /// </summary>
        /// <returns>Next id of the component type (collected in this class)</returns>
        public int GetNextBorderId()
        {
            return borders.Count;
        }

        /// <summary>
        /// Gets the next internal id of a fill component
        /// </summary>
        /// <returns>Next id of the component type (collected in this class)</returns>
        public int GetNextFillId()
        {
            return fills.Count;
        }

        /// <summary>
        /// Gets the next internal id of a font component
        /// </summary>
        /// <returns>Next id of the component type (collected in this class)</returns>
        public int GetNextFontId()
        {
            return fonts.Count;
        }

        /// <summary>
        /// Internal method to retrieve style components
        /// </summary>
        /// \remark <remarks>CellXF is not handled, since retrieved in the style reader in a different way</remarks>
        /// <param name="type">Type of the style components</param>
        /// <param name="index">Internal index of the style components</param>
        /// <returns>Style component or null if the component could not be retrieved</returns>
        private AbstractStyle GetComponent(Type type, int index)
        {
            try
            {
                if (type == typeof(NumberFormat))
                {
                    //Number format entries are handles differently, since identified by 'numFmtId'. Other components are identified by its entry index
                    NumberFormat numberFormat = numberFormats.Find(x => x.InternalID == index);
                    if (numberFormat == null)
                    {
                        throw new StyleException("The number format with the numFmtId: " + index + " was not found");
                    }
                    return numberFormat;
                }
                else if (type == typeof(Style))
                {
                    return styles[index];
                }
                else if (type == typeof(Border))
                {
                    return borders[index];
                }
                else if (type == typeof(Fill))
                {
                    return fills[index];
                }
                else // must be font (CellXF is not handled here)
                {
                    return fonts[index];
                }
            }
            catch (Exception)
            {
                // Ignore
            }
            return null;
        }

        /// <summary>
        /// Adds a color value to the color MRU list
        /// </summary>
        /// <param name="value">ARGB value</param>
        internal void AddMruColor(string value)
        {
            this.mruColors.Add(value);
        }

        /// <summary>
        /// Gets the mRU colors as list
        /// </summary>
        /// <returns>ARGB value</returns>
        internal List<string> GetMruColors()
        {
            return this.mruColors;
        }

        #endregion

    }
}
