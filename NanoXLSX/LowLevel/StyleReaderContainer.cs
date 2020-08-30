/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2020
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Exceptions;
using NanoXLSX.Styles;
using System;
using System.Collections.Generic;

namespace NanoXLSX.LowLevel
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
            else
            {
                throw new StyleException("StyleException", "The style definition of the type '" + t.ToString() + "' is unknown or not implemented yet");
            }
        }

        /// <summary>
        /// Returns a whole style by its index
        /// </summary>
        /// <param name="index">Index of the style</param>
        /// <param name="returnNullOnFail">If true, null will be returned if the style could not be retrieved. Otherwise an exception will be thrown</param>
        /// <exception cref="StyleException">Throws a StyleException if the style was not found and the parameter returnNullOnFail was set to false</exception>
        /// <returns>Style object or null if parameter returnNullOnFail was set to true and the component could not be retrieved</returns>
        public Style GetStyle(string index, bool returnNullOnFail = false)
        {
            int number;
            if (int.TryParse(index, out number))
            {
                return GetComponnet(typeof(Style), number, returnNullOnFail) as Style;
            }
            else if (returnNullOnFail)
            {
                return null;
            }
            else
            {
                throw new StyleException("StyleException", "The style definition could not be retrieved, because of the invalid style index '" + index + "'");
            }
        }

        /// <summary>
        /// Returns a whole style by its index. It also returns information about the type of the style
        /// </summary>
        /// <param name="index">Index of the style</param>
        /// <param name="isDateStyle">Out parameter that indicates whether the style represents a date</param>
        /// <param name="isTimeStyle">Out parameter that indicates whether the style represents a time</param>
        /// <param name="returnNullOnFail">If true, null will be returned if the style could not be retrieved. Otherwise an exception will be thrown</param>
        /// <exception cref="StyleException">Throws a StyleException if the style was not found and the parameter returnNullOnFail was set to false</exception>
        /// <returns>Style object or null if parameter returnNullOnFail was set to true and the component could not be retrieved</returns>
        public Style GetStyle(int index, out bool isDateStyle, out bool isTimeStyle, bool returnNullOnFail = false)
        {
            Style style = GetComponnet(typeof(Style), index, returnNullOnFail) as Style;
            if (style != null)
            {
                isDateStyle = NumberFormat.IsDateFormat(style.CurrentNumberFormat.Number);
                isTimeStyle = NumberFormat.IsTimeFormat(style.CurrentNumberFormat.Number);
            }
            else
            {
                isDateStyle = false;
                isTimeStyle = false;
            }
            return style;
        }

        /// <summary>
        /// Returns a cell XF component by its index
        /// </summary>
        /// <param name="index">Internal index of the style component</param>
        /// <param name="returnNullOnFail">If true, null will be returned if the component could not be retrieved. Otherwise an exception will be thrown</param>
        /// <exception cref="StyleException">Throws a StyleException if the component was not found and the parameter returnNullOnFail was set to false</exception>
        /// <remarks>The method is currently not used but prepared for usage when the style reader is fully implemented</remarks>
        /// <returns>Style component or null if parameter returnNullOnFail was set to true and the component could not be retrieved</returns>
        public CellXf GetCellXF(int index, bool returnNullOnFail = false)
        {
            return GetComponnet(typeof(CellXf), index, returnNullOnFail) as CellXf;
        }

        /// <summary>
        /// Returns a number format component by its index
        /// </summary>
        /// <param name="index">Internal index of the style component</param>
        /// <param name="returnNullOnFail">If true, null will be returned if the component could not be retrieved. Otherwise an exception will be thrown</param>
        /// <exception cref="StyleException">Throws a StyleException if the component was not found and the parameter returnNullOnFail was set to false</exception>
        /// <returns>Style component or null if parameter returnNullOnFail was set to true and the component could not be retrieved</returns>
        public NumberFormat GetNumberFormat(int index, bool returnNullOnFail = false)
        {
            return GetComponnet(typeof(NumberFormat), index, returnNullOnFail) as NumberFormat;
        }

        /// <summary>
        /// Returns a border component by its index
        /// </summary>
        /// <param name="index">Internal index of the style component</param>
        /// <param name="returnNullOnFail">If true, null will be returned if the component could not be retrieved. Otherwise an exception will be thrown</param>
        /// <exception cref="StyleException">Throws a StyleException if the component was not found and the parameter returnNullOnFail was set to false</exception>
        /// <remarks>The method is currently not used but prepared for usage when the style reader is fully implemented</remarks>
        /// <returns>Style component or null if parameter returnNullOnFail was set to true and the component could not be retrieved</returns>
        public Border GetBorder(int index, bool returnNullOnFail = false)
        {
            return GetComponnet(typeof(Border), index, returnNullOnFail) as Border;
        }

        /// <summary>
        /// Returns a fill component by its index
        /// </summary>
        /// <param name="index">Internal index of the style component</param>
        /// <param name="returnNullOnFail">If true, null will be returned if the component could not be retrieved. Otherwise an exception will be thrown</param>
        /// <exception cref="StyleException">Throws a StyleException if the component was not found and the parameter returnNullOnFail was set to false</exception>
        /// <remarks>The method is currently not used but prepared for usage when the style reader is fully implemented</remarks>
        /// <returns>Style component or null if parameter returnNullOnFail was set to true and the component could not be retrieved</returns>
        public Fill GetFill(int index, bool returnNullOnFail = false)
        {
            return GetComponnet(typeof(Fill), index, returnNullOnFail) as Fill;
        }

        /// <summary>
        /// Returns a font component by its index
        /// </summary>
        /// <param name="index">Internal index of the style component</param>
        /// <param name="returnNullOnFail">If true, null will be returned if the component could not be retrieved. Otherwise an exception will be thrown</param>
        /// <exception cref="StyleException">Throws a StyleException if the component was not found and the parameter returnNullOnFail was set to false</exception>
        /// <remarks>The method is currently not used but prepared for usage when the style reader is fully implemented</remarks>
        /// <returns>Style component or null if parameter returnNullOnFail was set to true and the component could not be retrieved</returns>
        public Font GetFont(int index, bool returnNullOnFail = false)
        {
            return GetComponnet(typeof(Font), index, returnNullOnFail) as Font;
        }

        /// <summary>
        /// Gets the next internal id of a style
        /// </summary>
        /// <returns>Next id of styles (collected in this class)</returns>
        public int GetNextStyleId()
        {
            return styles.Count + 1;
        }

        /// <summary>
        /// Gets the next internal id of a cell XF component
        /// </summary>
        /// <remarks>The method is currently not used but prepared for usage when the style reader is fully implemented</remarks>
        /// <returns>Next id of the component type (collected in this class)</returns>
        public int GetNextCellXFId()
        {
            return cellXfs.Count + 1;
        }

        /// <summary>
        /// Gets the next internal id of a number format component
        /// </summary>
        /// <returns>Next id of the component type (collected in this class)</returns>
        public int GetNextNumberFormatId()
        {
            return numberFormats.Count + 1;
        }

        /// <summary>
        /// Gets the next internal id of a border component
        /// </summary>
        /// <remarks>The method is currently not used but prepared for usage when the style reader is fully implemented</remarks>
        /// <returns>Next id of the component type (collected in this class)</returns>
        public int GetNextBorderId()
        {
            return borders.Count + 1;
        }

        /// <summary>
        /// Gets the next internal id of a fill component
        /// </summary>
        /// <remarks>The method is currently not used but prepared for usage when the style reader is fully implemented</remarks>
        /// <returns>Next id of the component type (collected in this class)</returns>
        public int GetNextFillId()
        {
            return fills.Count + 1;
        }

        /// <summary>
        /// Gets the next internal id of a font component
        /// </summary>
        /// <remarks>The method is currently not used but prepared for usage when the style reader is fully implemented</remarks>
        /// <returns>Next id of the component type (collected in this class)</returns>
        public int GetNextFontId()
        {
            return fonts.Count + 1;
        }

        /// <summary>
        /// Internal method to retrieve style components
        /// </summary>
        /// <param name="type">Type of the style components</param>
        /// <param name="index">Internal index of the style components</param>
        /// <param name="returnNullOnFail">If true, null will be returned if the component could not be retrieved. Otherwise an exception will be thrown</param>
        /// <exception cref="StyleException">Throws a StyleException if the component was not found and the parameter returnNullOnFail was set to false</exception>
        /// <returns>Style component or null if parameter returnNullOnFail was set to true and the component could not be retrieved</returns>
        private AbstractStyle GetComponnet(Type type, int index, bool returnNullOnFail)
        {
            try
            {
                if (type == typeof(CellXf))
                {
                    return cellXfs[index];
                }
                else if (type == typeof(NumberFormat))
                {
                    //Number format entries are handles differently, since identified by 'numFmtId'. Other components are identified by its entry index
                    NumberFormat numberFormat = numberFormats.Find(x => x.InternalID == index);
                    if (numberFormat == null)
                    {
                        throw new StyleException("StyleException", "The number format with the numFmtId: " + index + " was not found");
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
                else if (type == typeof(Font))
                {
                    return fonts[index];
                }
                else
                {
                    throw new StyleException("StyleException", "The style definition of the type '" + type.ToString() + "' is unknown or not implemented yet");
                }
            }
            catch (Exception ex)
            {
                if (returnNullOnFail)
                {
                    return null;
                }
                else
                {
                    throw new StyleException("StyleException", "The style definition could not be retrieved. Please see inner exception:", ex);
                }
            }
        }
        #endregion

    }
}
