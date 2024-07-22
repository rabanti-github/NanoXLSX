/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2024
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using NanoXLSX.Exceptions;

namespace NanoXLSX.Styles
{
    /// <summary>
    /// Class representing a style manager to maintain all styles and its components of a workbook. 
    /// This class is only internally used to compose the style environment right before saving an XLSX file
    /// </summary>
    internal class StyleManager
    {
        #region privateFields
        private List<AbstractStyle> borders;
        private List<AbstractStyle> cellXfs;
        private List<AbstractStyle> fills;
        private List<AbstractStyle> fonts;
        private List<AbstractStyle> numberFormats;
        private List<AbstractStyle> styles;
        #endregion

        #region constructors
        /// <summary>
        /// Default constructor
        /// </summary>
        public StyleManager()
        {
            borders = new List<AbstractStyle>();
            cellXfs = new List<AbstractStyle>();
            fills = new List<AbstractStyle>();
            fonts = new List<AbstractStyle>();
            numberFormats = new List<AbstractStyle>();
            styles = new List<AbstractStyle>();
        }
        #endregion

        #region methods

        /// <summary>
        /// Gets a component by its hash
        /// </summary>
        /// <param name="list">List to check</param>
        /// <param name="hash">Hash of the component</param>
        /// <returns>Determined component. If not found, null will be returned</returns>
        private static AbstractStyle GetComponentByHash(ref List<AbstractStyle> list, int hash)
        {
            int len = list.Count;
            for (int i = 0; i < len; i++)
            {
                if (list[i].GetHashCode() == hash)
                {
                    return list[i];
                }
            }
            return null;
        }

        /// <summary>
        /// Gets all borders of the style manager
        /// </summary>
        /// <returns>Array of borders</returns>
        public Border[] GetBorders()
        {
            return Array.ConvertAll(borders.ToArray(), x => (Border)x);
        }

        /// <summary>
        /// Gets the number of borders in the style manager
        /// </summary>
        /// <returns>Number of stored borders</returns>
        public int GetBorderStyleNumber()
        {
            return borders.Count;
        }

        /* ****************************** */

        /// <summary>
        /// Gets all fills of the style manager
        /// </summary>
        /// <returns>Array of fills</returns>
        public Fill[] GetFills()
        {
            return Array.ConvertAll(fills.ToArray(), x => (Fill)x);
        }

        /// <summary>
        /// Gets the number of fills in the style manager
        /// </summary>
        /// <returns>Number of stored fills</returns>
        public int GetFillStyleNumber()
        {
            return fills.Count;
        }

        /* ****************************** */

        /// <summary>
        /// Gets all fonts of the style manager
        /// </summary>
        /// <returns>Array of fonts</returns>
        public Font[] GetFonts()
        {
            return Array.ConvertAll(fonts.ToArray(), x => (Font)x);
        }

        /// <summary>
        /// Gets the number of fonts in the style manager
        /// </summary>
        /// <returns>Number of stored fonts</returns>
        public int GetFontStyleNumber()
        {
            return fonts.Count;
        }

        /* ****************************** */

        /// <summary>
        /// Gets all numberFormats of the style manager
        /// </summary>
        /// <returns>Array of numberFormats</returns>
        public NumberFormat[] GetNumberFormats()
        {
            return Array.ConvertAll(numberFormats.ToArray(), x => (NumberFormat)x);
        }

        /// <summary>
        /// Gets the number of numberFormats in the style manager
        /// </summary>
        /// <returns>Number of stored numberFormats</returns>
        public int GetNumberFormatStyleNumber()
        {
            return numberFormats.Count;
        }

        /* ****************************** */

        /// <summary>
        /// Gets all styles of the style manager
        /// </summary>
        /// <returns>Array of styles</returns>
        public Style[] GetStyles()
        {
            return Array.ConvertAll(styles.ToArray(), x => (Style)x);
        }

        /// <summary>
        /// Gets the number of styles in the style manager
        /// </summary>
        /// <returns>Number of stored styles</returns>
        public int GetStyleNumber()
        {
            return styles.Count;
        }

        /* ****************************** */


        /// <summary>
        /// Adds a style component to the manager
        /// </summary>
        /// <param name="style">Style to add</param>
        /// <returns>Added or determined style in the manager</returns>
        public Style AddStyle(Style style)
        {
            int hash = AddStyleComponent(style);
            return (Style)GetComponentByHash(ref styles, hash);
        }

        /// <summary>
        /// Adds a style component to the manager with an ID
        /// </summary>
        /// <param name="style">Component to add</param>
        /// <param name="id">Id of the component</param>
        /// <returns>Hash of the added or determined component</returns>
        private int AddStyleComponent(AbstractStyle style, int? id)
        {
            style.InternalID = id;
            return AddStyleComponent(style);
        }

        /// <summary>
        /// Adds a style component to the manager
        /// </summary>
        /// <param name="style">Component to add</param>
        /// <returns>Hash of the added or determined component</returns>
        private int AddStyleComponent(AbstractStyle style)
        {
            int hash = style.GetHashCode();
            if (style.GetType() == typeof(Border))
            {
                if (GetComponentByHash(ref borders, hash) == null)
                { borders.Add(style); }
                Reorganize(ref borders);
            }
            else if (style.GetType() == typeof(CellXf))
            {
                if (GetComponentByHash(ref cellXfs, hash) == null)
                { cellXfs.Add(style); }
                Reorganize(ref cellXfs);
            }
            else if (style.GetType() == typeof(Fill))
            {
                if (GetComponentByHash(ref fills, hash) == null)
                { fills.Add(style); }
                Reorganize(ref fills);
            }
            else if (style.GetType() == typeof(Font))
            {
                if (GetComponentByHash(ref fonts, hash) == null)
                { fonts.Add(style); }
                Reorganize(ref fonts);
            }
            else if (style.GetType() == typeof(NumberFormat))
            {
                if (GetComponentByHash(ref numberFormats, hash) == null)
                { numberFormats.Add(style); }
                Reorganize(ref numberFormats);
            }
            else if (style.GetType() == typeof(Style))
            {
                Style s = (Style)style;
                if (GetComponentByHash(ref styles, hash) == null)
                {
                    int? id;
                    if (!s.InternalID.HasValue)
                    {
                        id = int.MaxValue;
                        s.InternalID = id;
                    }
                    else
                    {
                        id = s.InternalID.Value;
                    }
                    int temp = AddStyleComponent(s.CurrentBorder, id);
                    s.CurrentBorder = (Border)GetComponentByHash(ref borders, temp);
                    temp = AddStyleComponent(s.CurrentCellXf, id);
                    s.CurrentCellXf = (CellXf)GetComponentByHash(ref cellXfs, temp);
                    temp = AddStyleComponent(s.CurrentFill, id);
                    s.CurrentFill = (Fill)GetComponentByHash(ref fills, temp);
                    temp = AddStyleComponent(s.CurrentFont, id);
                    s.CurrentFont = (Font)GetComponentByHash(ref fonts, temp);
                    temp = AddStyleComponent(s.CurrentNumberFormat, id);
                    s.CurrentNumberFormat = (NumberFormat)GetComponentByHash(ref numberFormats, temp);
                    styles.Add(s);
                }
                Reorganize(ref styles);
                hash = s.GetHashCode();
            }
            return hash;
        }

        /// <summary>
        /// Method to gather all styles of the cells in all worksheets
        /// </summary>
        /// <param name="workbook">Workbook to get all cells with possible style definitions</param>
        /// <returns>StyleManager object, to be processed by the save methods</returns>
        internal static StyleManager GetManagedStyles(Workbook workbook)
        {
            StyleManager styleManager = new StyleManager();
            styleManager.AddStyle(new Style("default", 0, true));
            Style borderStyle = new Style("default_border_style", 1, true);
            borderStyle.CurrentBorder = BasicStyles.DottedFill_0_125.CurrentBorder;
            borderStyle.CurrentFill = BasicStyles.DottedFill_0_125.CurrentFill;
            styleManager.AddStyle(borderStyle);

            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                foreach (KeyValuePair<string, Cell> cell in workbook.Worksheets[i].Cells)
                {
                    if (cell.Value.CellStyle != null)
                    {
                        Style resolvedStyle = styleManager.AddStyle(cell.Value.CellStyle);
                        workbook.Worksheets[i].Cells[cell.Key].SetStyle(resolvedStyle, true);
                    }
                }
                foreach(KeyValuePair<int, Column> column in workbook.Worksheets[i].Columns)
				{
                    if (column.Value.DefaultColumnStyle != null)
					{
                        Style resolvedStyle = styleManager.AddStyle(column.Value.DefaultColumnStyle);
                        workbook.Worksheets[i].Columns[column.Key].SetDefaultColumnStyle(resolvedStyle, true);
					}
				}
            }
            return styleManager;
        }

        /// <summary>
        /// Method to reorganize / reorder a list of style components
        /// </summary>
        /// <param name="list">List to reorganize as reference</param>
        private static void Reorganize(ref List<AbstractStyle> list)
        {
            int len = list.Count;
            list.Sort();
            int id = 0;
            for (int i = 0; i < len; i++)
            {
                list[i].InternalID = id;
                id++;
            }
        }
        #endregion
    }

}
