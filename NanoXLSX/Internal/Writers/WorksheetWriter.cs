/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2024
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Internal.Structures;
using NanoXLSX.Shared.Interfaces;
using NanoXLSX.Shared.Utils;
using System.Linq;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace NanoXLSX.Internal.Writers
{
    internal class WorksheetWriter
    {
        private static CultureInfo CULTURE = CultureInfo.InvariantCulture;

        private readonly Workbook workbook;
        private readonly SortedMap sharedStrings;
        private readonly XlsxWriter writer;

        internal WorksheetWriter(XlsxWriter writer)
        {
            this.workbook = writer.Workbook;
            this.sharedStrings = writer.SharedStrings;
            this.writer = writer;
        }
       

        /// <summary>
        /// Method to create a worksheet part as a raw XML string
        /// </summary>
        /// <param name="worksheet">worksheet object to process</param>
        /// <returns>Raw XML string</returns>
        /// <exception cref="NanoXLSX.Shared.Exceptions.FormatException">Throws a FormatException if a handled date cannot be translated to (Excel internal) OADate</exception>
        internal string CreateWorksheetPart(Worksheet worksheet)
        {
            worksheet.RecalculateAutoFilter();
            worksheet.RecalculateColumns();
            StringBuilder sb = new StringBuilder();
            sb.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">");
            if (worksheet.GetLastCellAddress().HasValue && worksheet.GetFirstCellAddress().HasValue)
            {
                sb.Append("<dimension ref=\"").Append(new Range(worksheet.GetFirstCellAddress().Value, worksheet.GetLastCellAddress().Value)).Append("\"/>");
            }
            if (worksheet.SelectedCells.Count > 0 || worksheet.PaneSplitTopHeight != null || worksheet.PaneSplitLeftWidth != null || worksheet.PaneSplitAddress != null ||
                worksheet.Hidden || worksheet.ZoomFactor != 100 || worksheet.ZoomFactors.Count > 1 || !worksheet.ShowGridLines || !worksheet.ShowRuler || !worksheet.ShowRowColumnHeaders || worksheet.ViewType != Worksheet.SheetViewType.normal)
            {
                CreateSheetViewString(worksheet, sb);
            }
            sb.Append("<sheetFormatPr");
            if (!HasPaneSplitting(worksheet))
            {
                // TODO: Find the right calculation to compensate baseColWidth when using pane splitting
                sb.Append(" defaultColWidth=\"")
             .Append(ParserUtils.ToString(worksheet.DefaultColumnWidth))
                .Append("\"");
            }
            sb.Append(" defaultRowHeight=\"")
             .Append(ParserUtils.ToString(worksheet.DefaultRowHeight))
             .Append("\" baseColWidth=\"")
             .Append(ParserUtils.ToString(worksheet.DefaultColumnWidth))
             .Append("\" x14ac:dyDescent=\"0.25\"/>");

            string colDefinitions = CreateColsString(worksheet);
            if (!string.IsNullOrEmpty(colDefinitions))
            {
                sb.Append("<cols>");
                sb.Append(colDefinitions);
                sb.Append("</cols>");
            }
            sb.Append("<sheetData>");
            CreateRowsString(worksheet, sb);
            sb.Append("</sheetData>");
            sb.Append(CreateMergedCellsString(worksheet));
            sb.Append(CreateSheetProtectionString(worksheet));

            if (worksheet.AutoFilterRange != null)
            {
                sb.Append("<autoFilter ref=\"").Append(worksheet.AutoFilterRange.Value.ToString()).Append("\"/>");
            }

            sb.Append("</worksheet>");
            return sb.ToString();
        }

        /// <summary>
        /// Method to create a row string
        /// </summary>
        /// <param name="dynamicRow">Dynamic row with List of cells, heights and hidden states</param>
        /// <param name="worksheet">Worksheet to process</param>
        /// <returns>Formatted row string</returns>
        /// <exception cref="NanoXLSX.Shared.Exceptions.FormatException">Throws a FormatException if a handled date cannot be translated to (Excel internal) OADate</exception>
        private string CreateRowString(DynamicRow dynamicRow, Worksheet worksheet)
        {
            int rowNumber = dynamicRow.RowNumber;
            string height = "";
            string hidden = "";
            if (worksheet.RowHeights.ContainsKey(rowNumber))
            {
                if (worksheet.RowHeights[rowNumber] != worksheet.DefaultRowHeight)
                {
                    height = " x14ac:dyDescent=\"0.25\" customHeight=\"1\" ht=\"" + ParserUtils.ToString(Utils.GetInternalRowHeight(worksheet.RowHeights[rowNumber])) + "\"";
                }
            }
            if (worksheet.HiddenRows.ContainsKey(rowNumber))
            {
                if (worksheet.HiddenRows[rowNumber])
                {
                    hidden = " hidden=\"1\"";
                }
            }
            StringBuilder sb = new StringBuilder();
            sb.Append("<row r=\"").Append((rowNumber + 1).ToString()).Append("\"").Append(height).Append(hidden).Append(">");
            string typeAttribute;
            string styleDef;
            string typeDef;
            string valueDef = "";
            bool boolValue;

            int col = 0;
            foreach (Cell item in dynamicRow.CellDefinitions)
            {
                // Data type must be resolved
                typeDef = " ";
                if (item.CellStyle != null)
                {
                    styleDef = " s=\"" + ParserUtils.ToString(item.CellStyle.InternalID.Value) + "\" ";
                }
                else
                {
                    styleDef = "";
                }
                if (item.DataType == Cell.CellType.BOOL)
                {
                    typeAttribute = "b";
                    typeDef = " t=\"" + typeAttribute + "\" ";
                    boolValue = (bool)item.Value;
                    if (boolValue) { valueDef = "1"; }
                    else { valueDef = "0"; }

                }
                // Number casting
                else if (item.DataType == Cell.CellType.NUMBER)
                {
                    typeAttribute = "n";
                    typeDef = " t=\"" + typeAttribute + "\" ";
                    Type t = item.Value.GetType();

                    if (t == typeof(byte)) { valueDef = ((byte)item.Value).ToString("G", CULTURE); }
                    else if (t == typeof(sbyte)) { valueDef = ((sbyte)item.Value).ToString("G", CULTURE); }
                    else if (t == typeof(decimal)) { valueDef = ((decimal)item.Value).ToString("G", CULTURE); }
                    else if (t == typeof(double)) { valueDef = ((double)item.Value).ToString("G", CULTURE); }
                    else if (t == typeof(float)) { valueDef = ((float)item.Value).ToString("G", CULTURE); }
                    else if (t == typeof(int)) { valueDef = ((int)item.Value).ToString("G", CULTURE); }
                    else if (t == typeof(uint)) { valueDef = ((uint)item.Value).ToString("G", CULTURE); }
                    else if (t == typeof(long)) { valueDef = ((long)item.Value).ToString("G", CULTURE); }
                    else if (t == typeof(ulong)) { valueDef = ((ulong)item.Value).ToString("G", CULTURE); }
                    else if (t == typeof(short)) { valueDef = ((short)item.Value).ToString("G", CULTURE); }
                    else if (t == typeof(ushort)) { valueDef = ((ushort)item.Value).ToString("G", CULTURE); }
                }
                // Date parsing
                else if (item.DataType == Cell.CellType.DATE)
                {
                    DateTime date = (DateTime)item.Value;
                    valueDef = Utils.GetOADateTimeString(date);
                }
                // Time parsing
                else if (item.DataType == Cell.CellType.TIME)
                {
                    TimeSpan time = (TimeSpan)item.Value;
                    valueDef = Utils.GetOATimeString(time);
                }
                else
                {
                    if (item.Value == null)
                    {
                        typeAttribute = null;
                        valueDef = null;
                    }
                    else // Handle sharedStrings
                    {
                        if (item.DataType == Cell.CellType.FORMULA)
                        {
                            typeAttribute = "str";
                            valueDef = item.Value.ToString();
                        }
                        else
                        {
                            typeAttribute = "s";
                            if (item.Value is IFormattableText)
                            {
                                valueDef = sharedStrings.Add((IFormattableText)item.Value, ParserUtils.ToString(sharedStrings.Count));
                            }
                            else
                            {
                                valueDef = sharedStrings.Add(new PlainText(item.Value.ToString()), ParserUtils.ToString(sharedStrings.Count));
                            }
                            this.writer.SharedStringsTotalCount++;
                        }
                    }
                    typeDef = " t=\"" + typeAttribute + "\" ";
                }
                if (item.DataType != Cell.CellType.EMPTY)
                {
                    sb.Append("<c r=\"").Append(item.CellAddress).Append("\"").Append(typeDef).Append(styleDef).Append(">");
                    if (item.DataType == Cell.CellType.FORMULA)
                    {
                        sb.Append("<f>").Append(XmlUtils.EscapeXmlChars(item.Value.ToString())).Append("</f>");
                    }
                    else
                    {
                        sb.Append("<v>").Append(XmlUtils.EscapeXmlChars(valueDef)).Append("</v>");
                    }
                    sb.Append("</c>");
                }
                else if (valueDef == null || item.DataType == Cell.CellType.EMPTY) // Empty cell
                {
                    sb.Append("<c r=\"").Append(item.CellAddress).Append("\"").Append(styleDef).Append("/>");
                }
                col++;
            }
            sb.Append("</row>");
            return sb.ToString();
        }

        /// <summary>
        /// Method to create the merged cells string of the passed worksheet
        /// </summary>
        /// <param name="sheet">Worksheet to process</param>
        /// <returns>Formatted string with merged cell ranges</returns>
        private string CreateMergedCellsString(Worksheet sheet)
        {
            if (sheet.MergedCells.Count < 1)
            {
                return string.Empty;
            }
            StringBuilder sb = new StringBuilder();
            sb.Append("<mergeCells count=\"").Append(ParserUtils.ToString(sheet.MergedCells.Count)).Append("\">");
            foreach (KeyValuePair<string, Range> item in sheet.MergedCells)
            {
                sb.Append("<mergeCell ref=\"").Append(item.Value.ToString()).Append("\"/>");
            }
            sb.Append("</mergeCells>");
            return sb.ToString();
        }

        /// <summary>
        /// Method to create the (sub) part of the sheet view (selected cells and panes) within the worksheet XML document
        /// </summary>
        /// <param name="worksheet">worksheet object to process</param>
        /// <param name="sb">reference to the StringBuilder</param>
        private void CreateSheetViewString(Worksheet worksheet, StringBuilder sb)
        {
            sb.Append("<sheetViews><sheetView workbookViewId=\"0\"");
            if (workbook.SelectedWorksheet == worksheet.SheetID - 1 && !worksheet.Hidden)
            {
                sb.Append(" tabSelected=\"1\"");
            }
            if (worksheet.ViewType != Worksheet.SheetViewType.normal)
            {
                if (worksheet.ViewType == Worksheet.SheetViewType.pageLayout)
                {
                    if (worksheet.ShowRuler)
                    {
                        sb.Append(" showRuler=\"1\"");
                    }
                    else
                    {
                        sb.Append(" showRuler=\"0\"");
                    }
                    sb.Append(" view=\"pageLayout\"");

                }
                else if (worksheet.ViewType == Worksheet.SheetViewType.pageBreakPreview)
                {
                    sb.Append(" view=\"pageBreakPreview\"");
                }
            }
            if (!worksheet.ShowGridLines)
            {
                sb.Append(" showGridLines=\"0\"");
            }
            if (!worksheet.ShowRowColumnHeaders)
            {
                sb.Append("  showRowColHeaders=\"0\"");
            }
            sb.Append(" zoomScale=\"").Append(ParserUtils.ToString(worksheet.ZoomFactor)).Append("\"");
            foreach (KeyValuePair<Worksheet.SheetViewType, int> scaleFactor in worksheet.ZoomFactors)
            {
                if (scaleFactor.Key == worksheet.ViewType)
                {
                    continue;
                }
                if (scaleFactor.Key == Worksheet.SheetViewType.normal)
                {
                    sb.Append(" zoomScaleNormal=\"").Append(ParserUtils.ToString(scaleFactor.Value)).Append("\"");
                }
                else if (scaleFactor.Key == Worksheet.SheetViewType.pageBreakPreview)
                {
                    sb.Append(" zoomScaleSheetLayoutView=\"").Append(ParserUtils.ToString(scaleFactor.Value)).Append("\"");
                }
                else if (scaleFactor.Key == Worksheet.SheetViewType.pageLayout)
                {
                    sb.Append(" zoomScalePageLayoutView=\"").Append(ParserUtils.ToString(scaleFactor.Value)).Append("\"");
                }
            }
            sb.Append(">");
            CreatePaneString(worksheet, sb);
            if (worksheet.SelectedCells.Count > 0)
            {
                sb.Append("<selection sqref=\"");
                for (int i = 0; i < worksheet.SelectedCells.Count; i++)
                {
                    sb.Append(worksheet.SelectedCells[i].ToString());
                    if (i < worksheet.SelectedCells.Count - 1)
                    {
                        sb.Append(" ");
                    }
                }
                sb.Append("\" activeCell=\"");
                sb.Append(worksheet.SelectedCells[0].StartAddress.ToString());
                sb.Append("\"/>");
            }
            sb.Append("</sheetView></sheetViews>");
        }

        /// <summary>
        /// Checks whether pane splitting is applied in the given worksheet
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns>True if applied, otherwise false</returns>
        private bool HasPaneSplitting(Worksheet worksheet)
        {
            if (worksheet.PaneSplitLeftWidth == null && worksheet.PaneSplitTopHeight == null && worksheet.PaneSplitAddress == null)
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// Method to create the enclosing part of the rows
        /// </summary>
        /// <param name="worksheet">Worksheet to process</param>
        /// <param name="sb">reference to the StringBuilder</param>
        private void CreateRowsString(Worksheet worksheet, StringBuilder sb)
        {
            List<DynamicRow> cellData = GetSortedSheetData(worksheet);
            string line;
            foreach (DynamicRow row in cellData)
            {
                line = CreateRowString(row, worksheet);
                sb.Append(line);
            }
        }

        /// <summary>
        /// Method to create the columns as XML string. This is used to define the width of columns
        /// </summary>
        /// <param name="worksheet">Worksheet to process</param>
        /// <returns>String with formatted XML data</returns>
        private string CreateColsString(Worksheet worksheet)
        {
            if (worksheet.Columns.Count > 0)
            {
                string col;
                string hidden = "";
                StringBuilder sb = new StringBuilder();
                foreach (KeyValuePair<int, Column> column in worksheet.Columns)
                {
                    if (column.Value.Width == worksheet.DefaultColumnWidth && !column.Value.IsHidden) { continue; }
                    if (worksheet.Columns.ContainsKey(column.Key))
                    {
                        if (worksheet.Columns[column.Key].IsHidden)
                        {
                            hidden = " hidden=\"1\"";
                        }
                    }
                    col = ParserUtils.ToString(column.Key + 1); // Add 1 for Address
                    float width = Utils.GetInternalColumnWidth(column.Value.Width);
                    sb.Append("<col customWidth=\"1\" width=\"").Append(ParserUtils.ToString(width)).Append("\" max=\"").Append(col).Append("\" min=\"").Append(col).Append("\"");
                    if (column.Value.DefaultColumnStyle != null)
                    {
                        sb.Append(" style=\"").Append(column.Value.DefaultColumnStyle.InternalID.Value.ToString("G", CULTURE)).Append("\"");
                    }
                    sb.Append(hidden).Append("/>");
                }
                string value = sb.ToString();
                if (value.Length > 0)
                {
                    return value;
                }
                return string.Empty;
            }
            return string.Empty;
        }

        /// <summary>
        /// Method to create the (sub) part of the pane (splitting and freezing) within the worksheet XML document
        /// </summary>
        /// <param name="worksheet">worksheet object to process</param>
        /// <param name="sb">reference to the StringBuilder</param>
        private void CreatePaneString(Worksheet worksheet, StringBuilder sb)
        {
            if (!HasPaneSplitting(worksheet))
            {
                return;
            }
            sb.Append("<pane");
            bool applyXSplit = false;
            bool applyYSplit = false;
            if (worksheet.PaneSplitAddress != null)
            {
                bool freeze = worksheet.FreezeSplitPanes != null && worksheet.FreezeSplitPanes.Value;
                int xSplit = worksheet.PaneSplitAddress.Value.Column;
                int ySplit = worksheet.PaneSplitAddress.Value.Row;
                if (xSplit > 0)
                {
                    if (freeze)
                    {
                        sb.Append(" xSplit=\"").Append(ParserUtils.ToString(xSplit)).Append("\"");
                    }
                    else
                    {
                        sb.Append(" xSplit=\"").Append(ParserUtils.ToString(CalculatePaneWidth(worksheet, xSplit))).Append("\"");
                    }
                    applyXSplit = true;
                }
                if (ySplit > 0)
                {
                    if (freeze)
                    {
                        sb.Append(" ySplit=\"").Append(ParserUtils.ToString(ySplit)).Append("\"");
                    }
                    else
                    {
                        sb.Append(" ySplit=\"").Append(ParserUtils.ToString(CalculatePaneHeight(worksheet, ySplit))).Append("\"");
                    }
                    applyYSplit = true;
                }
                if (freeze && applyXSplit && applyYSplit)
                {
                    sb.Append(" state=\"frozenSplit\"");
                }
                else if (freeze)
                {
                    sb.Append(" state=\"frozen\"");
                }
            }
            else
            {
                if (worksheet.PaneSplitLeftWidth != null)
                {
                    sb.Append(" xSplit=\"").Append(ParserUtils.ToString(Utils.GetInternalPaneSplitWidth(worksheet.PaneSplitLeftWidth.Value))).Append("\"");
                    applyXSplit = true;
                }
                if (worksheet.PaneSplitTopHeight != null)
                {
                    sb.Append(" ySplit=\"").Append(ParserUtils.ToString(Utils.GetInternalPaneSplitHeight(worksheet.PaneSplitTopHeight.Value))).Append("\"");
                    applyYSplit = true;
                }
            }
            if ((applyXSplit || applyYSplit) && worksheet.ActivePane != null)
            {
                switch (worksheet.ActivePane.Value)
                {
                    case Worksheet.WorksheetPane.bottomLeft:
                        sb.Append(" activePane=\"bottomLeft\"");
                        break;
                    case Worksheet.WorksheetPane.bottomRight:
                        sb.Append(" activePane=\"bottomRight\"");
                        break;
                    case Worksheet.WorksheetPane.topLeft:
                        sb.Append(" activePane=\"topLeft\"");
                        break;
                    case Worksheet.WorksheetPane.topRight:
                        sb.Append(" activePane=\"topRight\"");
                        break;
                }
            }
            string topLeftCell = worksheet.PaneSplitTopLeftCell.Value.GetAddress();
            sb.Append(" topLeftCell=\"").Append(topLeftCell).Append("\" ");
            sb.Append("/>");
            if (applyXSplit && !applyYSplit)
            {
                sb.Append("<selection pane=\"topRight\" activeCell=\"" + topLeftCell + "\"  sqref=\"" + topLeftCell + "\" />");
            }
            else if (applyYSplit && !applyXSplit)
            {
                sb.Append("<selection pane=\"bottomLeft\" activeCell=\"" + topLeftCell + "\"  sqref=\"" + topLeftCell + "\" />");
            }
            else if (applyYSplit && applyXSplit)
            {
                sb.Append("<selection activeCell=\"" + topLeftCell + "\"  sqref=\"" + topLeftCell + "\" />");
            }
        }

        /// <summary>
        /// Method to calculate the pane height, based on the number of rows
        /// </summary>
        /// <param name="worksheet">worksheet object to get the row definitions from</param>
        /// <param name="numberOfRows">Number of rows from the top to the split position</param>
        /// <returns>Internal height from the top of the worksheet to the pane split position</returns>
        private float CalculatePaneHeight(Worksheet worksheet, int numberOfRows)
        {
            float height = 0;
            for (int i = 0; i < numberOfRows; i++)
            {
                if (worksheet.RowHeights.ContainsKey(i))
                {
                    height += Utils.GetInternalRowHeight(worksheet.RowHeights[i]);
                }
                else
                {
                    height += Utils.GetInternalRowHeight(Worksheet.DEFAULT_ROW_HEIGHT);
                }
            }
            return Utils.GetInternalPaneSplitHeight(height);
        }



        /// <summary>
        /// Method to calculate the pane width, based on the number of columns
        /// </summary>
        /// <param name="worksheet">worksheet object to get the column definitions from</param>
        /// <param name="numberOfColumns">Number of columns from the left to the split position</param>
        /// <returns>Internal width from the left of the worksheet to the pane split position</returns>
        private float CalculatePaneWidth(Worksheet worksheet, int numberOfColumns)
        {
            float width = 0;
            for (int i = 0; i < numberOfColumns; i++)
            {
                if (worksheet.Columns.ContainsKey(i))
                {
                    width += Utils.GetInternalColumnWidth(worksheet.Columns[i].Width);
                }
                else
                {
                    width += Utils.GetInternalColumnWidth(Worksheet.DEFAULT_COLUMN_WIDTH);
                }
            }
            // Add padding of 75 per column
            return Utils.GetInternalPaneSplitWidth(width) + ((numberOfColumns - 1) * 0f);
        }

        /// <summary>
        /// Method to create the protection string of the passed worksheet
        /// </summary>
        /// <param name="sheet">Worksheet to process</param>
        /// <returns>Formatted string with protection statement of the worksheet</returns>
        private string CreateSheetProtectionString(Worksheet sheet)
        {
            if (!sheet.UseSheetProtection)
            {
                return string.Empty;
            }
            Dictionary<Worksheet.SheetProtectionValue, int> actualLockingValues = new Dictionary<Worksheet.SheetProtectionValue, int>();
            if (sheet.SheetProtectionValues.Count == 0)
            {
                actualLockingValues.Add(Worksheet.SheetProtectionValue.selectLockedCells, 1);
                actualLockingValues.Add(Worksheet.SheetProtectionValue.selectUnlockedCells, 1);
            }
            if (!sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.objects))
            {
                actualLockingValues.Add(Worksheet.SheetProtectionValue.objects, 1);
            }
            if (!sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.scenarios))
            {
                actualLockingValues.Add(Worksheet.SheetProtectionValue.scenarios, 1);
            }
            if (!sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.selectLockedCells))
            {
                if (!actualLockingValues.ContainsKey(Worksheet.SheetProtectionValue.selectLockedCells))
                {
                    actualLockingValues.Add(Worksheet.SheetProtectionValue.selectLockedCells, 1);
                }
            }
            if (!sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.selectUnlockedCells) || !sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.selectLockedCells))
            {
                if (!actualLockingValues.ContainsKey(Worksheet.SheetProtectionValue.selectUnlockedCells))
                {
                    actualLockingValues.Add(Worksheet.SheetProtectionValue.selectUnlockedCells, 1);
                }
            }
            if (sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.formatCells)) { actualLockingValues.Add(Worksheet.SheetProtectionValue.formatCells, 0); }
            if (sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.formatColumns)) { actualLockingValues.Add(Worksheet.SheetProtectionValue.formatColumns, 0); }
            if (sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.formatRows)) { actualLockingValues.Add(Worksheet.SheetProtectionValue.formatRows, 0); }
            if (sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.insertColumns)) { actualLockingValues.Add(Worksheet.SheetProtectionValue.insertColumns, 0); }
            if (sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.insertRows)) { actualLockingValues.Add(Worksheet.SheetProtectionValue.insertRows, 0); }
            if (sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.insertHyperlinks)) { actualLockingValues.Add(Worksheet.SheetProtectionValue.insertHyperlinks, 0); }
            if (sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.deleteColumns)) { actualLockingValues.Add(Worksheet.SheetProtectionValue.deleteColumns, 0); }
            if (sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.deleteRows)) { actualLockingValues.Add(Worksheet.SheetProtectionValue.deleteRows, 0); }
            if (sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.sort)) { actualLockingValues.Add(Worksheet.SheetProtectionValue.sort, 0); }
            if (sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.autoFilter)) { actualLockingValues.Add(Worksheet.SheetProtectionValue.autoFilter, 0); }
            if (sheet.SheetProtectionValues.Contains(Worksheet.SheetProtectionValue.pivotTables)) { actualLockingValues.Add(Worksheet.SheetProtectionValue.pivotTables, 0); }
            StringBuilder sb = new StringBuilder();
            sb.Append("<sheetProtection");
            string temp;
            foreach (KeyValuePair<Worksheet.SheetProtectionValue, int> item in actualLockingValues)
            {
                temp = Enum.GetName(typeof(Worksheet.SheetProtectionValue), item.Key); // Note! If the enum names differs from the OOXML definitions, this method will cause invalid OOXML entries
                sb.Append(" ").Append(temp).Append("=\"").Append(ParserUtils.ToString(item.Value)).Append("\"");
            }
            if (!string.IsNullOrEmpty(sheet.SheetProtectionPasswordHash))
            {
                sb.Append(" password=\"").Append(sheet.SheetProtectionPasswordHash).Append("\"");
            }
            sb.Append(" sheet=\"1\"/>");
            return sb.ToString();
        }

        /// <summary>
        /// Method to sort the cells of a worksheet as preparation for the XML document
        /// </summary>
        /// <param name="sheet">Worksheet to process</param>
        /// <returns>Sorted list of dynamic rows that are either defined by cells or row widths / hidden states. The list is sorted by row numbers (zero-based)</returns>
        private List<DynamicRow> GetSortedSheetData(Worksheet sheet)
        {
            List<Cell> temp = new List<Cell>();
            foreach (KeyValuePair<string, Cell> item in sheet.Cells)
            {
                temp.Add(item.Value);
            }
            temp.Sort();
            DynamicRow row = new DynamicRow(); ;
            Dictionary<int, DynamicRow> rows = new Dictionary<int, DynamicRow>();
            int rowNumber;
            if (temp.Count > 0)
            {
                rowNumber = temp[0].RowNumber;
                row.RowNumber = rowNumber;
                foreach (Cell cell in temp)
                {
                    if (cell.RowNumber != rowNumber)
                    {
                        rows.Add(rowNumber, row);
                        row = new DynamicRow();
                        row.RowNumber = cell.RowNumber;
                        rowNumber = cell.RowNumber;
                    }
                    row.CellDefinitions.Add(cell);
                }
                if (row.CellDefinitions.Count > 0)
                {
                    rows.Add(rowNumber, row);
                }
            }
            foreach (KeyValuePair<int, float> rowHeight in sheet.RowHeights)
            {
                if (!rows.ContainsKey(rowHeight.Key))
                {
                    row = new DynamicRow();
                    row.RowNumber = rowHeight.Key;
                    rows.Add(rowHeight.Key, row);
                }
            }
            foreach (KeyValuePair<int, bool> hiddenRow in sheet.HiddenRows)
            {
                if (!rows.ContainsKey(hiddenRow.Key))
                {
                    row = new DynamicRow();
                    row.RowNumber = hiddenRow.Key;
                    rows.Add(hiddenRow.Key, row);
                }
            }
            List<DynamicRow> output = rows.Values.ToList();
            output.Sort((r1, r2) => (r1.RowNumber.CompareTo(r2.RowNumber))); // Lambda sort
            return output;
        }


        #region helperClasses
        /// <summary>
        /// Class representing a row that is either empty or containing cells. Empty rows can also carry information about height or visibility
        /// </summary>
        internal class DynamicRow
        {
            private readonly List<Cell> cellDefinitions;
            public int RowNumber { get; set; }

            /// <summary>
            /// Gets the List of cells if not empty
            /// </summary>
            public List<Cell> CellDefinitions
            {
                get { return cellDefinitions; }
            }

            /// <summary>
            /// Default constructor. Defines an empty row if no additional operations are made on the object
            /// </summary>
            public DynamicRow()
            {
                this.cellDefinitions = new List<Cell>();
            }
        }
        #endregion

    }
}
