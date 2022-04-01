/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2022
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using NanoXLSX.Exceptions;
using NanoXLSX.LowLevel;
using NanoXLSX.Styles;

namespace NanoXLSX
{
    /// <summary>
    /// Class representing a workbook
    /// </summary>
    /// 
    public class Workbook
    {
        #region privateFields
        private string filename;
        private List<Worksheet> worksheets;
        private Worksheet currentWorksheet;
        private StyleManager styleManager;
        private Metadata workbookMetadata;
        private string workbookProtectionPassword;
        private bool lockWindowsIfProtected;
        private bool lockStructureIfProtected;
        private int selectedWorksheet;
        private Shortener shortener;
        private bool importInProgress = false;
        #endregion

        #region properties


        /// <summary>
        /// Gets the shortener object for the current worksheet
        /// </summary>
        public Shortener WS
        {
            get { return shortener; }
        }


        /// <summary>
        /// Gets the current worksheet
        /// </summary>
        public Worksheet CurrentWorksheet
        {
            get { return currentWorksheet; }
        }

        /// <summary>
        /// Gets or sets the filename of the workbook
        /// </summary>
        public string Filename
        {
            get { return filename; }
            set { filename = value; }
        }

        /// <summary>
        /// Gets whether the structure are locked if workbook is protected
        /// </summary>
        /// <see cref="SetWorkbookProtection"/>
        public bool LockStructureIfProtected
        {
            get { return lockStructureIfProtected; }
        }

        /// <summary>
        /// Gets whether the windows are locked if workbook is protected
        /// </summary>
        /// <see cref="SetWorkbookProtection"/> 
        public bool LockWindowsIfProtected
        {
            get { return lockWindowsIfProtected; }
        }

        /// <summary>
        /// Meta data object of the workbook
        /// </summary>
        public Metadata WorkbookMetadata
        {
            get { return workbookMetadata; }
            set { workbookMetadata = value; }
        }

        /// <summary>
        /// Gets the selected worksheet. The selected worksheet is not the current worksheet while design time but the selected sheet in the output file
        /// </summary>
        public int SelectedWorksheet
        {
            get { return selectedWorksheet; }
        }

        /// <summary>
        /// Gets the style manager of this workbook
        /// </summary>
        public StyleManager Styles
        {
            get { return styleManager; }
        }

        /// <summary>
        /// Gets or sets whether the workbook is protected
        /// </summary>
        public bool UseWorkbookProtection { get; set; }

        /// <summary>
        /// Gets the password used for workbook protection
        /// </summary>
        /// <see cref="SetWorkbookProtection"/>
        public string WorkbookProtectionPassword
        {
            get { return workbookProtectionPassword; }
        }

        /// <summary>
        /// Gets the list of worksheets in the workbook
        /// </summary>
        public List<Worksheet> Worksheets
        {
            get { return worksheets; }
        }


        /// <summary>
        /// Gets or sets whether the whole workbook is hidden
        /// </summary>
        /// <remarks>A hidden workbook can only be made visible, using another, already visible Excel window</remarks>
        public bool Hidden { get; set; }

        #endregion

        #region constructors
        /// <summary>
        /// Constructor with additional parameter to create a default worksheet. This constructor can be used to define a workbook that is saved as stream
        /// </summary>
        /// <param name="createWorkSheet">If true, a default worksheet with the name 'Sheet1' will be crated and set as current worksheet</param>
        public Workbook(bool createWorkSheet)
        {
            Init();
            if (createWorkSheet)
            {
                AddWorksheet("Sheet1");
            }
        }

        /// <summary>
        /// Constructor with additional parameter to create a default worksheet with the specified name. This constructor can be used to define a workbook that is saved as stream
        /// </summary>
        /// <param name="sheetName">Filename of the workbook.  The name will be sanitized automatically according to the specifications of Excel</param>
        public Workbook(string sheetName)
        {
            Init();
            AddWorksheet(sheetName, true);
        }

        /// <summary>
        /// Constructor with filename ant the name of the first worksheet
        /// </summary>
        /// <param name="filename">Filename of the workbook.  The name will be sanitized automatically according to the specifications of Excel</param>
        /// <param name="sheetName">Name of the first worksheet. The name will be sanitized automatically according to the specifications of Excel</param>
        public Workbook(string filename, string sheetName)
        {
            Init();
            this.filename = filename;
            AddWorksheet(sheetName, true);
        }

        /// <summary>
        /// Constructor with filename ant the name of the first worksheet
        /// </summary>
        /// <param name="filename">Filename of the workbook</param>
        /// <param name="sheetName">Name of the first worksheet</param>
        /// <param name="sanitizeSheetName">If true, the name of the worksheet will be sanitized automatically according to the specifications of Excel</param>
        public Workbook(string filename, string sheetName, bool sanitizeSheetName)
        {
            Init();
            this.filename = filename;
            AddWorksheet(Worksheet.SanitizeWorksheetName(sheetName, this));
        }

        #endregion

        #region methods_PICO

        /// <summary>
        /// Adds a style to the style manager
        /// </summary>
        /// <param name="style">Style to add</param>
        /// <returns>Returns the managed style of the style manager</returns>

        public Style AddStyle(Style style)
        {
            return styleManager.AddStyle(style);
        }

        /// <summary>
        /// Adds a style component to a style
        /// </summary>
        /// <param name="baseStyle">Style to append a component</param>
        /// <param name="newComponent">Component to add to the baseStyle</param>
        /// <returns>Returns the managed style of the style manager</returns>
        public Style AddStyleComponent(Style baseStyle, AbstractStyle newComponent)
        {

            if (newComponent.GetType() == typeof(Border))
            {
                baseStyle.CurrentBorder = (Border)newComponent;
            }
            else if (newComponent.GetType() == typeof(CellXf))
            {
                baseStyle.CurrentCellXf = (CellXf)newComponent;
            }
            else if (newComponent.GetType() == typeof(Fill))
            {
                baseStyle.CurrentFill = (Fill)newComponent;
            }
            else if (newComponent.GetType() == typeof(Font))
            {
                baseStyle.CurrentFont = (Font)newComponent;
            }
            else if (newComponent.GetType() == typeof(NumberFormat))
            {
                baseStyle.CurrentNumberFormat = (NumberFormat)newComponent;
            }
            return styleManager.AddStyle(baseStyle);
        }


        /// <summary>
        /// Adding a new Worksheet. The new worksheet will be defined as current worksheet
        /// </summary>
        /// <param name="name">Name of the new worksheet</param>
        /// <exception cref="WorksheetException">Throws a WorksheetNameAlreadxExistsException if the name of the worksheet already exists</exception>
        /// <exception cref="Exceptions.FormatException">Throws a FormatException if the name contains illegal characters or is out of range (length between 1 an 31 characters)</exception>
        public void AddWorksheet(string name)
        {
            foreach (Worksheet item in worksheets)
            {
                if (item.SheetName == name)
                {
                    throw new WorksheetException("WorksheetNameAlreadxExistsException", "The worksheet with the name '" + name + "' already exists.");
                }
            }
            int number = worksheets.Count + 1;
            Worksheet newWs = new Worksheet(name, number, this);
            currentWorksheet = newWs;
            worksheets.Add(newWs);
            shortener.SetCurrentWorksheet(currentWorksheet);
        }

        /// <summary>
        /// Adding a new Worksheet with a sanitizing option. The new worksheet will be defined as current worksheet
        /// </summary>
        /// <param name="name">Name of the new worksheet</param>
        /// <param name="sanitizeSheetName">If true, the name of the worksheet will be sanitized automatically according to the specifications of Excel</param>
        /// <exception cref="WorksheetException">WorksheetException is thrown if the name of the worksheet already exists and sanitizeSheetName is false</exception>
        /// <exception cref="Exceptions.FormatException">FormatException is thrown if the worksheet name contains illegal characters or is out of range (length between 1 an 31) and sanitizeSheetName is false</exception>
        public void AddWorksheet(String name, bool sanitizeSheetName)
        {
            if (sanitizeSheetName)
            {
                string sanitized = Worksheet.SanitizeWorksheetName(name, this);
                AddWorksheet(sanitized);
            }
            else
            {
                AddWorksheet(name);
            }
        }

        /// <summary>
        /// Adding a new Worksheet. The new worksheet will be defined as current worksheet
        /// </summary>
        /// <param name="worksheet">Prepared worksheet object</param>
        /// <exception cref="WorksheetException">WorksheetException is thrown if the name of the worksheet already exists</exception>
        /// <exception cref="Exceptions.FormatException">FormatException is thrown if the worksheet name contains illegal characters or is out of range (length between 1 an 31</exception>
        public void AddWorksheet(Worksheet worksheet)
        {
            for (int i = 0; i < worksheets.Count; i++)
            {
                if (worksheets[i].SheetName == worksheet.SheetName)
                {
                    throw new WorksheetException("WorksheetNameAlreadyExistsException", "The worksheet with the name '" + worksheet.SheetName + "' already exists.");
                }
            }
            int number = worksheets.Count + 1;
            worksheet.SheetID = number;
            currentWorksheet = worksheet;
            worksheets.Add(worksheet);
            worksheet.WorkbookReference = this;
        }

        /// <summary>
        /// Init method called in the constructors
        /// </summary>
        private void Init()
        {
            worksheets = new List<Worksheet>();
            styleManager = new StyleManager();
            styleManager.AddStyle(new Style("default", 0, true));
            Style borderStyle = new Style("default_border_style", 1, true);
            borderStyle.CurrentBorder = BasicStyles.DottedFill_0_125.CurrentBorder;
            borderStyle.CurrentFill = BasicStyles.DottedFill_0_125.CurrentFill;
            styleManager.AddStyle(borderStyle);
            workbookMetadata = new Metadata();
            shortener = new Shortener();
        }


        /// <summary>
        /// Removes the passed style from the style sheet
        /// </summary>
        /// <param name="style">Style to remove</param>
        /// <exception cref="StyleException">Throws a StyleException if the style was not found in the style collection (could not be referenced)</exception>
        public void RemoveStyle(Style style)
        {
            RemoveStyle(style, false);
        }

        /// <summary>
        /// Removes the defined style from the style sheet of the workbook
        /// </summary>
        /// <param name="styleName">Name of the style to be removed</param>
        /// <exception cref="StyleException">Throws a StyleException if the style was not found in the style collection (could not be referenced)</exception>
        public void RemoveStyle(string styleName)
        {
            RemoveStyle(styleName, false);
        }

        /// <summary>
        /// Removes the defined style from the style sheet of the workbook
        /// </summary>
        /// <param name="style">Style to remove</param>
        /// <param name="onlyIfUnused">If true, the style will only be removed if not used in any cell</param>
        /// <exception cref="StyleException">Throws a StyleException if the style was not found in the style collection (could not be referenced)</exception>
        public void RemoveStyle(Style style, bool onlyIfUnused)
        {
            if (style == null)
            {
                throw new StyleException("UndefinedStyleException", "The style to remove is not defined");
            }
            RemoveStyle(style.Name, onlyIfUnused);
        }

        /// <summary>
        /// Removes the defined style from the style sheet of the workbook
        /// </summary>
        /// <param name="styleName">Name of the style to be removed</param>
        /// <param name="onlyIfUnused">If true, the style will only be removed if not used in any cell</param>
        /// <exception cref="StyleException">Throws an UndefinedStyleException if the style was not found in the style collection (could not be referenced)</exception>
        public void RemoveStyle(string styleName, bool onlyIfUnused)
        {
            if (string.IsNullOrEmpty(styleName))
            {
                throw new StyleException("MissingReferenceException", "The style to remove is not defined (no name specified)");
            }
            if (onlyIfUnused)
            {
                bool styleInUse = false;
                for (int i = 0; i < worksheets.Count; i++)
                {
                    foreach (KeyValuePair<string, Cell> cell in worksheets[i].Cells)
                    {
                        if (cell.Value.CellStyle == null) { continue; }
                        if (cell.Value.CellStyle.Name == styleName)
                        {
                            styleInUse = true;
                            break;
                        }
                    }
                    if (styleInUse)
                    {
                        break;
                    }
                }
                if (!styleInUse)
                {
                    styleManager.RemoveStyle(styleName);
                }
            }
            else
            {
                styleManager.RemoveStyle(styleName);
            }
        }

        /// <summary>
        /// Removes the defined worksheet
        /// </summary>
        /// <param name="name">Name of the worksheet</param>
        /// <exception cref="WorksheetException">Throws a UnknownWorksheetException if the name of the worksheet is unknown or if removing would lead to an empty workbook.</exception>
        /// <remarks>Note that the selected Worksheet (index) must be changed to another worksheet when the selected one is removed.</remarks>
        public void RemoveWorksheet(string name)
        {
            bool exists = false;
            bool resetCurrent = false;
            int index = 0;
            for (int i = 0; i < worksheets.Count; i++)
            {
                if (worksheets[i].SheetName == name)
                {
                    index = i;
                    exists = true;
                    break;
                }
            }
            if (!exists)
            {
                throw new WorksheetException("UnknownWorksheetException", "The worksheet with the name '" + name + "' does not exist.");
            }
            if (worksheets[index].SheetName == currentWorksheet.SheetName)
            {
                resetCurrent = true;
            }
            worksheets.RemoveAt(index);
            if (worksheets.Count > 0)
            {
                for (int i = 0; i < worksheets.Count; i++)
                {
                    worksheets[i].SheetID = i + 1;
                    if (resetCurrent && i == 0)
                    {
                        currentWorksheet = worksheets[i];
                    }
                }
            }
            else
            {
                currentWorksheet = null;
            }
            if (selectedWorksheet > worksheets.Count - 1)
            {
                selectedWorksheet = worksheets.Count - 1;
            }
            ValidateWorksheets();
        }

        /// <summary>
        /// Method to resolve all merged cells in all worksheets. Only the value of the very first cell of the locked cells range will be visible. The other values are still present (set to EMPTY) but will not be stored in the worksheet.
        /// </summary>
        /// <exception cref="StyleException">Throws a StyleException if one of the styles of the merged cells cannot be referenced or is null</exception>
        public void ResolveMergedCells()
        {
            Style mergeStyle = BasicStyles.MergeCellStyle;
            int pos;
            List<Address> addresses;
            Cell cell;
            foreach (Worksheet sheet in worksheets)
            {
                foreach (KeyValuePair<string, Range> range in sheet.MergedCells)
                {
                    pos = 0;
                    addresses = Cell.GetCellRange(range.Value.StartAddress, range.Value.EndAddress) as List<Address>;
                    foreach (Address address in addresses)
                    {
                        if (!sheet.Cells.ContainsKey(address.ToString()))
                        {
                            cell = new Cell();
                            cell.DataType = Cell.CellType.EMPTY;
                            cell.RowNumber = address.Row;
                            cell.ColumnNumber = address.Column;
                            cell.WorksheetReference = sheet;
                            sheet.AddCell(cell, cell.ColumnNumber, cell.RowNumber);
                        }
                        else
                        {
                            cell = sheet.Cells[address.ToString()];
                        }
                        if (pos != 0)
                        {
                            cell.DataType = Cell.CellType.EMPTY;
                            cell.SetStyle(mergeStyle);
                        }
                        pos++;
                    }

                }
            }
        }

        /// <summary>
        /// Saves the workbook
        /// </summary>
        /// <exception cref="Exceptions.IOException">Throws IOException in case of an error</exception>
        /// <exception cref="RangeException">Throws a RangeException if the start or end address of a handled cell range was out of range</exception>
        /// <exception cref="Exceptions.FormatException">Throws a FormatException if a handled date cannot be translated to (Excel internal) OADate</exception>
        /// <exception cref="StyleException">Throws a StyleException if one of the styles of the workbook cannot be referenced or is null</exception>
        public void Save()
        {
            XlsxWriter l = new XlsxWriter(this);
            l.Save();
        }

        /// <summary>
        /// Saves the workbook asynchronous.
        /// </summary>
        /// <returns>Task object (void)</returns>
        /// <exception cref="Exceptions.IOException">May throw an IOException in case of an error. The asynchronous operation may hide the exception.</exception>
        /// <exception cref="RangeException">May throw a RangeException if the start or end address of a handled cell range was out of range. The asynchronous operation may hide the exception.</exception>
        /// <exception cref="Exceptions.FormatException">May throw a FormatException if a handled date cannot be translated to (Excel internal) OADate. The asynchronous operation may hide the exception.</exception>
        /// <exception cref="StyleException">May throw a StyleException if one of the styles of the workbook cannot be referenced or is null. The asynchronous operation may hide the exception.</exception>
        public async Task SaveAsync()
        {
            XlsxWriter l = new XlsxWriter(this);
            await l.SaveAsync();
        }

        /// <summary>
        /// Saves the workbook with the defined name
        /// </summary>
        /// <param name="filename">filename of the saved workbook</param>
        /// <exception cref="Exceptions.IOException">Throws IOException in case of an error</exception>
        /// <exception cref="RangeException">Throws a RangeException if the start or end address of a handled cell range was out of range</exception>
        /// <exception cref="Exceptions.FormatException">Throws a FormatException if a handled date cannot be translated to (Excel internal) OADate</exception>
        /// <exception cref="StyleException">Throws a StyleException if one of the styles of the workbook cannot be referenced or is null</exception>
        public void SaveAs(string filename)
        {
            string backup = filename;
            this.filename = filename;
            XlsxWriter l = new XlsxWriter(this);
            l.Save();
            this.filename = backup;
        }

        /// <summary>
        /// Saves the workbook with the defined name asynchronous.
        /// </summary>
        /// <param name="fileName">filename of the saved workbook</param>
        /// <returns>Task object (void)</returns>
        /// <exception cref="Exceptions.IOException">May throw an IOException in case of an error. The asynchronous operation may hide the exception.</exception>
        /// <exception cref="RangeException">May throw a RangeException if the start or end address of a handled cell range was out of range. The asynchronous operation may hide the exception.</exception>
        /// <exception cref="Exceptions.FormatException">May throw a FormatException if a handled date cannot be translated to (Excel internal) OADate. The asynchronous operation may hide the exception.</exception>
        /// <exception cref="StyleException">May throw a StyleException if one of the styles of the workbook cannot be referenced or is null. The asynchronous operation may hide the exception.</exception>
        public async Task SaveAsAsync(string fileName)
        {
            string backup = fileName;
            filename = fileName;
            XlsxWriter l = new XlsxWriter(this);
            await l.SaveAsync();
            filename = backup;
        }

        /// <summary>
        /// Save the workbook to a writable stream
        /// </summary>
        /// <param name="stream">Writable stream</param>
        /// <param name="leaveOpen">Optional parameter to keep the stream open after writing (used for MemoryStreams; default is false)</param>
        /// <exception cref="IOException">Throws IOException in case of an error</exception>
        /// <exception cref="RangeException">Throws a RangeException if the start or end address of a handled cell range was out of range</exception>
        /// <exception cref="FormatException">Throws a FormatException if a handled date cannot be translated to (Excel internal) OADate</exception>
        /// <exception cref="StyleException">Throws a StyleException if one of the styles of the workbook cannot be referenced or is null</exception>
        public void SaveAsStream(Stream stream, bool leaveOpen = false)
        {
            XlsxWriter l = new XlsxWriter(this);
            l.SaveAsStream(stream, leaveOpen);
        }

        /// <summary>
        /// Save the workbook to a writable stream asynchronous.
        /// </summary>
        /// <param name="stream">>Writable stream</param>
        /// <param name="leaveOpen">Optional parameter to keep the stream open after writing (used for MemoryStreams; default is false)</param>
        /// <returns>Task object (void)</returns>
        /// <exception cref="IOException">Throws IOException in case of an error. The asynchronous operation may hide the exception.</exception>
        /// <exception cref="RangeException">May throw a RangeException if the start or end address of a handled cell range was out of range. The asynchronous operation may hide the exception.</exception>
        /// <exception cref="FormatException">May throw a FormatException if a handled date cannot be translated to (Excel internal) OADate. The asynchronous operation may hide the exception.</exception>
        /// <exception cref="StyleException">May throw a StyleException if one of the styles of the workbook cannot be referenced or is null. The asynchronous operation may hide the exception.</exception>
        public async Task SaveAsStreamAsync(Stream stream, bool leaveOpen = false)
        {
            XlsxWriter l = new XlsxWriter(this);
            await l.SaveAsStreamAsync(stream, leaveOpen);
        }

        /// <summary>
        /// Sets the current worksheet
        /// </summary>
        /// <param name="name">Name of the worksheet</param>
        /// <returns>Returns the current worksheet</returns>
        /// <exception cref="WorksheetException">Throws a MissingReferenceException if the name of the worksheet is unknown</exception>
        public Worksheet SetCurrentWorksheet(string name)
        {
            bool exists = false;
            foreach (Worksheet item in worksheets)
            {
                if (item.SheetName == name)
                {
                    currentWorksheet = item;
                    exists = true;
                    break;
                }
            }
            if (!exists)
            {
                throw new WorksheetException("MissingReferenceException", "The worksheet with the name '" + name + "' does not exist.");
            }
            shortener.SetCurrentWorksheet(currentWorksheet);
            return currentWorksheet;
        }

        /// <summary>
        /// Sets the selected worksheet in the output workbook
        /// </summary>
        /// <remarks>This method does not set the current worksheet while design time. Use SetCurrentWorksheet instead for this</remarks>
        /// <param name="worksheetIndex">Zero-based worksheet index</param>
        /// <exception cref="RangeException">Throws a RangeException if the index of the worksheet is out of range</exception>
        /// <exception cref="WorksheetException">Throws a WorksheetException if the worksheet to be set selected is hidden</exception>
        public void SetSelectedWorksheet(int worksheetIndex)
        {
            if (worksheetIndex < 0 || worksheetIndex > worksheets.Count - 1)
            {
                throw new RangeException(RangeException.GENERAL, "The worksheet index " + worksheetIndex + " is out of range");
            }
            selectedWorksheet = worksheetIndex;
            ValidateWorksheets();
        }

        /// <summary>
        /// Sets or removes the workbook protection. If protectWindows and protectStructure are both false, the workbook will not be protected
        /// </summary>
        /// <param name="state">If true, the workbook will be protected, otherwise not</param>
        /// <param name="protectWindows">If true, the windows will be locked if the workbook is protected</param>
        /// <param name="protectStructure">If true, the structure will be locked if the workbook is protected</param>
        /// <param name="password">Optional password. If null or empty, no password will be set in case of protection</param>
        public void SetWorkbookProtection(bool state, bool protectWindows, bool protectStructure, string password)
        {
            lockWindowsIfProtected = protectWindows;
            lockStructureIfProtected = protectStructure;
            workbookProtectionPassword = password;
            if (!protectWindows && !protectStructure)
            {
                UseWorkbookProtection = false;
            }
            else
            {
                UseWorkbookProtection = state;
            }
        }

        /// <summary>
        /// Sets the selected worksheet in the output workbook
        /// </summary>
        /// <remarks>This method does not set the current worksheet while design time. Use SetCurrentWorksheet instead for this</remarks>
        /// <param name="worksheet">Worksheet object (must be in the collection of worksheets)</param>
        /// <exception cref="WorksheetException">Throws a WorksheetException if the worksheet was not found in the worksheet collection or if it is hidden</exception>
        public void SetSelectedWorksheet(Worksheet worksheet)
        {
            bool check = false;
            for (int i = 0; i < worksheets.Count; i++)
            {
                if (worksheets[i].Equals(worksheet))
                {
                    selectedWorksheet = i;
                    check = true;
                    break;
                }
            }
            if (!check)
            {
                throw new WorksheetException("UnknownWorksheetException", "The passed worksheet object is not in the worksheet collection.");
            }
            ValidateWorksheets();
        }

        /// <summary>
        /// Validates the worksheets regarding several conditions that must be met:<br/>
        /// - At least one worksheet must be defined<br/>
        /// - A hidden worksheet cannot be the selected one<br/>
        /// - At least one worksheet must be visible<br/>
        /// If one of the conditions is not met, an exception is thrown
        /// </summary>
        /// <remarks>If an import is in progress, these rules are disabled to avoid conflicts by the order of loaded worksheets</remarks>
        internal void ValidateWorksheets()
        {
            if (importInProgress)
            {
                // No validation during import
                return;
            }
            int worksheetCount = worksheets.Count;
            if (worksheetCount == 0)
            {
                throw new WorksheetException("NoWorksheetsException", "The workbook must contain at least one worksheet");
            }
            for(int i = 0; i < worksheetCount; i++)
            {
                if (worksheets[i].Hidden)
                {
                    if (i == selectedWorksheet)
                    {
                        throw new WorksheetException("InvalidWorksheetRefernceException", "The worksheet with the index " + selectedWorksheet + " cannot be set as selected, since it is set hidden");
                    }
                }
            }
        }

        #endregion

        #region methods_NANO

        /// <summary>
        /// Loads a workbook from a file
        /// </summary>
        /// <param name="filename">Filename of the workbook</param>
        /// <param name="options">Import options to override the data types of columns. These options can be used to cope with wrong interpreted data, caused by irregular styles</param>
        /// <returns>Workbook object</returns>
        /// <exception cref="Exceptions.IOException">Throws IOException in case of an error</exception>
        public static Workbook Load(String filename, ImportOptions options = null)
        {
            XlsxReader r = new XlsxReader(filename, options);
            r.Read();
            return r.GetWorkbook();
        }

        /// <summary>
        /// Loads a workbook from a stream
        /// </summary>
        /// <param name="stream">Stream containing the workbook</param>
        /// /// <param name="options">Import options to override the data types of columns. These options can be used to cope with wrong interpreted data, caused by irregular styles</param>
        /// <returns>Workbook object</returns>
        /// <exception cref="Exceptions.IOException">Throws IOException in case of an error</exception>
        public static Workbook Load(Stream stream, ImportOptions options = null)
        {
            XlsxReader r = new XlsxReader(stream, options);
            r.Read();
            return r.GetWorkbook();
        }

        /// <summary>
        /// Loads a workbook from a stream asynchronously
        /// </summary>
        /// <param name="stream">Stream containing the workbook</param>
        /// /// <param name="options">Import options to override the data types of columns. These options can be used to cope with wrong interpreted data, caused by irregular styles</param>
        /// <returns>Workbook object</returns>
        /// <exception cref="Exceptions.IOException">Throws IOException in case of an error</exception>
        public static async Task<Workbook> LoadAsync(Stream stream, ImportOptions options = null)
        {
            XlsxReader r = new XlsxReader(stream, options);
            await r.ReadAsync();
            return r.GetWorkbook();
        }

        /// <summary>
        /// Sets the import state. If an import is in progress, no validity checks on are performed to avoid conflicts by incomplete data (e.g. hidden worksheets)
        /// </summary>
        /// <param name="state">True if an import is in progress, otherwise false</param>
        internal void SetImportState(bool state)
        {
            this.importInProgress = state;
        }

        #endregion
    }

    #region doc
    /// <summary>
    /// Main namespace with all high-level classes and functions to create or read workbooks and worksheets
    /// </summary>
    [CompilerGenerated]
    class NamespaceDoc // This class is only for documentation purpose (Sandcastle)
    { }
    #endregion
}
