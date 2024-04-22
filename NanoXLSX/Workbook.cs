/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2024
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using NanoXLSX.Shared.Utils;
using NanoXLSX.Shared.Exceptions;
using NanoXLSX.Styles;
using NanoXLSX.Internal.Readers;
using NanoXLSX.Internal.Writers;
using NanoXLSX.Themes;

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
        private Metadata workbookMetadata;
        private string workbookProtectionPassword;
        private bool lockWindowsIfProtected;
        private bool lockStructureIfProtected;
        private int selectedWorksheet;
        private Shortener shortener;
        private List<string> mruColors = new List<string>();
        internal bool importInProgress = false;
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
        /// Gets or sets the filename of the workbook.
        /// </summary>
        /// <remarks>
        /// Note that the file name is not sanitized. If a filename is set that is not compliant to the file system, saving of the workbook may fail
        /// </remarks>
        public string Filename
        {
            get { return filename; }
            set { filename = value; }
        }

        /// <summary>
        /// Gets whether the structure are locked if workbook is protected. See also <see cref="SetWorkbookProtection"/>
        /// </summary>
        public bool LockStructureIfProtected
        {
            get { return lockStructureIfProtected; }
        }

        /// <summary>
        /// Gets whether the windows are locked if workbook is protected. See also <see cref="SetWorkbookProtection"/> 
        /// </summary>
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
        /// Gets or sets whether the workbook is protected
        /// </summary>
        public bool UseWorkbookProtection { get; set; }

        /// <summary>
        /// Gets the password used for workbook protection. See also <see cref="SetWorkbookProtection"/>
        /// </summary>
        /// <remarks>The password of this property is stored in plan text at runtime but not stored to a workbook. See also <see cref="WorkbookProtectionPasswordHash"/> for the generated hash</remarks>
        public string WorkbookProtectionPassword
        {
            get { return workbookProtectionPassword; }
        }

        /// <summary>
        /// Hash of the protected workbook, originated from <see cref="WorkbookProtectionPassword"/>
        /// </summary>
        /// <remarks>The plain text password cannot be recovered when loading a workbook. The hash is retrieved and can be reused, 
        /// if no changes are made in the area of workbook protection (<see cref="SetWorkbookProtection(bool, bool, bool, string)"/>)</remarks>
        public string WorkbookProtectionPasswordHash { get; internal set; }

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

        /// <summary>
        /// Gets or sets the theme of the workbook. The default is defined by <see cref="Theme.GetDefaultTheme"/>. However, the theme can be nullified
        /// </summary>
        public Theme WorkbookTheme { get; set; } = Theme.GetDefaultTheme();

        #endregion

        #region constructors
        /// <summary>
        /// Default constructor. No initial worksheet is created. Use <see cref="AddWorksheet(string)"/> (or overloads) to add one
        /// </summary>
        public Workbook()
        {
            Init();

        }

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
            if (sanitizeSheetName)
            {
                AddWorksheet(Worksheet.SanitizeWorksheetName(sheetName, this));
            }
            else
            {
                AddWorksheet(sheetName);
            }
        }

        #endregion

        #region methods_PICO

        /// <summary>
        /// Adds a color value (HEX; 6-digit RGB or 8-digit ARGB) to the MRU list
        /// </summary>
        /// <param name="color">RGB code in hex format (either 6 characters, e.g. FF00AC or 8 characters with leading alpha value). Alpha will be set to full opacity (FF) in case of 6 characters</param>
        public void AddMruColor(string color)
        {
            if (color != null && color.Length == 6)
            {
                color = "FF" + color;
            }
            Validators.ValidateColor(color, true);
            mruColors.Add(ParserUtils.ToUpper(color));
        }

        /// <summary>
        /// Gets the MRU color list
        /// </summary>
        /// <returns>Immutable list of color values</returns>
        public IReadOnlyList<string> GetMruColors()
        {
            return mruColors;
        }

        /// <summary>
        /// Clears the MRU color list
        /// </summary>
        public void ClearMruColors()
        {
            mruColors.Clear();
        }

        /// <summary>
        /// Adds a style to the style repository. This method is deprecated since it has no direct impact on the generated file.
        /// </summary>
        /// <param name="style">Style to add</param>
        /// <returns>Returns the managed style of the style repository</returns>
        /// 
        [Obsolete("This method has no direct impact on the generated file and is deprecated.")]
        public Style AddStyle(Style style)
        {
            return StyleRepository.Instance.AddStyle(style);
        }

        /// <summary>
        /// Adds a style component to a style. This method is deprecated since it has no direct impact on the generated file.
        /// </summary>
        /// <param name="baseStyle">Style to append a component</param>
        /// <param name="newComponent">Component to add to the baseStyle</param>
        /// <returns>Returns the modified style of the style repository</returns>
        [Obsolete("This method has no direct impact on the generated file and is deprecated.")]
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
            return StyleRepository.Instance.AddStyle(baseStyle);
        }

        /// <summary>
        /// Adding a new Worksheet. The new worksheet will be defined as current worksheet
        /// </summary>
        /// <param name="name">Name of the new worksheet</param>
        /// <exception cref="WorksheetException">Throws a WorksheetException if the name of the worksheet already exists</exception>
        /// <exception cref="NanoXLSX.Shared.Exceptions.FormatException">Throws a FormatException if the name contains illegal characters or is out of range (length between 1 an 31 characters)</exception>
        public void AddWorksheet(string name)
        {
            foreach (Worksheet item in worksheets)
            {
                if (item.SheetName == name)
                {
                    throw new WorksheetException("The worksheet with the name '" + name + "' already exists.");
                }
            }
            int number = GetNextWorksheetId();
            Worksheet newWs = new Worksheet(name, number, this);
            currentWorksheet = newWs;
            worksheets.Add(newWs);
            shortener.SetCurrentWorksheetInternal(currentWorksheet);
        }

        /// <summary>
        /// Adding a new Worksheet with a sanitizing option. The new worksheet will be defined as current worksheet
        /// </summary>
        /// <param name="name">Name of the new worksheet</param>
        /// <param name="sanitizeSheetName">If true, the name of the worksheet will be sanitized automatically according to the specifications of Excel</param>
        /// <exception cref="WorksheetException">WorksheetException is thrown if the name of the worksheet already exists and sanitizeSheetName is false</exception>
        /// <exception cref="NanoXLSX.Shared.Exceptions.FormatException">FormatException is thrown if the worksheet name contains illegal characters or is out of range (length between 1 an 31) and sanitizeSheetName is false</exception>
        public void AddWorksheet(string name, bool sanitizeSheetName)
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
        /// <exception cref="NanoXLSX.Shared.Exceptions.FormatException">FormatException is thrown if the worksheet name contains illegal characters or is out of range (length between 1 an 31)</exception>
        public void AddWorksheet(Worksheet worksheet)
        {
            AddWorksheet(worksheet, false);
        }

        /// <summary>
        /// Adding a new Worksheet. The new worksheet will be defined as current worksheet
        /// </summary>
        /// <param name="worksheet">Prepared worksheet object</param>
        /// <param name="sanitizeSheetName">If true, the name of the worksheet will be sanitized automatically according to the specifications of Excel</param>    
        /// <exception cref="WorksheetException">WorksheetException is thrown if the name of the worksheet already exists, when sanitation is false</exception>
        /// <exception cref="NanoXLSX.Shared.Exceptions.FormatException">FormatException is thrown if the worksheet name contains illegal characters or is out of range (length between 1 an 31) and sanitation is false</exception>
        public void AddWorksheet(Worksheet worksheet, bool sanitizeSheetName)
        {
            if (sanitizeSheetName)
            {
                string name = Worksheet.SanitizeWorksheetName(worksheet.SheetName, this);
                worksheet.SheetName = name;
            }
            else
            {
                if (string.IsNullOrEmpty(worksheet.SheetName))
                {
                    throw new WorksheetException("The name of the passed worksheet is null or empty.");
                }
                for (int i = 0; i < worksheets.Count; i++)
                {
                    if (worksheets[i].SheetName == worksheet.SheetName)
                    {
                        throw new WorksheetException("The worksheet with the name '" + worksheet.SheetName + "' already exists.");
                    }
                }
            }
            worksheet.SheetID = GetNextWorksheetId();
            currentWorksheet = worksheet;
            worksheets.Add(worksheet);
            worksheet.WorkbookReference = this;
        }

        /// <summary>
        /// Removes the passed style from the style sheet. This method is deprecated since it has no direct impact on the generated file.
        /// </summary>
        /// <param name="style">Style to remove</param>
        /// <remarks>Note: This method is available due to compatibility reasons. Added styles are actually not removed by it since unused styles are disposed automatically</remarks>
        [Obsolete("This method has no direct impact on the generated file and is deprecated.")]
        public void RemoveStyle(Style style)
        {
            RemoveStyle(style, false);
        }

        /// <summary>
        /// Removes the defined style from the style sheet of the workbook. This method is deprecated since it has no direct impact on the generated file.
        /// </summary>
        /// <param name="styleName">Name of the style to be removed</param>
        /// <remarks>Note: This method is available due to compatibility reasons. Added styles are actually not removed by it since unused styles are disposed automatically</remarks>
        [Obsolete("This method has no direct impact on the generated file and is deprecated.")]
        public void RemoveStyle(string styleName)
        {
            RemoveStyle(styleName, false);
        }

        /// <summary>
        /// Removes the defined style from the style sheet of the workbook
        /// </summary>
        /// <param name="style">Style to remove</param>
        /// <param name="onlyIfUnused">If true, the style will only be removed if not used in any cell</param>
        /// <remarks>Note: This method is available due to compatibility reasons. Added styles are actually not removed by it since unused styles are disposed automatically</remarks>
        [Obsolete("This method has no direct impact on the generated file and is deprecated.")]
        public void RemoveStyle(Style style, bool onlyIfUnused)
        {
            if (style == null) 
            {
                throw new StyleException("The style to remove is not defined");
            }
            RemoveStyle(style.Name, onlyIfUnused);
        }

        /// <summary>
        /// Removes the defined style from the style sheet of the workbook. This method is deprecated since it has no direct impact on the generated file.
        /// </summary>
        /// <param name="styleName">Name of the style to be removed</param>
        /// <param name="onlyIfUnused">If true, the style will only be removed if not used in any cell</param>
        /// <remarks>Note: This method is available due to compatibility reasons. Added styles are actually not removed by it since unused styles are disposed automatically</remarks>
        [Obsolete("This method has no direct impact on the generated file and is deprecated.")]
        public void RemoveStyle(string styleName, bool onlyIfUnused)
        {
            if (string.IsNullOrEmpty(styleName))
            {
                throw new StyleException("The style to remove is not defined (no name specified)");
            }
            // noOp / deprecated
        }

        /// <summary>
        /// Removes the defined worksheet based on its name. If the worksheet is the current or selected worksheet, the current and / or the selected worksheet will be set to the last worksheet of the workbook.
        /// If the last worksheet is removed, the selected worksheet will be set to 0 and the current worksheet to null.
        /// </summary>
        /// <param name="name">Name of the worksheet</param>
        /// <exception cref="WorksheetException">Throws a WorksheetException if the name of the worksheet is unknown</exception>
        public void RemoveWorksheet(string name)
        {
            Worksheet worksheetToRemove = worksheets.FirstOrDefault(w => w.SheetName == name);
            if (worksheetToRemove == null)
            {
                throw new WorksheetException("The worksheet with the name '" + name + "' does not exist.");
            }
            int index = worksheets.IndexOf(worksheetToRemove);
            bool resetCurrentWorksheet = worksheetToRemove == currentWorksheet;
            RemoveWorksheet(index, resetCurrentWorksheet);
        }

        /// <summary>
        /// Removes the defined worksheet based on its index. If the worksheet is the current or selected worksheet, the current and / or the selected worksheet will be set to the last worksheet of the workbook.
        /// If the last worksheet is removed, the selected worksheet will be set to 0 and the current worksheet to null.
        /// </summary>
        /// <param name="index">Index within the worksheets list</param>
        /// <exception cref="WorksheetException">Throws a WorksheetException if the index is out of range</exception>

        public void RemoveWorksheet(int index)
        {
            if (index < 0 || index >= worksheets.Count)
            {
                throw new WorksheetException("The worksheet index " + index + " is out of range");
            }
            bool resetCurrentWorksheet = worksheets[index] == currentWorksheet;
            RemoveWorksheet(index, resetCurrentWorksheet);
        }

        /// <summary>
        /// Method to resolve all merged cells in all worksheets. Only the value of the very first cell of the locked cells range will be visible. The other values are still present (set to EMPTY) but will not be stored in the worksheet.<br/>
        /// This is an internal method. There is no need to use it
        /// </summary>
        /// <exception cref="StyleException">Throws a StyleException if one of the styles of the merged cells cannot be referenced or is null</exception>
        internal void ResolveMergedCells()
        {
            foreach (Worksheet worksheet in worksheets)
            {
                worksheet.ResolveMergedCells();
            }
        }

        /// <summary>
        /// Saves the workbook
        /// </summary>
        /// <exception cref="NanoXLSX.Shared.Exceptions.IOException">Throws IOException in case of an error</exception>
        /// <exception cref="RangeException">Throws a RangeException if the start or end address of a handled cell range was out of range</exception>
        /// <exception cref="NanoXLSX.Shared.Exceptions.FormatException">Throws a FormatException if a handled date cannot be translated to (Excel internal) OADate</exception>
        public void Save()
        {
            XlsxWriter l = new XlsxWriter(this);
            l.Save();
        }

        /// <summary>
        /// Saves the workbook asynchronous.
        /// </summary>
        /// <returns>Task object (void)</returns>
        /// <exception cref="NanoXLSX.Shared.Exceptions.IOException">May throw an IOException in case of an error. The asynchronous operation may hide the exception.</exception>
        /// <exception cref="RangeException">May throw a RangeException if the start or end address of a handled cell range was out of range. The asynchronous operation may hide the exception.</exception>
        /// <exception cref="NanoXLSX.Shared.Exceptions.FormatException">May throw a FormatException if a handled date cannot be translated to (Excel internal) OADate. The asynchronous operation may hide the exception.</exception>
        public async Task SaveAsync()
        {
            XlsxWriter l = new XlsxWriter(this);
            await l.SaveAsync();
        }

        /// <summary>
        /// Saves the workbook with the defined name
        /// </summary>
        /// <param name="filename">filename of the saved workbook</param>
        /// <exception cref="NanoXLSX.Shared.Exceptions.IOException">Throws IOException in case of an error</exception>
        /// <exception cref="RangeException">Throws a RangeException if the start or end address of a handled cell range was out of range</exception>
        /// <exception cref="NanoXLSX.Shared.Exceptions.FormatException">Throws a FormatException if a handled date cannot be translated to (Excel internal) OADate</exception>
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
        /// <exception cref="NanoXLSX.Shared.Exceptions.IOException">May throw an IOException in case of an error. The asynchronous operation may hide the exception.</exception>
        /// <exception cref="RangeException">May throw a RangeException if the start or end address of a handled cell range was out of range. The asynchronous operation may hide the exception.</exception>
        /// <exception cref="NanoXLSX.Shared.Exceptions.FormatException">May throw a FormatException if a handled date cannot be translated to (Excel internal) OADate. The asynchronous operation may hide the exception.</exception>
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
        /// <exception cref="WorksheetException">Throws a WorksheetException if the name of the worksheet is unknown</exception>
        public Worksheet SetCurrentWorksheet(string name)
        {
            currentWorksheet = GetWorksheet(name);
            shortener.SetCurrentWorksheetInternal(currentWorksheet);
            return currentWorksheet;
        }

        /// <summary>
        /// Sets the current worksheet
        /// </summary>
        /// <param name="worksheetIndex">Zero-based worksheet index</param>
        /// <returns>Returns the current worksheet</returns>
        /// <exception cref="WorksheetException">Throws a WorksheetException if the name of the worksheet is unknown</exception>
        public Worksheet SetCurrentWorksheet(int worksheetIndex)
        {
            currentWorksheet = GetWorksheet(worksheetIndex);
            shortener.SetCurrentWorksheetInternal(currentWorksheet);
            return currentWorksheet;
        }

        /// <summary>
        /// Sets the current worksheet
        /// </summary>
        /// <param name="worksheet">Worksheet object (must be in the collection of worksheets)</param>
        /// <exception cref="WorksheetException">Throws a WorksheetException if the worksheet was not found in the worksheet collection</exception>
        public void SetCurrentWorksheet(Worksheet worksheet)
        {
            int index = worksheets.IndexOf(worksheet);
            if (index < 0)
            {
                throw new WorksheetException("The passed worksheet object is not in the worksheet collection.");
            }
            currentWorksheet = worksheets[index];
            shortener.SetCurrentWorksheetInternal(worksheet);
        }

        /// <summary>
        /// Sets the selected worksheet in the output workbook
        /// </summary>
        /// <param name="name">Name of the worksheet</param>
        /// <exception cref="WorksheetException">Throws a WorksheetException if the name of the worksheet is unknown</exception>
        public void SetSelectedWorksheet(string name)
        {
            int index = worksheets.FindIndex(w => w.SheetName == name);
            if (index < 0)
            {
                throw new WorksheetException("No worksheet with the name '" + name + "' was found in this workbook.");
            }
            selectedWorksheet = index;
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
                throw new RangeException("The worksheet index " + worksheetIndex + " is out of range");
            }
            selectedWorksheet = worksheetIndex;
            ValidateWorksheets();
        }

        /// <summary>
        /// Sets the selected worksheet in the output workbook
        /// </summary>
        /// <remarks>This method does not set the current worksheet while design time. Use SetCurrentWorksheet instead for this</remarks>
        /// <param name="worksheet">Worksheet object (must be in the collection of worksheets)</param>
        /// <exception cref="WorksheetException">Throws a WorksheetException if the worksheet was not found in the worksheet collection or if it is hidden</exception>
        public void SetSelectedWorksheet(Worksheet worksheet)
        {
            selectedWorksheet = worksheets.IndexOf(worksheet);
            if (selectedWorksheet < 0)
            {
                throw new WorksheetException("The passed worksheet object is not in the worksheet collection.");
            }
            ValidateWorksheets();
        }

        /// <summary>
        /// Gets a worksheet from this workbook by name
        /// </summary>
        /// <param name="name">Name of the worksheet</param>
        /// <returns>Worksheet with the passed name</returns>
        /// <exception cref="WorksheetException">Throws a WorksheetException if the worksheet was not found in the worksheet collection</exception>
        public Worksheet GetWorksheet(string name)
        {
            int index = worksheets.FindIndex(w => w.SheetName == name);
            if (index < 0)
            {
                throw new WorksheetException("No worksheet with the name '" + name + "' was found in this workbook.");
            }
            return worksheets[index];
        }

        /// <summary>
        /// Gets a worksheet from this workbook by index
        /// </summary>
        /// <param name="index">Index of the worksheet</param>
        /// <returns>Worksheet with the passed index</returns>
        /// <exception cref="WorksheetException">Throws a RangeException if the worksheet was not found in the worksheet collection</exception>
        public Worksheet GetWorksheet(int index)
        {
            if (index < 0 || index > worksheets.Count - 1)
            {
                throw new RangeException("The worksheet index " + index + " is out of range");
            }
            return worksheets[index];
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
            WorkbookProtectionPasswordHash = Utils.GeneratePasswordHash(password);
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
        /// Copies a worksheet of the current workbook by its name
        /// </summary>
        /// <param name="sourceWorksheetName">Name of the worksheet to copy, originated in this workbook</param>
        /// <param name="newWorksheetName">Name of the new worksheet (copy)</param>
        /// <param name="sanitizeSheetName">If true, the new name will be automatically sanitized if a name collision occurs</param>
        /// <remarks>The copy is not set as current worksheet. The existing one is kept</remarks>
        /// <returns>Copied worksheet</returns>
        public Worksheet CopyWorksheetIntoThis(string sourceWorksheetName, string newWorksheetName, bool sanitizeSheetName = true)
        {
            Worksheet sourceWorksheet = GetWorksheet(sourceWorksheetName);
            return CopyWorksheetTo(sourceWorksheet, newWorksheetName, this, sanitizeSheetName);
        }

        /// <summary>
        /// Copies a worksheet of the current workbook by its index
        /// </summary>
        /// <param name="sourceWorksheetIndex">Index of the worksheet to copy, originated in this workbook</param>
        /// <param name="newWorksheetName">Name of the new worksheet (copy)</param>
        /// <param name="sanitizeSheetName">If true, the new name will be automatically sanitized if a name collision occurs</param>
        /// <remarks>The copy is not set as current worksheet. The existing one is kept</remarks>
        /// <returns>Copied worksheet</returns>
        public Worksheet CopyWorksheetIntoThis(int sourceWorksheetIndex, string newWorksheetName, bool sanitizeSheetName = true)
        {
            Worksheet sourceWorksheet = GetWorksheet(sourceWorksheetIndex);
            return CopyWorksheetTo(sourceWorksheet, newWorksheetName, this, sanitizeSheetName);
        }

        /// <summary>
        /// Copies a worksheet of any workbook into the current workbook
        /// </summary>
        /// <param name="sourceWorksheet">Worksheet to copy</param>
        /// <param name="newWorksheetName">Name of the new worksheet (copy)</param>
        /// <param name="sanitizeSheetName">If true, the new name will be automatically sanitized if a name collision occurs</param>
        /// <remarks>The copy is not set as current worksheet. The existing one is kept. The source worksheet can originate from any workbook</remarks>
        /// <returns>Copied worksheet</returns>
        public Worksheet CopyWorksheetIntoThis(Worksheet sourceWorksheet, string newWorksheetName, bool sanitizeSheetName = true)
        {
            return CopyWorksheetTo(sourceWorksheet, newWorksheetName, this, sanitizeSheetName);
        }

        /// <summary>
        /// Copies a worksheet of the current workbook by its name into another workbook
        /// </summary>
        /// <param name="sourceWorksheetName">Name of the worksheet to copy, originated in this workbook</param>
        /// <param name="newWorksheetName">Name of the new worksheet (copy)</param>
        /// <param name="targetWorkbook">Workbook to copy the worksheet into</param>
        /// <param name="sanitizeSheetName">If true, the new name will be automatically sanitized if a name collision occurs</param>
        /// <remarks>The copy is not set as current worksheet. The existing one is kept</remarks>
        /// <returns>Copied worksheet</returns>
        public Worksheet CopyWorksheetTo(string sourceWorksheetName, string newWorksheetName, Workbook targetWorkbook, bool sanitizeSheetName = true)
        {
            Worksheet sourceWorksheet = GetWorksheet(sourceWorksheetName);
            return CopyWorksheetTo(sourceWorksheet, newWorksheetName, targetWorkbook, sanitizeSheetName);
        }

        /// <summary>
        /// Copies a worksheet of the current workbook by its index into another workbook
        /// </summary>
        /// <param name="sourceWorksheetIndex">Index of the worksheet to copy, originated in this workbook</param>
        /// <param name="newWorksheetName">Name of the new worksheet (copy)</param>
        /// <param name="targetWorkbook">Workbook to copy the worksheet into</param>
        /// <param name="sanitizeSheetName">If true, the new name will be automatically sanitized if a name collision occurs</param>
        /// <remarks>The copy is not set as current worksheet. The existing one is kept</remarks>
        /// <returns>Copied worksheet</returns>
        public Worksheet CopyWorksheetTo(int sourceWorksheetIndex, string newWorksheetName, Workbook targetWorkbook, bool sanitizeSheetName = true)
        {
            Worksheet sourceWorksheet = GetWorksheet(sourceWorksheetIndex);
            return CopyWorksheetTo(sourceWorksheet, newWorksheetName, targetWorkbook, sanitizeSheetName);
        }


        /// <summary>
        /// Copies a worksheet of any workbook into the another workbook
        /// </summary>
        /// <param name="sourceWorksheet">Worksheet to copy</param>
        /// <param name="newWorksheetName">Name of the new worksheet (copy)</param>
        /// <param name="targetWorkbook">Workbook to copy the worksheet into</param>
        /// <param name="sanitizeSheetName">If true, the new name will be automatically sanitized if a name collision occurs</param>
        /// <remarks>The copy is not set as current worksheet. The existing one is kept</remarks>
        /// <returns>Copied worksheet</returns>
        public static Worksheet CopyWorksheetTo(Worksheet sourceWorksheet, string newWorksheetName, Workbook targetWorkbook, bool sanitizeSheetName = true)
        {
            if (targetWorkbook == null)
            {
                throw new WorksheetException("The target workbook cannot be null");
            }
            if (sourceWorksheet == null)
            {
                throw new WorksheetException("The source worksheet cannot be null");
            }
            Worksheet copy = sourceWorksheet.Copy();
            copy.SetSheetName(newWorksheetName);
            Worksheet currentWorksheet = targetWorkbook.CurrentWorksheet;
            targetWorkbook.AddWorksheet(copy, sanitizeSheetName);
            targetWorkbook.SetCurrentWorksheet(currentWorksheet);
            return copy;
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
                throw new WorksheetException("The workbook must contain at least one worksheet");
            }
            for (int i = 0; i < worksheetCount; i++)
            {
                if (worksheets[i].Hidden)
                {
                    if (i == selectedWorksheet)
                    {
                        throw new WorksheetException("The worksheet with the index " + selectedWorksheet + " cannot be set as selected, since it is set hidden");
                    }
                }
            }
        }

        /// <summary>
        /// Removes the worksheet at the defined index and relocates current and selected worksheet references
        /// </summary>
        /// <param name="index">Index within the worksheets list</param>
        /// <param name="resetCurrentWorksheet">If true, the current worksheet will be relocated to the last worksheet in the list</param>
        private void RemoveWorksheet(int index, bool resetCurrentWorksheet)
        {
            worksheets.RemoveAt(index);
            if (worksheets.Count > 0)
            {
                for (int i = 0; i < worksheets.Count; i++)
                {
                    worksheets[i].SheetID = i + 1;
                }
                if (resetCurrentWorksheet)
                {
                    currentWorksheet = worksheets[worksheets.Count - 1];
                }
                if (selectedWorksheet == index || selectedWorksheet > worksheets.Count - 1)
                {
                    selectedWorksheet = worksheets.Count - 1;
                }
            }
            else
            {
                currentWorksheet = null;
                selectedWorksheet = 0;
            }
            ValidateWorksheets();
        }

        /// <summary>
        /// Gets the next free worksheet ID
        /// </summary>
        /// <returns>Worksheet ID</returns>
        private int GetNextWorksheetId()
        {
            if (worksheets.Count == 0)
            {
                return 1;
            }
            return worksheets.Max(w => w.SheetID) + 1;
        }

        /// <summary>
        /// Init method called in the constructors
        /// </summary>
        private void Init()
        {
            worksheets = new List<Worksheet>();
            workbookMetadata = new Metadata();
            shortener = new Shortener(this);
        }

        #endregion

        #region methods_NANO

        /// <summary>
        /// Loads a workbook from a file
        /// </summary>
        /// <param name="filename">Filename of the workbook</param>
        /// <param name="options">Import options to override the data types of columns or cells. These options can be used to cope with wrong interpreted data, caused by irregular styles</param>
        /// <returns>Workbook object</returns>
        /// <exception cref="NanoXLSX.Shared.Exceptions.IOException">Throws IOException in case of an error</exception>
        public static Workbook Load(string filename, ImportOptions options = null)
        {
            XlsxReader r = new XlsxReader(filename, options);
            r.Read();
            return r.GetWorkbook();
        }

        /// <summary>
        /// Loads a workbook from a stream
        /// </summary>
        /// <param name="stream">Stream containing the workbook</param>
        /// /// <param name="options">Import options to override the data types of columns or cells. These options can be used to cope with wrong interpreted data, caused by irregular styles</param>
        /// <returns>Workbook object</returns>
        /// <exception cref="NanoXLSX.Shared.Exceptions.IOException">Throws IOException in case of an error</exception>
        public static Workbook Load(Stream stream, ImportOptions options = null)
        {
            XlsxReader r = new XlsxReader(stream, options);
            r.Read();
            return r.GetWorkbook();
        }

        /// <summary>
        /// Loads a workbook from a file asynchronously
        /// </summary>
        /// <param name="filename">Filename of the workbook</param>
        /// <param name="options">Import options to override the data types of columns or cells. These options can be used to cope with wrong interpreted data, caused by irregular styles</param>
        /// <returns>Workbook object</returns>
        /// <exception cref="Exceptions.IOException">Throws IOException in case of an error</exception>
        public static async Task<Workbook> LoadAsync(string filename, ImportOptions options = null)
        {
            XlsxReader r = new XlsxReader(filename, options);
            await r.ReadAsync();
            return r.GetWorkbook();
        }

        /// <summary>
        /// Loads a workbook from a stream asynchronously
        /// </summary>
        /// <param name="stream">Stream containing the workbook</param>
        /// /// <param name="options">Import options to override the data types of columns or cells. These options can be used to cope with wrong interpreted data, caused by irregular styles</param>
        /// <returns>Workbook object</returns>
        /// <exception cref="NanoXLSX.Shared.Exceptions.IOException">Throws IOException in case of an error</exception>
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
