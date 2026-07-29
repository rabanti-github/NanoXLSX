/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Collections.Generic;
using System.Linq;
using NanoXLSX.Colors;
using NanoXLSX.Exceptions;
using NanoXLSX.Interfaces;
using NanoXLSX.Registry;
using NanoXLSX.Themes;
using NanoXLSX.Utils;

namespace NanoXLSX
{
    /// <summary>
    /// Class representing a workbook
    /// </summary>
    /// 
    public class Workbook
    {
        static Workbook()
        {
            PlugInLoader.Initialize();
        }

        #region privateFields
        private string filename;
        private List<Worksheet> worksheets;
        private Worksheet currentWorksheet;
        private Metadata workbookMetadata;
        private IPassword workbookProtectionPassword;
        private bool lockWindowsIfProtected;
        private bool lockStructureIfProtected;
        private int selectedWorksheet;
        private Shortener shortener;
        private readonly List<Color> mruColors = new List<Color>();
        private readonly List<DefinedName> definedNames = new List<DefinedName>();
        internal bool importInProgress; // Used by NanoXLSX.Reader
        #endregion

        #region properties

        /// <summary>
        /// Optional auxiliary data object. This object is used to store additional information about the workbook. 
        /// The data is not stored in the file but can be used by plug-ins
        /// </summary>
        internal AuxiliaryData AuxiliaryData { get; private set; }

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
        /// \remark <remarks>
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
        /// Password instance of the protected workbook. If a password was set, the pain text representation and the hash can be read from the instance
        /// </summary>
        /// \remark <remarks>The password of this property is stored in plain text at runtime but not stored to a workbook. The plain text password cannot be recovered when loading a workbook. The hash is retrieved and can be reused, 
        /// if no changes are made in the area of workbook protection (<see cref="SetWorkbookProtection(bool, bool, bool, string)"/>)</remarks>
        public virtual IPassword WorkbookProtectionPassword { get { return workbookProtectionPassword; } internal set => workbookProtectionPassword = value; }

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
        /// \remark <remarks>A hidden workbook can only be made visible, using another, already visible Excel window</remarks>
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

        #region methods

        /// <summary>
        /// Adds a color value (HEX; 6-digit RGB or 8-digit ARGB) to the MRU list
        /// </summary>
        /// <param name="color">RGB code in hex format (either 6 characters, e.g. FF00AC or 8 characters with leading alpha value). Alpha will be set to full opacity (FF) in case of 6 characters</param>
        public void AddMruColor(string color)
        {
            Validators.ValidateGenericColor(color);
            mruColors.Add(Color.CreateRgb(color));
        }

        /// <summary>
        /// Adds a generic color value. This can be an RGB/ARGB color, Auto, Theme, Indexed or System color
        /// </summary>
        /// <param name="color">Color instance</param>
        public void AddMruColor(Color color)
        {
            mruColors.Add(color);
        }

        /// <summary>
        /// Gets the MRU color list
        /// </summary>
        /// <returns>Immutable list of color instances</returns>
        public IReadOnlyList<Color> GetMruColors()
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
        /// Adds a defined name, pointing to a single cell, to the workbook
        /// </summary>
        /// <param name="name">Unique name of defined name</param>
        /// <param name="worksheet">Worksheet that contains the target cell (cannot be null)</param>
        /// <param name="cellAddress">Cell address as string (cannot be null)</param>
        /// <param name="localWorksheet">If not null, the defined name will only be available on the given worksheet. The combination of 'name' and 'worksheet' must be unique</param>
        /// <param name="comment">Optional comment of the defined name</param>
        /// <returns>Returns the added defined name object</returns>
        /// <exception cref="WorksheetException">Thrown if the worksheet or cell address is null</exception>
        /// <exception cref="FormatException">Thrown if the name of the defined name or the cell address is invalid</exception>
        public DefinedName AddDefinedNameCell(string name, Worksheet worksheet, string cellAddress, Worksheet localWorksheet = null, string comment = null)
        {
            Address address = new Address(cellAddress);
            return AddDefinedNameCell(name, worksheet, address, localWorksheet, comment);
        }

        /// <summary>
        /// Adds a defined name, pointing to a single cell, to the workbook
        /// </summary>
        /// <param name="name">Unique name of defined name</param>
        /// <param name="worksheet">Worksheet that contains the target cell (cannot be null)</param>
        /// <param name="column">Column number (zero-based) of the target cell</param>
        /// <param name="row">Row number (zero-based) of the target cell</param>
        /// <param name="localWorksheet">If not null, the defined name will only be available on the given worksheet. The combination of 'name' and 'worksheet' must be unique</param>
        /// <param name="comment">Optional comment of the defined name</param>
        /// <returns>Returns the added defined name object</returns>
        /// <exception cref="WorksheetException">Thrown if the worksheet or cell address is null</exception>
        /// <exception cref="FormatException">Thrown if the name of the defined name or the cell address is invalid</exception>
        public DefinedName AddDefinedNameCell(string name, Worksheet worksheet, int column, int row, Worksheet localWorksheet = null, string comment = null)
        {
            Address address = new Address(column, row);
            return AddDefinedNameCell(name, worksheet, address, localWorksheet, comment);
        }

        /// <summary>
        /// Adds a defined name, pointing to a single cell, to the workbook
        /// </summary>
        /// <param name="name">Unique name of defined name</param>
        /// <param name="worksheet">Worksheet that contains the target cell (cannot be null)</param>
        /// <param name="cellAddress">Address object of the target cell</param>
        /// <param name="localWorksheet">If not null, the defined name will only be available on the given worksheet. The combination of 'name' and 'worksheet' must be unique</param>
        /// <param name="comment">Optional comment of the defined name</param>
        /// <returns>Returns the added defined name object</returns>
        /// <exception cref="WorksheetException">Thrown if the worksheet or cell address is null</exception>
        /// <exception cref="FormatException">Thrown if the name of the defined name or the cell address is invalid</exception>
        public DefinedName AddDefinedNameCell(string name, Worksheet worksheet, Address cellAddress, Worksheet localWorksheet = null, string comment = null)
        {
            if (worksheet == null)
            {
                throw new WorksheetException("A defined name to a cell must have a worksheet");
            }
            if (cellAddress == null)
            {
                throw new WorksheetException("The cell address pointing to a defined name cannot be null");
            }
            return AddDefinedName(name, DefinedName.NameType.Cell, cellAddress, worksheet, localWorksheet, comment);
        }

        /// <summary>
        /// Adds a defined name, pointing to a cell range, to the workbook
        /// </summary>
        /// <param name="name">Unique name of defined name</param>
        /// <param name="worksheet">Worksheet that contains the target cell (cannot be null)</param>
        /// <param name="rangeAddress">Address of the target range as string</param>
        /// <param name="localWorksheet">If not null, the defined name will only be available on the given worksheet. The combination of 'name' and 'worksheet' must be unique</param>
        /// <param name="comment">Optional comment of the defined name</param>
        /// <returns>Returns the added defined name object</returns>
        /// <exception cref="WorksheetException">Thrown if the worksheet or range address is null</exception>
        /// <exception cref="FormatException">Thrown if the name of the defined name or the range address is invalid</exception>
        public DefinedName AddDefinedNameRange(string name, Worksheet worksheet, string rangeAddress, Worksheet localWorksheet = null, string comment = null)
        {
            Range range = new Range(rangeAddress);
            return AddDefinedNameRange(name, worksheet, range, localWorksheet, comment);
        }
        /// <summary>
        /// Adds a defined name, pointing to a cell range, to the workbook
        /// </summary>
        /// <param name="name">Unique name of defined name</param>
        /// <param name="worksheet">Worksheet that contains the target cell (cannot be null)</param>
        /// <param name="startAddress">Start address object of the target range</param>
        /// <param name="endAddress">End address object of the target range</param>
        /// <param name="localWorksheet">If not null, the defined name will only be available on the given worksheet. The combination of 'name' and 'worksheet' must be unique</param>
        /// <param name="comment">Optional comment of the defined name</param>
        /// <returns>Returns the added defined name object</returns>
        /// <exception cref="WorksheetException">Thrown if the worksheet or range address is null</exception>
        /// <exception cref="FormatException">Thrown if the name of the defined name or the range address is invalid</exception>
        public DefinedName AddDefinedNameRange(string name, Worksheet worksheet, Address startAddress, Address endAddress, Worksheet localWorksheet = null, string comment = null)
        {
            Range range = new Range(startAddress, endAddress);
            return AddDefinedNameRange(name, worksheet, range, localWorksheet, comment);
        }
        /// <summary>
        /// Adds a defined name, pointing to a cell range, to the workbook
        /// </summary>
        /// <param name="name">Unique name of defined name</param>
        /// <param name="worksheet">Worksheet that contains the target cell (cannot be null)</param>
        /// <param name="startColumn">Start column number (zero-based) of the target range</param>
        /// <param name="startRow">Start row number (zero-based) of the target range</param>
        /// <param name="endColumn">End column number (zero-based) of the target range</param>
        /// <param name="endRow">End row number (zero-based) of the target range</param>
        /// <param name="localWorksheet">If not null, the defined name will only be available on the given worksheet. The combination of 'name' and 'worksheet' must be unique</param>
        /// <param name="comment">Optional comment of the defined name</param>
        /// <returns>Returns the added defined name object</returns>
        /// <exception cref="WorksheetException">Thrown if the worksheet or range address is null</exception>
        /// <exception cref="FormatException">Thrown if the name of the defined name or the range address is invalid</exception>
        public DefinedName AddDefinedNameRange(string name, Worksheet worksheet, int startColumn, int startRow, int endColumn, int endRow, Worksheet localWorksheet = null, string comment = null)
        {
            Range range = new Range(startColumn, startRow, endColumn, endRow);
            return AddDefinedNameRange(name, worksheet, range, localWorksheet, comment);
        }
        /// <summary>
        /// Adds a defined name, pointing to a cell range, to the workbook
        /// </summary>
        /// <param name="name">Unique name of defined name</param>
        /// <param name="worksheet">Worksheet that contains the target cell (cannot be null)</param>
        /// <param name="rangeAddress">Range object of the target range</param>
        /// <param name="localWorksheet">If not null, the defined name will only be available on the given worksheet. The combination of 'name' and 'worksheet' must be unique</param>
        /// <param name="comment">Optional comment of the defined name</param>
        /// <returns>Returns the added defined name object</returns>
        /// <exception cref="WorksheetException">Thrown if the worksheet or range address is null</exception>
        /// <exception cref="FormatException">Thrown if the name of the defined name or the range address is invalid</exception>
        public DefinedName AddDefinedNameRange(string name, Worksheet worksheet, Range rangeAddress, Worksheet localWorksheet = null, string comment = null)
        {
            if (worksheet == null)
            {
                throw new WorksheetException("A defined name to a cell must have a worksheet");
            }
            if (rangeAddress == null)
            {
                throw new WorksheetException("The range address pointing to a defined name cannot be null");
            }
            return AddDefinedName(name, DefinedName.NameType.Range, rangeAddress, worksheet, localWorksheet, comment);
        }

        /// <summary>
        /// Adds a defined name, pointing to a constant value, to the workbook
        /// </summary>
        /// <param name="name">Unique name of defined name</param>
        /// <param name="value">Constant value</param>
        /// <param name="localWorksheet">If not null, the defined name will only be available on the given worksheet. The combination of 'name' and 'worksheet' must be unique</param>
        /// <param name="comment">Optional comment of the defined name</param>
        /// <returns>Returns the added defined name object</returns>
        /// <exception cref="WorksheetException">Thrown if the constant value is null</exception>
        /// <exception cref="FormatException">Thrown if the name of the defined name is invalid</exception>
        /// \remark <remarks>A constant should be one of the compatible types: string, int, uint, double, float, long, ulong, short, ushort, decimal, byte, sbyte, DateTime, TimeSpan, bool. Other types will be treated as string, using <see cref="object.ToString()"/>.</remarks>
        public DefinedName AddDefinedNameConstant(string name, object value, Worksheet localWorksheet = null, string comment = null)
        {
            if (value == null)
            {
                throw new WorksheetException("A constant value pointing to a defined name cannot be nul");
            }
            return AddDefinedName(name, DefinedName.NameType.Constant, value, null, localWorksheet, comment);
        }

        /// <summary>
        /// Adds a defined name, pointing to a formula expression, to the workbook. Do not add a leading equal sign (=) to the formula. 
        /// For simple cell references (e.g. "A1") use <see cref="AddDefinedNameCell(string, Worksheet, Address, Worksheet, string)"/> or one of the overloaded methods. 
        /// For cell range references (e.g. "A1:C3") use <see cref="AddDefinedNameRange(string, Worksheet, Address, Address, Worksheet, string)"/> or one of the overloaded methods. 
        /// </summary>
        /// <param name="name">Unique name of defined name</param>
        /// <param name="formula">Formula value</param>
        /// <param name="localWorksheet">If not null, the defined name will only be available on the given worksheet. The combination of 'name' and 'worksheet' must be unique</param>
        /// <param name="comment">Optional comment of the defined name</param>
        /// <returns>Returns the added defined name object</returns>
        /// <exception cref="WorksheetException">Thrown if the formula value is null or empty</exception>
        /// <exception cref="FormatException">Thrown if the name of the defined name is invalid</exception>
        /// \remark <remarks>The formula is not evaluated. Also do not add references to external workbooks (will cause an exception on save), as longs no NanoXLSX extension is loaded that can handle external links.</remarks>
        public DefinedName AddDefinedNameFormula(string name, string formula, Worksheet localWorksheet = null, string comment = null)
        {
            if (string.IsNullOrEmpty(formula))
            {
                throw new WorksheetException("A formula value pointing to a defined name cannot be null or empty");
            }
            // Note: Added strings like '[0]' will out-of-the-box throw an exception on saveing a workbook 
            return AddDefinedName(name, DefinedName.NameType.Formula, formula, null, localWorksheet, comment);
        }

        /// <summary>
        /// Internal method to add a defined name
        /// </summary>
        /// <param name="name">Name of the defined name.</param>
        /// <param name="type">Type of the reference.</param>
        /// <param name="value">Reference text (cell, range, formula, or constant).</param>
        /// <param name="targetWorksheet">Target worksheet in case of a cell or range. Pass null for constant or formula.</param>
        /// <param name="localWorksheet">Optional worksheet that scopes the defined name. Pass null for workbook scope.</param>
        /// <param name="comment">Optional comment.</param>
        /// <returns>Returns the added defined name.</returns>
        internal DefinedName AddDefinedName(string name, DefinedName.NameType type, object value, Worksheet targetWorksheet, Worksheet localWorksheet, string comment)
        {
            // Validation is implemented in DefinedName class
            DefinedName definedName = new DefinedName(this, type, name, value, targetWorksheet, localWorksheet, comment);
            AddDefinedName(definedName);
            return definedName;
        }

        /// <summary>
        /// Internal method to add a defined name object without validation
        /// </summary>
        /// <param name="definedName">Defined name object</param>
        internal void AddDefinedName(DefinedName definedName)
        {
            definedNames.Add(definedName);
        }

        /// <summary>
        /// Removes the defined name with the supplied name and scope.
        /// </summary>
        /// <param name="name">Name of the defined name to remove.</param>
        /// <param name="localSheet">Worksheet scope, or null for workbook scope.</param>
        /// <returns>True if a matching defined name was removed, false if no match was found.</returns>
        /// \remark <remarks><b>Important:</b> Removing defined names may remove references in worksheets. If a workbook is saved in this state, it can lead to broken formula references, indicated by "#NAME?"</remarks>
        public bool RemoveDefinedName(string name, Worksheet localSheet = null)
        {
            int index = FindDefinedNameIndex(name, localSheet);
            if (index < 0)
            {
                return false;
            }
            DefinedName definedName = definedNames[index];
            InvalidateDefinedNameReferences(definedName);
            definedNames.RemoveAt(index);
            return true;
        }

        /// <summary>
        /// Gets the defined name with the supplied name and scope.
        /// </summary>
        /// <param name="name">Name of the defined name.</param>
        /// <param name="localSheet">Worksheet scope, or null for workbook scope.</param>
        /// <returns>The matching <see cref="DefinedName"/> instance, or null if no match was found.</returns>
        public DefinedName GetDefinedName(string name, Worksheet localSheet = null)
        {
            int index = FindDefinedNameIndex(name, localSheet);
            return index < 0 ? null : definedNames[index];
        }

        /// <summary>
        /// Gets a read-only view of all defined names in this workbook in insertion order.
        /// </summary>
        /// <returns>Read-only list of <see cref="DefinedName"/> instances.</returns>
        public IReadOnlyList<DefinedName> GetDefinedNames()
        {
            return definedNames;
        }

        /// <summary>
        /// Locates the index of a defined name by name and scope.
        /// </summary>
        /// <param name="name">Name to find.</param>
        /// <param name="localSheet">Worksheet scope, or null for workbook scope.</param>
        /// <returns>Index in the internal list, or -1 if not found.</returns>
        internal int FindDefinedNameIndex(string name, Worksheet localSheet)
        {
            for (int i = 0; i < definedNames.Count; i++)
            {
                DefinedName candidate = definedNames[i];
                if (string.Equals(candidate.Name, name, System.StringComparison.Ordinal)
                    && ReferenceEquals(candidate.LocalSheet, localSheet))
                {
                    return i;
                }
            }
            return -1;
        }

        /// <summary>
        /// Adding a new Worksheet. The new worksheet will be defined as current worksheet
        /// </summary>
        /// <param name="name">Name of the new worksheet</param>
        /// <exception cref="WorksheetException">Throws a WorksheetException if the name of the worksheet already exists</exception>
        /// <exception cref="NanoXLSX.Exceptions.FormatException">Throws a FormatException if the name contains illegal characters or is out of range (length between 1 an 31 characters)</exception>
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
        /// <exception cref="NanoXLSX.Exceptions.FormatException">FormatException is thrown if the worksheet name contains illegal characters or is out of range (length between 1 an 31) and sanitizeSheetName is false</exception>
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
        /// <exception cref="NanoXLSX.Exceptions.FormatException">FormatException is thrown if the worksheet name contains illegal characters or is out of range (length between 1 an 31)</exception>
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
        /// <exception cref="NanoXLSX.Exceptions.FormatException">FormatException is thrown if the worksheet name contains illegal characters or is out of range (length between 1 an 31) and sanitation is false</exception>
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
        /// Method to resolve all merged cells in all worksheets. Only the value of the very first cell of the locked cells range will be visible. The other values are still present (set to EMPTY) but will not be stored in the worksheet.<br />
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
        /// \remark <remarks>This method does not set the current worksheet while design time. Use SetCurrentWorksheet instead for this</remarks>
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
        /// \remark <remarks>This method does not set the current worksheet while design time. Use SetCurrentWorksheet instead for this</remarks>
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
            workbookProtectionPassword.SetPassword(password);
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
        /// \remark <remarks>The copy is not set as current worksheet. The existing one is kept</remarks>
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
        /// \remark <remarks>The copy is not set as current worksheet. The existing one is kept</remarks>
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
        /// \remark <remarks>The copy is not set as current worksheet. The existing one is kept. The source worksheet can originate from any workbook</remarks>
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
        /// \remark <remarks>The copy is not set as current worksheet. The existing one is kept</remarks>
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
        /// \remark <remarks>The copy is not set as current worksheet. The existing one is kept</remarks>
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
        /// \remark <remarks>The copy is not set as current worksheet. The existing one is kept</remarks>
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
        /// Validates the worksheets regarding several conditions that must be met:<br />
        /// - At least one worksheet must be defined<br />
        /// - A hidden worksheet cannot be the selected one<br />
        /// - At least one worksheet must be visible<br />
        /// If one of the conditions is not met, an exception is thrown
        /// </summary>
        /// \remark <remarks>If an import is in progress, these rules are disabled to avoid conflicts by the order of loaded worksheets</remarks>
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
        /// Removes all references of invalid defined names in Worksheets
        /// </summary>
        /// <param name="definedName">Defined name object to invalidate</param>
        private void InvalidateDefinedNameReferences(DefinedName definedName)
        {
            foreach (Worksheet worksheet in worksheets)
            {
                foreach (KeyValuePair<string, Cell> cell in worksheet.Cells)
                {
                    FormulaData formula = cell.Value.Formula;
                    if (formula != null && ReferenceEquals(formula.DefinedNameReference, definedName))
                    {
                        formula.DefinedNameReference = null;
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
            workbookProtectionPassword = new LegacyPassword(LegacyPassword.PasswordType.WorkbookProtection);
            AuxiliaryData = new AuxiliaryData();
        }


        #endregion
    }
}
