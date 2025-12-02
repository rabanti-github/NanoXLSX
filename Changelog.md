# Change Log

## v3.0.0-rc.3

---
Release Date: **02.12.2025**

- Formal changes to enforce displaying target frameworks in NuGet meta package

Note: The version numbers of the dependencies `NanoXLSX.Core`, `NanoXLSX.Reader` and `NanoXLSX.Writer` have not changed with this release. There are also no functional changes



## v3.0.0-rc.2

---
Release Date: **27.11.2025**

- Refactoring of several enums in `NanoXLSX.Core`, `NanoXLSX.Reader` and `NanoXLSX.Writer` from lowercase start to uppercase start for better consistency


# v3.0.0-rc.1

---
Release Date: **25.11.2025**

- Initial release candidate of the NanoXLSX library split into three separate libraries:
  - NanoXLSX.Core
  - NanoXLSX.Reader
  - NanoXLSX.Writer

## v2.6.7

---
Release Date: **01.10.2025**

- Fixed handling of worksheet protection (regression bug)
- Code cleanup

## v2.6.6

---
Release Date: **29.09.2025**

- Fixed handling of worksheet protection (selecting locked or unlocked cells)
- Added test case

Note: The default value of `Style.CurrentCellXf.Locked` is now true, to be consistent with Excel behavior. This change only affects worksheets with protection enabled and may require
  explicit unlocking of cells that should remain editable

## v2.6.5

---
Release Date: **13.09.2025**

- Added import option to ignore invalid column widths or row heights. Concept provided by pokorny
- Added test case

## v2.6.4

---
Release Date: **19.07.2025**

- Added support for in-line string values (non-formatted). Change provided by Misir
- Added test case

## v2.6.3

---
Release Date: **26.04.2025**

- Fixed a bug that prevented adding new worksheets when a pane split was defined
- Changed handling of reading workbooks, when docProps are missing (formal change)
- Added test case

## v2.6.2

---
Release Date: **24.01.2025**

- Fixed a regression bug in the Cell function ConvertArray
- Added test cases

## v2.6.1

---
Release Date: **19.01.2025**

- Fixed a bug on writing default column styles (not persisted in some cases)
- Adapted style reader: When a workbook is loaded, not defined color values of Border styles are now empty strings (were null), as if a new style is created
- Code maintenance

Note: The color values of Border styles are handled identical on writing XLSX files, either if null or empty. The change of the reader behavior was to enforce the "What You Can Write Is What You Can Read" policy of the library (writing an empty string as color value should lead to an empty string on read).

## v2.6.0

---
Release Date: **12.01.2025**

- Added InsertRow and InsertColumn functions. Functionality provided by Alexander Schlecht
- Added FirstCellByValue, FirstOrDefaultCell, CellsByValue functions. Functionality provided by Alexander Schlecht
- Added ReplaceCellValue function. Functionality provided by Alexander Schlecht
- Code maintenance

## v2.5.2

---
Release Date: **24.11.2024**

- Fixed a bug of the column address (letter) resolution. Column letters above 'Z' were resolved incorrectly
- Changed async handing of the workbook reader, to avoid deadlocks. Change provided by Jarren Long
- Simplified project structure (unified .Net 4.x and Standard). Change provided by Jarren Long
- Added tests for column address resolution

## v2.5.1

---
Release Date: **26.10.2024**

- Fixed a bug regarding the determination of the first data cell in an empty worksheet. Bug fix provided by Martin Stránský

## v2.5.0

---
Release Date: **22.07.2024**

- Adapted handling of the font scheme in styles. The scheme is now determined automatically
- Added column option to define a default column style
- Added tests

## v2.4.0

---
Release Date: **21.04.2024**

- Added handling to load workbooks from files asynchronously. Concept provided by John Leyva
- Fixed a bug when loading a workbook asynchronously from a stream. Bug fix provided by John Leyva
- Fixed a bug when the column auto-filter is a single cell address. Bug fix provided by pokorny
- Fixed a bug regarding style enumeration when reading a workbook. Bug fix provided by Martin Stránský
- Added new  and adapted existing test cases

## v2.3.3

---
Release Date: **24.02.2024**

- Fixed a bug in the GetFirstDataCellAddress function
- Fixed test cases

## v2.3.2

---
Release Date: **24.02.2024**

- Fixed a bug when reading min and max values in the GetLastDataColumnNumber function. Bug fix provided by pokorny
- Code maintenance

## v2.3.1

---
Release Date: **22.01.2024**

- Fixed a bug when reading fill styles. Bug fix provided by Marq Watkin
- Fixed a bug regarding casting floats to integers, in the worksheet reader. Bug fix provided by wappenull
- Removed broken debug code in tests
- Code maintenance

## v2.3.0

---
Release Date: **07.09.2023**

- Added worksheet option for zoom factors
- Added worksheet option for view types (e.g. page break preview)
- Added worksheet option to show or hide grid lines
- Added worksheet option to show or hide columns and row headers
- Added worksheet option to show or hide rulers in page layout view type

## v2.2.0

---
Release Date: **23.04.2023**

- Added new import option to cast all single values into decimals. Feature implementation provided by Tim M. Madsen
- Adapted hex color validation (clarified number of necessary characters)
- Internal changes of build processes (Documentation generation is performed now by a GitHub Action)

## v2.1.1

---
Release Date: **04.03.2023**

- Fixed a bug when a workbook contains charts instead of worksheets. Bug fix provided by Iivari Mokelainen
- Minor code maintenance

## v2.1.0

---
Release Date: **08.11.2022**

- Added a several methods in the Worksheet class to add multiple ranges of selected cells
- Fixed a bug in the reader function to read worksheets with multiple ranges of selected cells
- Fixed a bug in several readers to cope (internally) with bools, represented by numbers and textual expressions
- Removed internal escaping of custom number format codes for now
- Updated example in demo

Note: It seems that newer versions of Excel may store boolean attributes internally now as texts (true/false) and not anymore as numbers (1/0).
      This release adds compatibility to read this newer format but will currently store files still in the old format

Note 2: The incomplete internal escaping of custom number format codes was removed due to the potential high complexity.
        Escaping must be performed currently by hand, according to OOXML specs: Part 1 - Fundamentals And Markup Language Reference, Chapter 18.8.31

## v2.0.4

---
Release Date: **04.10.2022**

- Fixed a bug in the reader function of dates and times on hosts with locales different than en-US (and others)
- Added == and != operator overload on Address and Range struct
- Code maintenance

## v2.0.3

---
Release Date: **01.10.2022**

- Fixed a bug in the functions to write and read font values (styles)
- Adapted tests according to specs
- Updated documentation
- Added some internal notes to prepare the development of the next mayor version

## v2.0.2

---
Release Date: **29.09.2022**

- Fixed a bug in the functions to write and read custom number formats
- Fixed behavior of empty cells and added re-evaluation if values are set by the Value property
- Adapted and added further tests
- Removed several obsolete files and fixed project links

Note:

- When defining a custom number format, now the CustomFormatCode property must always be defined as well, since an empty value leads to an invalid Workbook 
- When a cell is now created (by constructor) with the type EMPTY, any passed value will be discarded in this cell

## v2.0.1

---
Release Date: **10.09.2022**

- Fixed a bug when loading workbooks on hosts with locales different than en-US (and others)

## v2.0.0

---
Release Date: **03.09.2022 - Major Release**

### Workbook and Shortener

- Added a list of MRU colors that can be defined in the Workbook class (methods AddMruColor, ClearMruColors)
- Added an exposed property for the workbook protection password hash (will be filled when loading a workbook)
- Added the method SetSelectedWorksheet by name in the Workbook class
- Added two methods GetWorksheet by name or index in the Workbook class
- Added the methods CopyWorksheetIntoThis and CopyWorksheetTo with several overloads in the Workbook class
- Added the method RemoveWorksheet by index with the option of resetting the current worksheet, in the Workbook class
- Added the method SetCurrentWorksheet by index in the Workbook class
- Added the method SetSelectedWorksheet by name in the Workbook class
- Added a Shortener-Class constructor with a workbook reference
- The shortener methods Down and Right have now an option to keep row and column positions
- Added two shortener methods Up and Left
- Made several non-functional style assigning methods deprecated in the Workbook class (will be removed in future versions)

### Worksheet

- Added an exposed property for the worksheet protection password hash (will be filled when loading a workbook)
- Added the methods GetFirstDataColumnNumber, GetFirstDataColumnNumber, GetFirstDataRowNumber, GetFirstRowNumber, GetLastDataColumnNumber, GetFirstCellAddress, GetFirstDataCellAddress, GetLastDataColumnNumber, GetLastDataRowNumber, GetLastRowNumber, GetLastCellAddress,  GetLastCellAddress and GetLastDataCellAddress
- Added the methods GetRow and GetColumns by address string or index
- Added the method Copy to copy a worksheet (deep copy)
- Added a constructor with only the worksheet name as parameter
- Added and option in GoToNextColumn and GoToNextRow to either keep the current row or column
- Added the methods RemoveRowHeight and RemoveAllowedActionOnSheetProtection
- Renamed columnAddress and rowAddress to columnNumber and rowNumber in the AddNextCell, AddCellFormula and RemoveCell methods
- Added several validations for worksheet data

### Cells, Rows and Columns

- In a Cell object, the address can now have reference modifiers ($)
- The worksheet reference in the Cell constructor was removed. Assigning to a worksheet is now managed automatically by the worksheet when adding a cell
- Added a property CellAddressType in the Cell class
- Cells can now have null as value, interpreted as empty
- Added a new overloaded method ResolveCellCoordinate to resolve the address type as well
- Added the methods ValidateColumnNumber and ValidateRowNumber in the Cell class
- In Address, the constructor with string and address type now only needs a string, since reference modifiers ($) are resolved automatically
- Address objects are now comparable
- Implemented better address validation
- Range start and end addresses are swapped automatically, if reversed

### Styles

- Font class has now an enum of possible underline values (e.g. instead of a bool)
- CellXf class supports now indentation
- A new, internal style repository was introduced to streamline the style management
- Color (RGB) values are now validated (Fill class has a method ValidateColor)
- Style components have now more appropriate default values
- MRU colors are now not collected from defined style colors anymore, but from the MRU list in the workbook object
- The ToString function of styles and all sub parts will now give a complete outline of all elements
- Fixed several issues with style comparison
- Several style default values were introduced as constants

### Formulas

- Added uint as possible formula value. Valid types are int, uint, long, ulong, float, double, byte, sbyte, decimal, short and ushort
- Added several validity checks

### Reader

- Added default values for dates, times and culture info in the import options
- Added global casting import options: AllNumbersToDouble, AllNumbersToDecimal, AllNumbersToInt, EverythingToString
- Added column casting import options: Double, Decimal
- Added global import options: EnforcePhoneticCharacterImport, EnforceEmptyValuesAsString, DateTimeFormat, TemporalCultureInfo
- Added a meta data reader for workbook meta data
- All style elements that can be written can also be read
- All workbook elements that can be written can also be read (exception: passwords cannot be recovered)
- All worksheet elements that can be written can also be read (exception: passwords cannot be recovered)
- Better handling of dates and times, especially with invalid (too low and too high numbers) values

### Misc

- Added a unit test project with several thousand, partially parametrized test cases
- Added several constants for boundary dates in the Utils class
- Added several methods for pane splitting in the Utils class
- Exposed the (legacy) password generation method in the Utils class
- Updated documentation among the whole project
- Exceptions have no sub titles anymore
- Overhauled the whole writer
- Removed lot of dead code for better maintenance

## v1.8.7

---
Release Date: **06.08.2022**

- Fixed a bug when setting a workbook protection password

## v1.8.6

---
Release Date: **02.04.2022**

- Added an import option to display phonetic characters (like Ruby Characters / Furigana / Zhuyin / Pinyin are now discarded) in strings

Note: Phonetic characters are discarded by default. If the import option "EnforcePhoneticCharacterImport" is set to true, the phonetic transcription will be displayed in brackets, right after the characters to be transcribed

## v1.8.5

---
Release Date: **27.03.2022**

- Fixed a follow-up issue on finding first/last cell addresses on explicitly defined, empty cells
- Code maintenance

## v1.8.4

---
Release Date: **20.03.2022**

- Fixed a regression bug, caused by changes of v1.8.3

## v1.8.3

---
Release Date: **10.03.2022**

- Added functions to determine the first cell address, column number or row number of a worksheet
- Adapted internal style handling
- Adapted the internal building of XML documents
- Fixed a bug in the handling of border colors

## v1.8.2

---
Release Date: **20.12.2021**

- Added hidden property for worksheets when loading a workbook

Note: The reader functionality on worksheets is not feature complete yet. Additional information like panes, splitting, column and row sizes are currently in development

## v1.8.1

---
Release Date: **12.09.2021**

- Fixed a bug when hiding worksheets

Note: It is not possible anymore to remove all worksheets from a workbook, or to set a hidden one as active. This would lead to an invalid Excel file

## v1.8.0

---
Release Date: **10.07.2021**

- Added functions to split (and freeze) a worksheet horizontally and vertically into panes
- Added a property to set the visibility of a workbook
- Added a property to set the visibility of worksheets
- Added two examples in the demo for the introduced split, freeze and visibility functionalities
- Added the possibility to define column widths and row height even if there are no cells defined
- Fixed the internal representation of column widths and row heights
- Minor code maintenance

Note: The column widths and row heights may change slightly with this release, since the actual (internal) width and height is now applied when setting a non-standard column width or row height

## v1.7.0

---
Release Date: **05.06.2021**

- Added functions to determine the last row, column or cell with data
- Fixed documentation formatting issues
- Updated readme and documentation

## v1.6.0

---
Release Date: **18.04.2021**

- Introduced library version for .NET Standard 2.0 (and assigned demos)
- Updated project structure (two projects for .NET >=4.5 and two for .NET Standard 2.0)
- Added function SetStyle in the Worksheet class
- Added demo for the new SetStyle function
- Changed behavior of empty cells. They are now not string but implicit numeric cells
- Added new function ResolveEnclosedAddresses in Range class
- Added new function GetAddressScope in Cell class
- Fixed the validation of cell addresses (single cell)
- Defined several immutable lists as return values to IReadOnlyList
- Minor code maintenance

Thanks to the following people (in the order of contribution date):

- Shobb for the introduction of IReadOnlyList
- John Lenz for the port to .NET Standard
- Ned Marinov for the proposal of the new SetStyle function

## v1.5.0

---
Release Date: **10.12.2020**

- Added indentation property of horizontal text alignment (CellXF) as style 
- Added example in demo for text indentation
- Code Cleanup

## v1.4.1

---
Release Date: **13.09.2020**

- Fixed a bug regarding numeric cells in the worksheet reader. Bug fix provided by John Lenz
- Minor code maintenance
- Updated readme and documentation

## v1.4.0

---
Release Date: **30.08.2020**

- Added style reader to resolve dates and times properly
- Added new data type TIME, represented by TimeSpan objects in reader and writer
- Changed namespace from 'Styles' to 'NanoXLSX.Styles'
- Added time (TimeSpan) examples to the demos
- Added a check to ensure dates are not set beyond 9999-12-31 (limitation of OAdate)
- Updated documentation
- Fixed some code formatting issues

### Notes

- To be consistent, the namespace of 'Styles' was changed to 'NanoXLSX.Styles'. Minor changes may be necessary in existing code if styles were used
- Currently, the style reader resolves only number formats to determine dates and times, as well as custom formats. Other components like fonts, borders or fills are neglected at the moment

## v1.3.6

---
Release Date: **19.07.2020**

- Fixed a bug in the reader regarding dates, times and booleans
- Fixed a bug in the method AddNextCellFormula

Note: Fixes provided by Silvio Burger and Thiago Souza. The fix for the reader bug is currently a work-around

## v1.3.5

---
Release Date: **10.01.2020**

- Fixed a bug in the reader regarding decimal numbers (for locales where the decimal pointer is not a dot)
- Formal changes

## v1.3.4

---
Release Date: **01.12.2019**

- Fixed a bug of reorganized worksheets (when deleted in Excel)
- Fixed a bug in the handling of shared strings
- Minor code maintenance

## v1.3.3

---
Release Date: **20.05.2019**

- Fixed a bug in the handling of streams (streams can be left open now)
- Updated stream demo
- Code Cleanup
- Removed executable folder, since executables are available through releases, compilation or NuGet

## v1.3.2

---
Release Date: **08.12.2018**

- Improved the performance of adding stylized cells by factor 10 to 100

## v1.3.1

---
Release Date: **04.11.2018**

- Fixed a bug in the style handling of merged cells. Bug fix provided by David Courtel for PicoXLSX

## v1.3.0

---
Release Date: **06.10.2018**

- Added missing features of PicoXLSX (synced with PicoXLSX version 2.6.1)
- Added asynchronous methods SaveAsync, SaveAsAsync, SaveAsStreamAsync and LoadAsync
- Added a new example for the introduced async methods
- Renamed namespace Exception to Exceptions
- Renamed namespace Style to Styles
- Fixed a bug regarding formulas in the reader
- Added support for dates in the reader
- Documentation Update
- Removed redundant code

## v1.2.4

---
Release Date: **24.08.2018**

- Fixed a bug regarding formulas in the reader
- Added support for dates in the reader
- Documentation Update

## v1.2.3

---
Release Date: **24.08.2018**

- Initial Release (synced to v 1.2.3 of NanoXLSX4j for Java)
