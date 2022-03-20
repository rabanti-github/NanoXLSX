# Change Log

## v2.11.4

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

Note: The column widths and row heights may change slightly with this release, since now the actual (internal) width and height is applied when setting a non-standard column width or row height

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
