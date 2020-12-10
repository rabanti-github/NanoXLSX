# Change Log

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
