# Change Log

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

