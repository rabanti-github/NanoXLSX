# Change Log - NanoXLSX.Reader

## v3.2.1

---

Release Date: **(25.08.2026)** <sup>(DMY)</sup>

- Fixed a bug regarding path handling on reading files. Bug fix provided by yolkin-games
- Added tests
- Code maintenance


## v3.2.0

---

Release Date: **(20.08.2026)** <sup>(DMY)</sup>

- Added reading of defined names.
- Added reading of formula metadata and cached formula values, including cached error results.
- Added reading of standalone cell error values.
- Improved workbook-part and relationship discovery for more robust XLSX processing.
- Tolerant reader mode now skips invalid custom number formats with missing format codes;strict validation continues to reject them.
- Improved internal workbook finalization and formula/defined-name resolution.
- Added infrastructure for resolving external workbook references through compatibility plug-ins.
- Introduced a discovery reader to make the reading process more robust (collect data about all parts in a XLSX file)
- Updated readers with the new reader interfaces
- Deprecated RelationshipReader (replaced by DiscoveryReader)
- Added unit test for discovery
- Version bump of NanoXLSX.Core to v3.2.0
- Code maintenance

## v3.1.0

---
Release Date: **(04.05.2026)** <sup>(DMY)</sup>

- Version bump of NanoXLSX.Core to v3.1.0
- Improved reader performance (memory consumption, load time). Reading a workbook should now be up to 3 times faster.

## v3.0.1

---
Release Date: **(24.04.2026)** <sup>(DMY)</sup>

- Fixed internal async handling of the workbook reader, to avoid deadlocks in WinForms/WPF projects, when async is not used (regression).
- Fixed order of worksheets when manually changed
- Added filename to workbook when reading from file


## v3.0.0

---
Release Date: **(28.02.2026)** <sup>(DMY)</sup>

- Final release of NanoXLSX.Reader
- See the [main changelog](https://github.com/rabanti-github/NanoXLSX/blob/master/Changelog.md) for a comprehensive summary of all changes since v2.6.7

## v3.0.0-rc.4 + v3.0.0-rc.5

---
Release Date: **22.01.2026** <sup>(DMY)</sup>

- Added reader handling for the Font properties: `Font.Outline`, `Font.Shadow`, `Font.Condense` and `Font.Extend`
- Moved internal interfaces to NanoXLSX.Core (namespace `NanoXLSX.Interfaces.Reader`)
- Changed plug-in handling
- Version bump rc.4 to rc.5

## v3.0.0-rc.3

---
Release Date: **04.01.2026** <sup>(DMY)</sup>

- Changed handling of colors in the style reader (Fills) to consider:
  - sRGB colors (RGB / ARGB)
  - Indexed colors
  - Theme colors
  - System colors
  - Auto colors
  - Tint values

## v3.0.0-rc.2

---
Release Date: **27.11.2025** <sup>(DMY)</sup>

- Refactoring of several enums from lowercase start to uppercase start for better consistency

## v3.0.0-rc.1

---
Release Date: **25.11.2025** <sup>(DMY)</sup>

- Initial release of the reader library
