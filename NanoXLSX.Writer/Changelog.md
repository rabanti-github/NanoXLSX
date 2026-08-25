# Change Log - NanoXLSX.Writer

## v3.2.1

---

Release Date: **(25.08.2026)** <sup>(DMY)</sup>

- Fixed a bug regarding path handling on writing files
- Added tests
- Code maintenance

## v3.2.0

---
Release Date: **(20.08.2026)** <sup>(DMY)</sup>

- Added writing of defined names.
- Added writing of extended formula metadata and cached formula values.
- Added writing of cell error values.
- Updated internal writer processing and package-part handling.
- Added infrastructure for writing external workbook references through compatibility plug-ins.
- Version bump of NanoXLSX.Core to v3.2.0
- Implemented writing of defined names (cell references)
- Code maintenance


## v3.1.0

---
Release Date: **(04.05.2026)** <sup>(DMY)</sup>

- Version bump of NanoXLSX.Core to v3.1.0
- Optimized writer performance (memory consumption, save time)
- Updated internal worksheet iteration to use `Worksheet.CellValues`, eliminating per-cell string allocation during save. Requires **NanoXLSX.Core ≥ 3.1.0**.



## v3.0.0

---
Release Date: **(28.02.2026)** <sup>(DMY)</sup>

- Final release of NanoXLSX.Writer
- See the [main changelog](https://github.com/rabanti-github/NanoXLSX/blob/master/Changelog.md) for a comprehensive summary of all changes since v2.6.7

## v3.0.0-rc.4 + v3.0.0-rc.5

---
Release Date: **22.01.2026** <sup>(DMY)</sup>

- Added writer handling for the Font properties: `Font.Outline`, `Font.Shadow`, `Font.Condense` and `Font.Extend`
- Moved internal interfaces to NanoXLSX.Core (namespace `NanoXLSX.Interfaces.Writer`)
- Version bump rc.4 to rc.5

## v3.0.0-rc.3

---
Release Date: **04.01.2026** <sup>(DMY)</sup>

- Changed handling of colors in the style writer (Fills) to consider:
  - sRGB colors (RGB / ARGB)
  - Indexed colors
  - Theme colors
  - System colors
  - Auto colors
  - Tint values
- Internal change of structured text handling

## v3.0.0-rc.2

---
Release Date: **27.11.2025** <sup>(DMY)</sup>

- Refactoring of several enums from lowercase start to uppercase start for better consistency

## v3.0.0-rc.1

---
Release Date: **25.11.2025** <sup>(DMY)</sup>

- Initial release of the writer library
