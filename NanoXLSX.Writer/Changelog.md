# Change Log - NanoXLSX.Writer

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
