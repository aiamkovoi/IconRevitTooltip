# IconRevitTooltip

Shows Revit file version, build, worksharing, locale and save count on hover in Windows Explorer. Works with `.rvt` and `.rfa` files, including older versions (2016+). Revit installation is not required.

**Locale** (ENG/RUS/etc.) shows the language of Revit used when the file was last saved. Useful because system family and parameter names depend on the language version.

![screenshot](image.png)

## Install

1. Right-click `install.bat` → **Run as Administrator**
2. Explorer restarts automatically  
3. Hover over any `.rvt` or `.rfa` file

Pre-built binaries are included in `dist/`. No Visual Studio needed.

To build from source, delete `dist/` and re-run `install.bat` — MSBuild will be detected automatically.

## Uninstall

Run `uninstall.bat` as Administrator.

## ⚠️ Note

Revit 2024+ registers its own tooltip handler. This extension overrides it during install (originals are backed up and restored on uninstall). After updating Revit you may need to re-run `install.bat`.

## Requirements

.NET Framework 4.7 (comes with Revit).

## License

MIT

Based on [ShowRevitVersion](https://github.com/Tereami/ShowRevitVersion) by Tereami.
