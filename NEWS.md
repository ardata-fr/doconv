# doconv 0.3.3 (dev version)

## Issues

- Fail with informative error message if PowerShell (PS) execution strategy does not allow
  running PS scripts. PS scripts are required for certain actions on Windows (#2).

# doconv 0.3.2

## Features

- add support for 'RTF'.

## Issues

- For 'Windows' users, figure resolution should now remain the same
when exporting 'Word' to 'PDF'.

# doconv 0.3.1

## Issues

- dont load tinytest
- drop officer from suggests

# doconv 0.3.0

## Features

* new function `expect_snapshot_html()` for visual testing HTML documents.
* new function `msoffice_available()` to test if 'Word' and 'PowerPoint'
are available.

## Issues

* fix tinytest registration


# doconv 0.2.0

## Features

* new fonction `expect_snapshot_doc()` for visual testing.

# doconv 0.1.4

## Features

* new fonction `docx_update()` to refresh all TOC and fields.
* new `dpi` parameters for image resolution

## Changes

* internals: png filenames are now defined with a correct mask
* [breaking change]: python, docx2pdf are not required anymore. 
* `tools::R_user_dir()` is used instead of package 'rappdirs'.

# doconv 0.1.3

## Issues

* Export functions `init_working_directory`, `rm_working_directory` and `working_directory` 
to let users manage docx2pdf working directory and comply with CRAN policy.
* Use `tempfile()` to make libreoffice write in a temporary directory instead of 
using and managing `working_directory()`.
* fix for `working_directory()` so that when deleted, no empty directory is left.

# doconv 0.1.2

* Add argument `UserInstallation` to function `to_pdf()`

# doconv 0.1.1

* Added `check_libreoffice_export()` that checks 'LibreOffice' can export to PDF.

