# doconv 0.4.0

## Features

- New exported function `run_in_shiny_app()` to run an arbitrary
  command against a live Shiny application. Manages the app lifecycle
  (start, wait for HTTP readiness, stop) via processx; the command
  to execute is passed through `...` to `processx::run()`.
- New exported functions `expect_snapshot_ggplots()` and
  `expect_snapshot_flextables()` for visual regression testing of
  ggplot2 plots and flextable objects within testthat test suites.
- New exported helper `skip_if_not_snapshot_png()` to skip tests
  when snapshot dependencies (ragg, gdtools) are not available.
- `to_miniature()` gains `ncol` and `ncol_landscape` parameters for
  easier page layout. `ncol` groups pages N-per-row; adding
  `ncol_landscape` enables orientation-aware layout where portrait and
  landscape pages use different column counts.
- `docx2pdf()` and `to_miniature()` gain a `show_markup` parameter.
  When `TRUE`, tracked changes and comments are rendered visibly in the
  PDF output on Windows. Requires Microsoft Word.

## Issues

- Forward slashes in file paths are now converted to backslashes before
  injection into PowerShell scripts, fixing failures when paths contain
  spaces (#5).
- PowerShell scripts now use `try/finally` to guarantee `Word.Quit()` and
  `PowerPoint.Quit()` are called even when an error occurs, preventing
  orphan COM processes on Windows.
- `Word.Visible = $False` is now set before opening the document in
  PowerShell scripts, eliminating the brief window flash on startup.
- PowerPoint's `Presentations.Open()` now uses `WithWindow=$false` to
  open presentations without a visible window.
- AppleScript scripts now use `try/on error` blocks to ensure cleanup
  (close document + conditional quit) runs even when an error occurs,
  and the error is re-raised so that `processx::run()` sees a non-zero
  exit status.
- File paths are now sanitized before injection into script templates:
  `sprintf()` replaced by `sub()` to handle `%` in paths, and special
  characters (`"`, `` ` ``, `$`, `\`) are escaped for PowerShell and
  AppleScript respectively.

# doconv 0.3.3

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

