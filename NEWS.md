# doconv 0.1.4

* internals: png filenames are now defined with a correct mask

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

