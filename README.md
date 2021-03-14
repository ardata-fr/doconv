
<!-- README.md is generated from README.Rmd. Please edit that file -->

# doconv

The tool offers a set of functions for converting ‘Microsoft Word’ or
‘Microsoft PowerPoint’ documents to ‘PDF’ format and also for
converting them to images in the form of thumbnails. In order to work,
‘LibreOffice’ must be installed on the machine and possibly ‘python’
and ‘Microsoft Word’.

<!-- badges: start -->

[![R build
status](https://github.com/ardata-fr/doconv/workflows/R-CMD-check/badge.svg)](https://github.com/ardata-fr/doconv/actions)
[![Lifecycle:
maturing](https://img.shields.io/badge/lifecycle-experimental-orange.svg)](https://www.tidyverse.org/lifecycle/#experimental)
<!-- badges: end -->

## Installation

You can install the latest version from GitHub with:

``` r
# install.packages("devtools")
devtools::install_github("ardata-fr/doconv")
```

## Example

``` r
library(doconv)
```

## Generate thumbails from file

You can generate thumbails as an image by using `to_miniature`:

``` r
docx_file <- system.file(package = "doconv", "doc-examples/bookdown.docx")
to_miniature(
  filename = docx_file, 
  row = c(1, 1, 1, 2, 2, 2),
  use_docx2pdf = TRUE)
```

<img src="man/figures/README-unnamed-chunk-3-1.png" width="100%" />

It uses ‘LibreOffice’ to convert Word or PowerPoint documents to PDF. It
probably works with other types of document. If option
`use_docx2pdf=TRUE`, [docx2pdf](https://github.com/AlJohri/docx2pdf) is
used instead of ‘LibreOffice’ to convert Word files to PDF; you can only
use that option if ‘Word’ and ‘docx2pdf’ is installed on your machine.

## Convert a PowerPoint file to PDF

``` r
docx_file <- system.file(package = "doconv", "doc-examples/first_example.pptx")
to_pdf(docx_file, output = "first_example.pdf")
#> [1] "first_example.pdf"
to_miniature("first_example.pdf", width = 750)
```

<img src="man/figures/README-unnamed-chunk-4-1.png" width="100%" />

### Convert a Word file to PDF

``` r
to_pdf(docx_file, output = "bookdown.pdf")
#> [1] "bookdown.pdf"
```

## Related work

  - Packages [docxtractr](http://cran.r-project.org/package=docxtractr)
    is providing `convert_to_pdf()` that works very well.
