% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/docx2pdf.R
\name{docx_update}
\alias{docx_update}
\title{Update docx fields}
\usage{
docx_update(input)
}
\arguments{
\item{input}{file input}
}
\value{
the name of the produced pdf (the same value as \code{output})
}
\description{
Update all fields and table of contents of a
Word document using "Microsoft Word".
This function will not work if "Microsoft Word" is not available
on your machine.

The calls to "Microsoft Word" are made differently depending on the
operating system. On "Windows", a "PowerShell" script using COM
technology is used to control "Microsoft Word". On macOS, an "AppleScript"
script is used to control "Microsoft Word".
}
\examples{
library(locatexec)
if (exec_available('word')) {
  file <- system.file(package = "doconv",
    "doc-examples/example.docx")
  docx_out <- tempfile(fileext = ".docx")
  file.copy(file, docx_out)
  docx_update(input = docx_out)
}
}
