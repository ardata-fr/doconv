xlsx_file <- function() {
  system.file(package = "doconv", "doc-examples", "example.xlsx")
}

test_that("workbooks use the calc filter with one page per sheet", {
  expect_match(libreoffice_pdf_filter("a.XLSX"), "^pdf:calc_pdf_Export:.*SinglePageSheets")
  expect_match(libreoffice_pdf_filter("a.ods", show_markup = TRUE), "SinglePageSheets")
  expect_identical(guess_exec("xlsx"), c(xlsx = "excel"))
  expect_identical(guess_exec("xlsm"), c(xlsm = "excel"))
})

test_that("to_pdf() exports each sheet on a single page with LibreOffice", {
  skip_on_cran()
  skip_if_not(locatexec::exec_available("libreoffice"))
  # LibreOffice, even where Excel is available
  local_mocked_bindings(exec_available = function(exec, ...) exec == "libreoffice")

  pdf_file <- to_pdf(xlsx_file(), output = tempfile(fileext = ".pdf"))
  expect_identical(pdftools::pdf_info(pdf_file)$pages, 2L)
  text <- pdftools::pdf_text(pdf_file)
  expect_match(text[1], "column_15")
  expect_match(text[2], "apples")

  x <- to_miniature(xlsx_file())
  expect_s3_class(x, "magick-image")
})

test_that("xlsx2pdf() exports all sheets with Excel", {
  skip_on_cran()
  skip_if_not(locatexec::exec_available("excel"))

  pdf_file <- xlsx2pdf(xlsx_file(), output = tempfile(fileext = ".pdf"))
  text <- paste(pdftools::pdf_text(pdf_file), collapse = "\n")
  # the wide sheet fits the page width, and its 60 rows a single page
  expect_identical(pdftools::pdf_info(pdf_file)$pages, 2L)
  expect_match(pdftools::pdf_text(pdf_file)[1], "column_15")
  expect_match(text, "apples")

  x <- to_miniature(xlsx_file())
  expect_s3_class(x, "magick-image")
})
