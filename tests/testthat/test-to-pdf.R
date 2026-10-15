test_that("to_pdf() and check_libreoffice_export() clean their temporary files", {
  skip_on_cran()
  skip_if_not(locatexec::exec_available("libreoffice"))

  txt_file <- tempfile(fileext = ".txt")
  writeLines("hello", txt_file)
  pdf_file <- tempfile(fileext = ".pdf")

  before <- list.files(tempdir(), all.files = TRUE, no.. = TRUE)
  to_pdf(txt_file, output = pdf_file)
  check_libreoffice_export()
  after <- list.files(tempdir(), all.files = TRUE, no.. = TRUE)

  expect_true(file.exists(pdf_file))
  expect_setequal(setdiff(after, before), basename(pdf_file))
})
