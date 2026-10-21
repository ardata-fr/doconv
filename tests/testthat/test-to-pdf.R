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

test_that("to_pdf() stops LibreOffice after timeout", {
  skip_on_cran()
  skip_if_not(locatexec::exec_available("libreoffice"))

  txt_file <- tempfile(fileext = ".txt")
  writeLines("hello", txt_file)

  expect_error(
    to_pdf(txt_file, output = tempfile(fileext = ".pdf"), timeout = 0.01),
    "did not finish within"
  )
  # 0 disables the timeout
  pdf_file <- to_pdf(txt_file, output = tempfile(fileext = ".pdf"), timeout = 0)
  expect_true(file.exists(pdf_file))
})

test_that("LibreOffice is run without R's LD_LIBRARY_PATH on Linux", {
  skip_if_not(locatexec::is_unix())
  old <- Sys.getenv("LD_LIBRARY_PATH", unset = NA)
  on.exit(if (is.na(old)) Sys.unsetenv("LD_LIBRARY_PATH") else Sys.setenv(LD_LIBRARY_PATH = old))
  Sys.setenv(LD_LIBRARY_PATH = "/r/lib:/usr/lib/x86_64-linux-gnu")

  res <- processx::run("sh", c("-c", "echo \"[$LD_LIBRARY_PATH]\""), env = libreoffice_env())
  expect_identical(trimws(res$stdout), "[]")
})

test_that("to_pdf() fails when LibreOffice exits without writing a PDF", {
  # LibreOffice can exit with status 0 after an export error
  local_mocked_bindings(
    exec_available = function(...) FALSE,
    libreoffice_exec = function(...) "soffice",
    run = function(...) {
      list(status = 0L, stdout = "", stderr = "Error: Please verify input parameters...",
           timeout = FALSE)
    }
  )
  txt_file <- tempfile(fileext = ".txt")
  writeLines("hello", txt_file)
  pdf_file <- tempfile(fileext = ".pdf")

  expect_error(
    to_pdf(txt_file, output = pdf_file),
    "did not produce a PDF. Error: Please verify input parameters"
  )
  expect_false(file.exists(pdf_file))
})
