test_that("show_markup is validated", {
  expect_error(to_pdf(tempfile(), show_markup = NA), "must be TRUE or FALSE")
  expect_error(docx2pdf(tempfile(), show_markup = "yes"), "must be TRUE or FALSE")
  expect_error(to_miniature(tempfile(), show_markup = c(TRUE, FALSE)), "must be TRUE or FALSE")
})

test_that("the LibreOffice filter draws comments only with show_markup", {
  expect_identical(libreoffice_pdf_filter("a.docx"), "pdf:writer_pdf_Export")
  expect_match(libreoffice_pdf_filter("a.DOCX", show_markup = TRUE), "ExportNotesInMargin")
  expect_identical(libreoffice_pdf_filter("a.pptx", show_markup = TRUE), "pdf:writer_pdf_Export")
})

test_that("docx_hide_revisions() asks for the final view of tracked changes", {
  skip_if_not_installed("officer")
  dir <- tempfile()
  dir.create(dir)

  docx <- markup_docx()
  hidden <- docx_hide_revisions(docx, dir = dir)
  expect_identical(normalizePath(dirname(hidden)), normalizePath(dir))
  settings <- utils::unzip(hidden, files = "word/settings.xml", exdir = tempfile())
  settings <- paste(readLines(settings, warn = FALSE), collapse = "\n")
  expect_match(settings, '<w:revisionView w:markup="false" w:insDel="false"', fixed = TRUE)

  # nothing to do without tracked changes, or with another format
  plain <- tempfile(fileext = ".docx")
  print(officer::read_docx(), target = plain)
  expect_identical(docx_hide_revisions(plain, dir = dir), plain)
  expect_identical(docx_hide_revisions("a.rtf", dir = dir), "a.rtf")
})

test_that("to_pdf() shows tracked changes and comments only with show_markup, with LibreOffice", {
  skip_on_cran()
  skip_if_not(locatexec::exec_available("libreoffice"))
  skip_if_not_installed("officer")
  # LibreOffice, even where Word is available
  local_mocked_bindings(exec_available = function(exec, ...) exec == "libreoffice")
  docx <- markup_docx()

  final <- pdf_text_of(to_pdf(docx, output = tempfile(fileext = ".pdf")))
  expect_match(final, "INSERTEDTEXT")
  expect_no_match(final, "DELETEDTEXT")
  expect_no_match(final, "THE COMMENT TEXT")

  markup <- pdf_text_of(to_pdf(docx, output = tempfile(fileext = ".pdf"), show_markup = TRUE))
  expect_match(markup, "INSERTEDTEXT")
  expect_match(markup, "DELETEDTEXT")
  expect_match(markup, "THE COMMENT TEXT")
})
