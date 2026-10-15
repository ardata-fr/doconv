test_that("to_miniature() uses dpi with Word documents", {
  skip_on_cran()
  skip_if_not(locatexec::exec_available("libreoffice"))
  # convert with LibreOffice, so that the test does not need Word
  local_mocked_bindings(exec_available = function(...) FALSE)

  docx <- system.file(package = "doconv", "doc-examples/minimal-word.docx")
  width_50 <- magick::image_info(to_miniature(docx, dpi = 50))$width
  width_100 <- magick::image_info(to_miniature(docx, dpi = 100))$width
  expect_equal(width_100 / width_50, 2, tolerance = 0.05)
})

test_that("file extensions are not case sensitive", {
  pdf_file <- file.path(tempfile(), "doc.PDF")
  dir.create(dirname(pdf_file))
  grDevices::pdf(pdf_file, width = 2, height = 2)
  graphics::plot.new()
  graphics::plot.new()
  grDevices::dev.off()

  expect_length(doconv:::pdf_to_images(pdf_file, dpi = 30), 2L)
  expect_s3_class(to_miniature(pdf_file, dpi = 30), "magick-image")

  # default output must not be the input file
  default_output <- function(fun, input) {
    eval(formals(fun)$output, list(input = input))
  }
  expect_identical(default_output(docx2pdf, "doc.DOCX"), "doc.pdf")
  expect_identical(default_output(pptx2pdf, "doc.PPTX"), "doc.pdf")
})
