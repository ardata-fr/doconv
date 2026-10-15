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
