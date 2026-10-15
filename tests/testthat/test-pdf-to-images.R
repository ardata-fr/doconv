test_that("pdf_to_images() keeps the page order", {
  pdf_file <- tempfile(fileext = ".pdf")
  grDevices::pdf(pdf_file, width = 2, height = 2)
  for (i in 1:12) {
    graphics::plot.new()
    graphics::text(.5, .5, i, cex = 4)
  }
  grDevices::dev.off()

  img_list <- doconv:::pdf_to_images(pdf_file, dpi = 30)
  expect_length(img_list, 12L)

  # each page rendered alone, in order
  pages <- vapply(seq_len(12), function(i) {
    png_file <- pdftools::pdf_convert(
      pdf_file, format = "png", pages = i, dpi = 30, verbose = FALSE,
      filenames = paste0(tempfile(), "_%d.%s")
    )
    paste(as.integer(magick::image_data(magick::image_read(png_file))), collapse = "")
  }, character(1))
  images <- vapply(img_list, function(img) {
    paste(as.integer(magick::image_data(img)), collapse = "")
  }, character(1))
  expect_identical(match(images, pages), 1:12)
})
