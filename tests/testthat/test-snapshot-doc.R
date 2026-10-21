test_that("expect_snapshot_doc() skips when Word is not available", {
  local_mocked_bindings(exec_available = function(...) FALSE)
  docx <- system.file(package = "doconv", "doc-examples/example.docx")
  expect_condition(
    expect_snapshot_doc(name = "docx-without-word", x = docx, engine = "testthat"),
    "'word' must be available",
    class = "skip"
  )
})

test_that("expect_snapshot_doc() recognises .doc files", {
  local_mocked_bindings(exec_available = function(...) FALSE)
  doc_file <- tempfile(fileext = ".doc")

  # with testthat, the test is skipped as Word is not available
  expect_condition(
    expect_snapshot_doc(name = "doc-without-word", x = doc_file, engine = "testthat"),
    "'word' must be available",
    class = "skip"
  )

  # with tinytest, the result reports that Word is required
  res <- doconv:::expect_snapshot_tinytest(current = doc_file, name = "doc-without-word")
  expect_match(attr(res, "info"), "'word' must be available")
})

test_that("diff_image() returns the ratio of different pixels", {
  write_img <- function(arr) {
    path <- tempfile(fileext = ".png")
    magick::image_write(magick::image_read(arr), path, format = "png")
    path
  }
  red <- array(0, c(10, 10, 4))
  red[, , 1] <- 1
  red[, , 4] <- 1
  ref <- write_img(red)

  # identical files
  expect_identical(doconv:::diff_image(ref, ref), 0)

  # same pixels, different files: rgb versus opaque rgba
  expect_equal(doconv:::diff_image(ref, write_img(red[, , 1:3])), 0)

  # one pixel with another color, one pixel with another alpha
  changed <- red
  changed[1, 1, 2] <- 1 / 255
  changed[2, 2, 4] <- 0.5
  expect_equal(doconv:::diff_image(ref, write_img(changed)), 0.02)

  # unequal dimensions
  expect_identical(doconv:::diff_image(ref, write_img(red[1:5, , ])), -1)
})
