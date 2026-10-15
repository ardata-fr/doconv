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
