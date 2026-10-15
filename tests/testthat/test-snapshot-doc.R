test_that("expect_snapshot_doc() skips when Word is not available", {
  local_mocked_bindings(exec_available = function(...) FALSE)
  docx <- system.file(package = "doconv", "doc-examples/example.docx")
  expect_condition(
    expect_snapshot_doc(name = "docx-without-word", x = docx, engine = "testthat"),
    "'word' must be available",
    class = "skip"
  )
})
