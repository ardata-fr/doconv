docx <- system.file(package = "doconv", "doc-examples/example.docx")
fake_run <- function(status) {
  function(...) list(status = status, stdout = "", stderr = "", timeout = FALSE)
}

test_that("docx_update() on macOS reports a failure with an explicit error", {
  local_mocked_bindings(is_osx = function() TRUE, run = fake_run(1L))
  expect_error(doconv:::docx_update_osx(docx), "could not update")
})

test_that("docx_update() does not create the working directory", {
  wd <- tempfile()
  local_mocked_bindings(working_directory = function() wd)

  local_mocked_bindings(is_osx = function() TRUE, run = fake_run(0L))
  expect_true(doconv:::docx_update_osx(docx))
  expect_false(dir.exists(wd))

  local_mocked_bindings(is_windows = function() TRUE)
  expect_true(doconv:::docx_update_win(docx))
  expect_false(dir.exists(wd))
})
