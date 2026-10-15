test_that("Office calls from several R processes run one at a time", {
  skip_on_cran()
  skip_if_not_installed("filelock")

  # the child processes use the function of the tested package
  fun <- paste(deparse(doconv:::with_office_lock), collapse = "\n")
  script <- tempfile(fileext = ".R")
  writeLines(c(
    "R_user_dir <- tools::R_user_dir",
    paste("with_office_lock <-", fun),
    "out <- commandArgs(TRUE)[1]",
    "with_office_lock({",
    "  start <- Sys.time()",
    "  Sys.sleep(1)",
    "  writeLines(format(as.numeric(c(start, Sys.time())), digits = 16), out)",
    "})"
  ), script)

  outs <- replicate(2, tempfile(fileext = ".txt"))
  procs <- lapply(outs, function(out) {
    processx::process$new(doconv:::rscript_exec(), c(script, out))
  })
  for (p in procs) p$wait(timeout = 30000)
  expect_true(all(file.exists(outs)))

  times <- lapply(outs, function(f) as.numeric(readLines(f)))
  times <- times[order(vapply(times, `[`, numeric(1), 1))]
  # the second process starts after the end of the first one
  expect_gte(times[[2]][1], times[[1]][2])
})
