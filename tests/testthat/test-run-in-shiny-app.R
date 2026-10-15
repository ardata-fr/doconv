test_that("run_in_shiny_app() fails fast when the app stops", {
  skip_on_cran()
  skip_if_not_installed("curl")

  start <- Sys.time()
  expect_error(
    run_in_shiny_app("stop('boom')", port = 9991, timeout = 30,
                     command = "echo", args = "ok"),
    "app stopped.*boom"
  )
  elapsed <- as.numeric(difftime(Sys.time(), start, units = "secs"))
  expect_lt(elapsed, 20)
})

test_that("run_in_shiny_app() fails when the port is already in use", {
  skip_on_cran()
  skip_if_not_installed("curl")

  # another app is answering on the port
  other <- processx::process$new(
    doconv:::rscript_exec(),
    c("-e", "options(help.ports = 9992); tools::startDynamicHelp(TRUE); Sys.sleep(60)")
  )
  on.exit(other$kill(), add = TRUE)
  ready <- FALSE
  for (i in 1:30) {
    ready <- doconv:::http_responds("http://127.0.0.1:9992/")
    if (ready) break
    Sys.sleep(.5)
  }
  skip_if_not(ready)

  expect_error(
    run_in_shiny_app("stop('boom')", port = 9992, timeout = 30,
                     command = "echo", args = "ok"),
    "already in use"
  )
})
