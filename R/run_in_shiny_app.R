#' @title Run a command against a running Shiny application
#' @description Starts a Shiny app in a background process, waits for it
#' to be ready, executes an arbitrary command via [processx::run()],
#' and stops the app on exit.
#'
#' @param app_expr A character string containing an R expression to launch
#'   the Shiny app, executed via `Rscript -e`. The expression must start
#'   the app on the specified `port`
#'   (e.g. `"shiny::runApp('app/', port = 9999)"`).
#' @param port HTTP port used by the Shiny app. Default is `9999`.
#' @param timeout Maximum number of seconds to wait for the app to be
#'   ready. Default is `120`.
#' @param ... Arguments passed to [processx::run()]. At minimum,
#'   `command` must be provided.
#' @return The value returned by [processx::run()] (a list with
#'   `status`, `stdout`, `stderr`, `timeout`).
#' @export
#' @examples
#' \dontrun{
#' res <- run_in_shiny_app(
#'   app_expr = "shiny::runApp('app/', port = 9999)",
#'   command = "npm", args = c("run", "scenario", "import-flow"),
#'   wd = tempdir()
#' )
#' }
run_in_shiny_app <- function(
    app_expr,
    port = 9999,
    timeout = 120,
    ...) {

  if (!requireNamespace("curl", quietly = TRUE)) {
    stop(
      "Package 'curl' is required by `run_in_shiny_app()`. ",
      "Install it with: install.packages('curl')",
      call. = FALSE
    )
  }

  # --- start the Shiny app ---
  app <- processx::process$new(
    command = "Rscript",
    args = c("-e", app_expr),
    supervise = TRUE
  )
  on.exit(app$kill(), add = TRUE)

  url <- sprintf("http://127.0.0.1:%s/", port)
  wait_for_http(url, timeout = timeout)

  # --- run the command ---
  processx::run(...)
}

#' Poll an HTTP URL until it responds with status 200
#' @param url URL to poll.
#' @param timeout Maximum wait time in seconds.
#' @param interval Seconds between attempts.
#' @return `TRUE` (invisibly) on success; throws an error on timeout.
#' @noRd
wait_for_http <- function(url, timeout = 30, interval = 1) {
  deadline <- Sys.time() + timeout

  repeat {
    ready <- tryCatch(
      {
        resp <- curl::curl_fetch_memory(url)
        resp$status_code == 200L
      },
      error = function(e) FALSE
    )
    if (ready) return(invisible(TRUE))
    if (Sys.time() > deadline) {
      stop(
        "Timeout waiting for app at ", url,
        " (", timeout, "s elapsed).",
        call. = FALSE
      )
    }
    Sys.sleep(interval)
  }
}
