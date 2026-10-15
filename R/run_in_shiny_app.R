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

  url <- sprintf("http://127.0.0.1:%s/", port)
  # another app answering on the port would be used instead of this one
  if (http_responds(url)) {
    stop("port ", port, " is already in use, ", url, " is responding.",
         call. = FALSE)
  }

  # --- start the Shiny app ---
  # stderr goes to a file: a pipe that is never read can fill up and
  # block the app
  stderr_file <- tempfile(fileext = ".txt")
  app <- processx::process$new(
    command = rscript_exec(),
    args = c("-e", app_expr),
    stderr = stderr_file,
    supervise = TRUE
  )
  on.exit({
    app$kill()
    unlink(stderr_file)
  }, add = TRUE)

  wait_for_http(url, timeout = timeout, process = app,
                stderr_file = stderr_file)

  # --- run the command ---
  processx::run(...)
}

#' Poll an HTTP URL until it responds with status 200
#' @param url URL to poll.
#' @param timeout Maximum wait time in seconds.
#' @param interval Seconds between attempts.
#' @param process Optional processx process serving `url`; the wait stops
#'   with an error as soon as it is no longer alive.
#' @param stderr_file Optional file where `process` writes its standard
#'   error, its last lines are added to the error message.
#' @return `TRUE` (invisibly) on success; throws an error on timeout or
#'   when `process` has stopped.
#' @noRd
wait_for_http <- function(url, timeout = 30, interval = 1,
                          process = NULL, stderr_file = NULL) {
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
    if (!is.null(process) && !process$is_alive()) {
      msg <- character()
      if (!is.null(stderr_file) && file.exists(stderr_file)) {
        msg <- readLines(stderr_file, warn = FALSE)
        msg <- msg[seq_along(msg) > length(msg) - 20]
      }
      stop(
        "The app stopped before ", url, " was ready (exit status ",
        process$get_exit_status(), ").",
        if (length(msg)) paste0("\n", paste(msg, collapse = "\n")),
        call. = FALSE
      )
    }
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

#' Check if an HTTP URL is responding, whatever the status code
#' @param url URL to check.
#' @return `TRUE` if a server answers, `FALSE` otherwise.
#' @noRd
http_responds <- function(url) {
  tryCatch(
    {
      curl::curl_fetch_memory(url)
      TRUE
    },
    error = function(e) FALSE
  )
}
