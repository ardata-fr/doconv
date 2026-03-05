#' @export
#' @title Convert xlsx to pdf
#' @description Convert xlsx to pdf directly using "Microsoft Excel".
#'
#' This function will not work if "Microsoft Excel" is not available
#' on your machine.
#'
#' The calls to "Microsoft Excel" are made differently depending on the
#' operating system:
#' - On "Windows", a "PowerShell" script using COM
#' technology is used to control "Microsoft Excel".
#' - On macOS, an "AppleScript" script is used to control "Microsoft Excel".
#' @section Windows authorizations:
#' If your execution policy is set to "RemoteSigned", 'doconv' will
#' not be able to run powershell script. Set it to "Unrestricted" and it
#' should work. If you are in a managed and administrated environment,
#' you may not be able to use 'doconv' because of execution policies.
#' @section Macos manual authorizations:
#' On macOS the call is happening into a working
#' directory managed with function [working_directory()].
#'
#' Manual interventions are necessary to authorize 'Excel'
#' application to write in a single directory: the working directory.
#' These permissions must be set manually, this is required by the macOS security
#' policy.
#' @param input,output file input and optional file output. If output
#' file is not provided, the value will be the value of input file with
#' extension 'pdf'.
#' @examples
#' library(locatexec)
#' if (exec_available("excel")) {
#'   xlsx2pdf(
#'     input = system.file(
#'       package = "doconv",
#'       "doc-examples/example.xlsx"
#'     ),
#'     output = tempfile(fileext = ".pdf")
#'   )
#' }
#' @return the name of the produced pdf (the same value as `output`)
xlsx2pdf <- function(input, output = gsub("\\.(xlsx|xlsm)$", ".pdf", input)) {

  if (is_osx()) {
    xlsx2pdf_osx(input = input, output = output)
  } else if (is_windows()) {
    xlsx2pdf_win(input = input, output = output)
  } else {
    stop("xlsx2pdf is only available on 'macOS' and 'Windows' systems.", call. = FALSE)
  }
}


#' @importFrom locatexec is_osx
#' @importFrom processx run
xlsx2pdf_osx <- function(input, output = gsub("\\.(xlsx|xlsm)$", ".pdf", input)) {

  if (!is_osx()) {
    stop("xlsx2pdf_osx() should only be used on 'macOS' systems.", call. = FALSE)
  }

  if (!file.exists(input)) {
    stop("input does not exist", call. = FALSE)
  }

  input <- absolute_path(input)
  output <- absolute_path(output)

  init_working_directory()
  default_root <- working_directory()

  output_name <- file.path(default_root, gsub("\\.(xlsx|xlsm)$", ".pdf", basename(input)))

  script_sourcefile <- system.file(package = "doconv", "scripts", "applescripts", "xlsx2pdf.applescript")
  script_path <- tempfile(fileext = ".applescript")
  script_str <- readLines(script_sourcefile, encoding = "UTF-8")
  script_str[1] <- sub("%s", escape_path_applescript(input), script_str[1], fixed = TRUE)
  script_str[2] <- sub("%s", escape_path_applescript(output_name), script_str[2], fixed = TRUE)
  writeLines(script_str, script_path, useBytes = TRUE)

  res <- run("osascript", script_path, error_on_status = FALSE)

  success <- res$status == 0

  if (success) {
    success <- file.copy(from = output_name, to = output, overwrite = TRUE)
  }

  rm_working_directory()

  if (!success) stop("could not convert ", input, call. = FALSE)

  output
}


#' @importFrom locatexec is_windows
xlsx2pdf_win <- function(input, output = gsub("\\.(xlsx|xlsm)$", ".pdf", input)) {

  if (!is_windows()) {
    stop("xlsx2pdf_win() should only be used on 'Windows' systems.", call. = FALSE)
  }

  if (!file.exists(input)) {
    stop("input does not exist", call. = FALSE)
  }

  input <- absolute_path(input)
  output <- absolute_path(output)

  script_sourcefile <- system.file(
    package = "doconv", "scripts", "powershell", "xlsx2pdf.ps1")
  script_path <- tempfile(fileext = ".ps1")
  script_str <- readLines(script_sourcefile, encoding = "UTF-8")
  script_str[1] <- sub("%s", escape_path_ps(input), script_str[1], fixed = TRUE)
  script_str[2] <- sub("%s", escape_path_ps(normalizePath(output, mustWork = FALSE)), script_str[2], fixed = TRUE)
  writeLines(script_str, script_path, useBytes = TRUE)

  res <- run("powershell", args = c("-file", script_path), error_on_status = FALSE)
  success <- res$status == 0

  # fail with informative error if conversion fails due to PS Execution Policy
  if (!success && grepl("UnauthorizedAccess", res$stderr)) {
    stop_on_wrong_ps_exec_policy()
  }

  if (!success) stop("could not convert ", input, call. = FALSE)

  output
}
