#' @title Is 'docx2pdf' available
#' @description Checks if 'docx2pdf' is available within a given 'python' distribution.
#' @param error Whether to signal an error if 'docx2pdf' is not found
#' @return a single logical value.
#' @export
#' @importFrom locatexec python_exec
#' @examples
#' library(locatexec)
#'
#' if(exec_available("python") &&
#'    exec_version("python") > numeric_version("3")) {
#'   docx2pdf_available()
#' }
#' @family tools for docx2pdf
docx2pdf_available <- function(error = FALSE) {
  suppressWarnings(
    info <- try(system2(python_exec(), args = c("-c", shQuote("import docx2pdf", type = "cmd")),
                    stderr = TRUE, stdout = TRUE), silent = TRUE)
  )
  out <- !1 %in% attr(info, "status")

  if (error && !out) {
    stop(
      "'docx2pdf' is not available or cannot be found. ",
      "Please use `docx2pdf_install()` to install it.",
      "If you think it is installed, you may have to define the python directory with `find_python(dir='/path/to/python')`"
    )
  }

  out
}

#' @export
#' @title Uninstall 'docx2pdf'
#' @description Removes 'docx2pdf' within a given 'python' distribution.
#' @return a single logical value, FALSE if the operation failed, TRUE otherwise.
#' @examples
#' library(locatexec)
#'
#' if(exec_available("python") &&
#'    exec_version("python") > numeric_version("3") &&
#'    docx2pdf_available()) {
#'   \donttest{
#'   # this will uninstall and install 'docx2pdf' (i.e. `pip install docx2pdf`)
#'   # in your python environnement. It can run during 10s and
#'   # require user to have write permission in python environnement.
#'   docx2pdf_uninstall()
#'   docx2pdf_install()
#'   }
#' }
#' @family tools for docx2pdf
docx2pdf_uninstall <- function() {

  suppressWarnings(
    info <- try(system2(pip_exec(), args = c("uninstall", "-y", "docx2pdf"),
                        stderr = TRUE, stdout = TRUE), silent = TRUE)
  )
  out <- !1 %in% attr(info, "status")
  out
}



#' @export
#' @title Install 'docx2pdf'
#' @description Downloads and installs 'docx2pdf' within a given 'python' distribution.
#' @param force Whether to force to install (override) 'docx2pdf'.
#' @return a single logical value, FALSE if the operation failed, TRUE otherwise.
#' @examples
#' library(locatexec)
#'
#' if(exec_available("python") &&
#'    exec_version("python") > numeric_version("3") &&
#'    !docx2pdf_available()) {
#'   \donttest{
#'   # this will install and uninstall 'docx2pdf' (i.e. `pip install docx2pdf`)
#'   # in your python environnement. It can run during 10s and
#'   # require user to have write permission in python environnement.
#'   docx2pdf_install()
#'   docx2pdf_uninstall()
#'   }
#' }
#' @importFrom locatexec pip_exec exec_available
#' @family tools for docx2pdf
docx2pdf_install <- function(force = FALSE) {
  if (force) {
    docx2pdf_uninstall()
  }

  suppressWarnings(
    info <- try(system2(pip_exec(), args = c("install", "docx2pdf"),
      stderr = TRUE, stdout = TRUE), silent = TRUE)
  )
  out <- !1 %in% attr(info, "status")
  out
}


#' @importFrom locatexec is_windows
docx2pdf_exec <- function() {
  exec_available("python", error = TRUE)
  docx2pdf_available(error = TRUE)

  bin <- "docx2pdf"
  if(is_windows()) bin <- paste0("Scripts/", bin)

  file.path(dirname(python_exec()), bin)
}

#' @export
#' @title Convert docx to pdf
#' @description Convert docx to pdf directly using "Microsoft Word".
#' This function will not work if "Microsoft Word" is not available
#' on your machine.
#'
#' On Windows, this is implemented via win32com while on macOS this is
#' implemented via JXA (Javascript for Automation, aka AppleScript in JS).
#'
#' This is a simple call to python module 'docx2pdf' into a working
#' directory managed with function [working_directory()].
#' @param input,output file input and optional file output (default
#' to input with pdf extension).
#' @examples
#' library(locatexec)
#' if (exec_available('python') && docx2pdf_available()) {
#'   file <- system.file(package = "doconv",
#'     "doc-examples/example.docx")
#'
#'   out <- docx2pdf(input = file,
#'     output = tempfile(fileext = ".pdf"))
#'
#'   if (file.exists(out)) {
#'     message(basename(out), " is existing now.")
#'   }
#' }
#' @return the name of the produced pdf (the same value as `output`)
#' @family tools for docx2pdf
docx2pdf <- function(input, output = gsub("\\.docx$", ".pdf", input)) {

  if (!file.exists(input)) {
    stop("input does not exist")
  }

  input <- absolute_path(input)
  output <- absolute_path(output)

  init_working_directory()
  default_root <- working_directory()

  file.copy(from = input,
            to = default_root,
            overwrite = TRUE)
  output_name <- file.path(default_root, gsub("\\.docx$", ".pdf", basename(input)))

  suppressWarnings(
    info <- try(
      system2(
        docx2pdf_exec(),
        args = shQuote(default_root, type = "cmd"),
        stderr = TRUE, stdout = TRUE), silent = TRUE)
  )
  out <- !1 %in% attr(info, "status")
  if(!out) {
    stop(paste0(info, collapse = "\n"))
  }
  success <- file.copy(from = output_name, to = output, overwrite = TRUE)

  unlink(file.path(default_root, basename(input)), force = TRUE)
  rm_working_directory()
  if(!success) stop("could not convert ", input)
  output
}

