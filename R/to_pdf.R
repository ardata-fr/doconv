#' @export
#' @title Convert documents to pdf
#' @description Convert documents to pdf using Libre Office. It
#' supports very well "Microsoft PowerPoint" to PDF. "Microsoft Word"
#' can also be converted but some Word features are not supported
#' such as sections.
#' @param input,output file input and optional file output.
#' @examples
#' library(locatexec)
#' if (exec_available("libreoffice")) {
#'   file <- tempfile(fileext = ".pptx")
#'   file.copy(
#'     system.file(package = "doconv",
#'       "doc-examples/first_example.pptx"),
#'     file)
#'
#'   out <- to_pdf(input = file)
#' }
#' @importFrom locatexec libreoffice_exec
#' @importFrom locatexec is_windows
to_pdf <- function(input, output = gsub("\\.[[:alnum:]]+$", ".pdf", input)) {
  if (!file.exists(input)) {
    stop("input does not exist")
  }
  input <- absolute_path(input)

  init_working_directory(force = TRUE)
  default_root <- working_directory()

  file.copy(
    from = input,
    to = default_root,
    overwrite = TRUE
  )
  file_set_origin <- list.files(default_root, full.names = TRUE)
  old_warn <- getOption("warn")
  options(warn = -1)
  info <- try(
    system2(
      libreoffice_exec(),
      args = c("--headless",
               if(!is_windows()) "\"-env:UserInstallation=file:///tmp/LibreOffice_Conversion_${USER}\"",
               "--convert-to", "pdf:writer_pdf_Export",
               "--outdir", shQuote(default_root, type = "cmd"),
               shQuote(input, type = "cmd")),
      stderr = TRUE, stdout = TRUE), silent = TRUE)
  options(warn = old_warn)
  out <- !1 %in% attr(info, "status")
  if(!out) {
    stop(paste0(info, collapse = "\n"))
  }

  file_set_new <- list.files(default_root, full.names = TRUE)
  file_set_new <- setdiff(file_set_new, file_set_origin)
  success <- file.copy(from = file_set_new, to = output, overwrite = TRUE)
  unlink(file.path(default_root, basename(input)), force = TRUE)
  unlink(file_set_new, force = TRUE)
  if (any(!success)) stop("could not convert ", shQuote(input))
  output
}

