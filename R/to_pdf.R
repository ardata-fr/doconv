#' @export
#' @title Convert documents to pdf
#' @description Convert documents to pdf using Libre Office. It
#' supports very well "Microsoft PowerPoint" to PDF. "Microsoft Word"
#' can also be converted but some Word features are not supported
#' such as sections.
#'
#' Windows users must be warned the program is slow on your platform. Performances
#' are not excellent but fast enough on other platform.
#' @param input,output file input and optional file output. If output
#' file is not provided, the value will be the value of input file with
#' extension "pdf".
#' @param use_docx2pdf if TRUE (and if 'Microsoft Word' executable
#' can be found as well as 'docx2pdf'), docx2pdf will be used to
#' convert 'Word' documents to PDF. This makes it possible to have a
#' PDF identical to the 'Word' display whereas with 'LibreOffice', this
#' is not always the case.
#' @param UserInstallation use this value to set a non-default user profile path
#' for "LibreOffice". If not provided a temporary dir is created. It makes possibles
#' to use more than a single session of "LibreOffice."
#' @param timeout timeout in seconds, ignored if 0.
#' @section Ubuntu platforms:
#' On some Ubuntu platforms, 'LibreOffice' require to add in
#' the environment variable `LD_LIBRARY_PATH` the following path:
#' `/usr/lib/libreoffice/program` (you should see the message
#' "libreglo.so cannot open shared object file" if it is the case). This
#' can be done with R
#' command `Sys.setenv(LD_LIBRARY_PATH = "/usr/lib/libreoffice/program/")`
#' @examples
#' library(locatexec)
#' if (exec_available("libreoffice") && check_libreoffice_export()) {
#'
#'   out_pptx <- tempfile(fileext = ".pdf")
#'   file <- system.file(package = "doconv",
#'     "doc-examples/example.pptx")
#'
#'   \donttest{to_pdf(input = file, output = out_pptx)}
#'
#'   out_docx <- tempfile(fileext = ".pdf")
#'   file <- system.file(package = "doconv",
#'     "doc-examples/example.docx")
#'
#'   \donttest{to_pdf(input = file, output = out_docx)}
#'
#' }
#' @return the name of the produced pdf (the same value as `output`),
#' invisibly.
#' @importFrom locatexec libreoffice_exec
#' @importFrom locatexec is_windows
to_pdf <- function(input, output = gsub("\\.[[:alnum:]]+$", ".pdf", input),
                   use_docx2pdf = FALSE, timeout = 120, UserInstallation = NULL) {
  if (!file.exists(input)) {
    stop("input does not exist")
  }
  input <- absolute_path(input)

  default_root <- tempfile()
  dir.create(default_root, showWarnings = FALSE)
  on.exit(
    unlink(default_root, recursive = TRUE, force = TRUE)
  )

  file.copy(
    from = input,
    to = default_root,
    overwrite = TRUE
  )

  if(is.null(UserInstallation)){
    UserInstallation <- absolute_path(tempfile(pattern = "lo_", fileext = ""))
    on.exit(unlink(UserInstallation, recursive = TRUE, force = TRUE))
  } else {
    UserInstallation <- absolute_path(UserInstallation)
  }


  if(grepl("\\.(doc|docx)$", input) && use_docx2pdf && exec_available("word") && docx2pdf_available()){
    docx2pdf(input = input, output = output)
  } else {
    file_set_origin <- list.files(default_root, full.names = TRUE)
    suppressWarnings(
      info <- try(
        system2(
          libreoffice_exec(),
          args = c("--headless",
                   if(!is_windows()) sprintf("\"-env:UserInstallation=file://%s\"", UserInstallation)
                   else sprintf("\"-env:UserInstallation=file:///%s\"", UserInstallation),
                   "--convert-to", "pdf:writer_pdf_Export",
                   "--outdir", shQuote(default_root, type = "cmd"),
                   shQuote(input, type = "cmd")),
          stderr = TRUE, stdout = TRUE, timeout = timeout), silent = TRUE)
    )
    out <- !1 %in% attr(info, "status")
    if(!out) {
      stop(paste0(info, collapse = "\n"))
    }
    file_set_new <- list.files(default_root, full.names = TRUE)
    file_set_new <- setdiff(file_set_new, file_set_origin)
    success <- file.copy(from = file_set_new, to = output, overwrite = TRUE)
    unlink(file.path(default_root, basename(input)), force = TRUE)
    unlink(file_set_new, force = TRUE)
    if (any(!success)) {
      stop("could not convert ", shQuote(input))
    }
  }


  invisible(output)
}

#' @export
#' @title Check if PDF export is functional
#' @description Test if 'LibreOffice' can export to PDF.
#' An attempt to export to PDF is made to confirm that
#' the PDF export is functional.
#' @param UserInstallation use this value to set a non-default user profile path
#' for "LibreOffice". If not provided a temporary dir is created. It makes possibles
#' to use more than a single session of "LibreOffice."
#' @return a single logical value.
#' @examples
#' library(locatexec)
#' if(exec_available("libreoffice")){
#'   \donttest{check_libreoffice_export()}
#' }
check_libreoffice_export <- function(UserInstallation = NULL) {

  default_root <- tempfile()
  dir.create(default_root, showWarnings = FALSE)
  on.exit(
    unlink(default_root, recursive = TRUE, force = TRUE)
  )

  input <- file.path(default_root, "minimal-word.docx")

  file.copy(
    system.file(package = "doconv", "doc-examples", "minimal-word.docx"),
    default_root
  )

  if(is.null(UserInstallation)){
    UserInstallation <- absolute_path(tempfile(pattern = "lo_", fileext = ""))
    on.exit(unlink(UserInstallation, recursive = TRUE, force = TRUE))
  } else {
    UserInstallation <- absolute_path(UserInstallation)
  }

  suppressWarnings(
    try(
      system2(
        libreoffice_exec(),
        args = c("--headless",# useless unless with older versions
                 if(!is_windows()) sprintf("\"-env:UserInstallation=file://%s\"", UserInstallation)
                 else sprintf("\"-env:UserInstallation=file:///%s\"", UserInstallation),
                 "--convert-to", "pdf:writer_pdf_Export",
                 "--outdir", shQuote(default_root, type = "cmd"),
                 shQuote(input, type = "cmd")),
        stderr = TRUE, stdout = TRUE, timeout = 15), silent = TRUE)
  )
  expected <- file.path(default_root, "minimal-word.pdf")
  success <- file.exists(expected)

  success
}

