#' @export
#' @title Convert documents to pdf
#' @description Convert documents to pdf with a script using
#' 'Office' or 'Libre Office'.
#'
#' If 'Microsoft Word' and 'Microsoft PowerPoint' are available,
#' files 'docx', 'doc', 'rtf' and 'pptx' will be converted to
#' PDF with 'Office' via a script.
#'
#' If 'Microsoft Word' and 'Microsoft PowerPoint' are not available
#' (on linux for example), 'Libre Office' will be used to convert
#' documents. In that case the rendering can be different from
#' the original document. It supports very well 'Microsoft PowerPoint'
#' to PDF. 'Microsoft Word' can also be converted but some Word
#' features are not supported, such as sections.
#'
#' @param input,output file input and optional file output. If output
#' file is not provided, the value will be the value of input file with
#' extension 'pdf'.
#' @param UserInstallation use this value to set a non-default user profile path
#' for 'LibreOffice'. If not provided a temporary dir is created. It makes possibles
#' to use more than a single session of 'LibreOffice'.
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
                   timeout = 120, UserInstallation = NULL) {
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


  if(grepl("\\.(doc|docx|rtf)$", input) && exec_available("word")){
    docx2pdf(input = input, output = output)
  } else if(grepl("\\.(ppt|pptx)$", input) && exec_available("powerpoint")){
    pptx2pdf(input = input, output = output)
  } else {
    file_set_origin <- list.files(default_root, full.names = TRUE)

    res <- run(
          libreoffice_exec(),
          args = c("--headless",
                   if(!is_windows()) sprintf("-env:UserInstallation=file://%s", UserInstallation)
                   else sprintf("-env:UserInstallation=file:///%s", UserInstallation),
                   "--convert-to", "pdf:writer_pdf_Export",
                   "--outdir", default_root,
                   input),
          error_on_status = FALSE)
    success <- res$status == 0

    if(!success) stop("could not convert ", input, call. = FALSE)

    file_set_new <- list.files(default_root, full.names = TRUE)
    file_set_new <- setdiff(file_set_new, file_set_origin)
    success <- file.copy(from = file_set_new, to = output, overwrite = TRUE)
    unlink(file.path(default_root, basename(input)), force = TRUE)
    unlink(file_set_new, force = TRUE)
    if (any(!success)) {
      stop("could not convert ", input, call. = FALSE)
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
    default_root, overwrite = TRUE
  )

  if(is.null(UserInstallation)){
    UserInstallation <- absolute_path(tempfile(pattern = "lo_", fileext = ""))
    on.exit(unlink(UserInstallation, recursive = TRUE, force = TRUE))
  } else {
    UserInstallation <- absolute_path(UserInstallation)
  }

  res <- run(command = libreoffice_exec(),
      args = c("--headless",# useless unless with older versions
               if(!is_windows()) sprintf("-env:UserInstallation=file://%s", UserInstallation)
               else sprintf("-env:UserInstallation=file:///%s", UserInstallation),
               "--convert-to", "pdf:writer_pdf_Export",
               "--outdir", default_root,
               input),
      error_on_status = FALSE,
      timeout = 15
      )
  success <- res$status == 0

  expected <- file.path(default_root, "minimal-word.pdf")
  success <- file.exists(expected)

  success
}
