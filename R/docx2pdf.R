#' @export
#' @title Convert docx to pdf
#' @description Convert docx to pdf directly using "Microsoft Word".
#"
#' This function will not work if "Microsoft Word" is not available
#' on your machine.
#'
#' The calls to "Microsoft Word" are made differently depending on the
#' operating system:
#' - On "Windows", a "PowerShell" script using COM
#' technology is used to control "Microsoft Word". The resulting PDF
#' is containing a browsable TOC.
#' - On macOS, an "AppleScript" script is used to control "Microsoft Word".
#' The resulting PDF is not containing a browsable TOC as when on 'Windows'.
#' @section Macos manual authorizations:
#' On macOS the call is happening into a working
#' directory managed with function [working_directory()].
#'
#' Manual interventions are necessary to authorize 'Word' and
#' 'PowerPoint' applications to write in a single directory: the working directory.
#' These permissions must be set manually, this is required by the macOS security
#' policy. We think that this is not a problem because it is unlikely that you will
#' use a Mac machine as a server.
#'
#' You must click "allow" two times to:
#'
#' 1. allow R to run 'AppleScript' scripts that will control Word
#' 2. allow Word to write to the working directory.
#'
#' This are one-time operations.
#'
#' @param input,output file input and optional file output (default
#' to input with pdf extension).
#' @examples
#' library(locatexec)
#' if (exec_available('word')) {
#'   file <- system.file(package = "doconv",
#'     "doc-examples/example.docx")
#'
#'   out <- docx2pdf(input = file,
#'     output = tempfile(fileext = ".pdf"))
#'
#'   if (file.exists(out)) {
#'     message(basename(out), " is existing now.")
#'   }
#'
#'   if (require("officer")) {
#'     doc <- read_docx()
#'     doc <- body_add_fpar(doc,
#'         value = fpar(
#'           run_word_field("DOCPROPERTY \"coco\" \\* MERGEFORMAT")))
#'     doc <- set_doc_properties(doc, coco = "test")
#'     file <- print(doc, target = "output.docx") |>
#'     docx_update(file)
#'   }
#' }
#' @return the name of the produced pdf (the same value as `output`)
docx2pdf <- function(input, output = gsub("\\.docx$", ".pdf", input)) {

  if (is_osx()) {
    docx2pdf_osx(input = input, output = output)
  } else if (is_windows()) {
    docx2pdf_win(input = input, output = output)
  } else {
    stop("docx2pdf is only available on 'macOS' and 'Windows' systems.", call. = FALSE)
  }
}

#' @importFrom locatexec is_osx
#' @importFrom processx run
docx2pdf_osx <- function(input, output = gsub("\\.docx$", ".pdf", input)){

  if (!is_osx()) {
    stop("docx2pdf_osx() should only be used on 'macOS' systems.", call. = FALSE)
  }

  if (!file.exists(input)) {
    stop("input does not exist", call. = FALSE)
  }

  input <- absolute_path(input)
  output <- absolute_path(output)

  init_working_directory()
  default_root <- working_directory()

  output_name <- file.path(default_root, gsub("\\.docx$", ".pdf", basename(input)))

  script_sourcefile <- system.file(package = "doconv", "scripts", "applescripts", "docx2pdf.applescript")
  script_path <- tempfile(fileext = ".applescript")
  script_str <- readLines(script_sourcefile, encoding = "UTF-8")
  script_str[1] <- sprintf(script_str[1], input)
  script_str[2] <- sprintf(script_str[2], output_name)
  writeLines(script_str, script_path, useBytes = TRUE)

  res <- run("osascript", script_path, error_on_status = FALSE)
  success <- res$status == 0

  if(success) {
    success <- file.copy(from = output_name, to = output, overwrite = TRUE)
  }
  rm_working_directory()

  if(!success) stop("could not convert ", input, call. = FALSE)

  output
}

#' @importFrom locatexec is_windows
docx2pdf_win <- function(input, output = gsub("\\.docx$", ".pdf", input)){

  if (!is_windows()) {
    stop("docx2pdf_win() should only be used on 'Windows' systems.", call. = FALSE)
  }

  if (!file.exists(input)) {
    stop("input does not exist", call. = FALSE)
  }

  input <- absolute_path(input)
  output <- absolute_path(output)

  init_working_directory()
  default_root <- working_directory()

  output_name <- file.path(default_root, gsub("\\.docx$", ".pdf", basename(input)))

  script_sourcefile <- system.file(
    package = "doconv", "scripts", "powershell", "docx2pdf.ps1")
  script_path <- tempfile(fileext = ".ps1")
  script_str <- readLines(script_sourcefile, encoding = "UTF-8")
  script_str[1] <- sprintf(script_str[1], input)
  script_str[2] <- sprintf(script_str[2], output_name)
  writeLines(script_str, script_path, useBytes = TRUE)

  res <- run("powershell", args = c("-file", script_path), error_on_status = FALSE)
  success <- res$status == 0

  if(success) {
    success <- file.copy(from = output_name, to = output, overwrite = TRUE)
  }

  rm_working_directory()

  if(!success) stop("could not convert ", input, call. = FALSE)

  output


}

#' @export
#' @title Update docx fields
#' @description Update all fields and table of contents of a
#' Word document using "Microsoft Word".
#' This function will not work if "Microsoft Word" is not available
#' on your machine.
#'
#' The calls to "Microsoft Word" are made differently depending on the
#' operating system. On "Windows", a "PowerShell" script using COM
#' technology is used to control "Microsoft Word". On macOS, an "AppleScript"
#' script is used to control "Microsoft Word".
#' @param input file input
#' @examples
#' library(locatexec)
#' if (exec_available('word')) {
#'   file <- system.file(package = "doconv",
#'     "doc-examples/example.docx")
#'   docx_out <- tempfile(fileext = ".docx")
#'   file.copy(file, docx_out)
#'   docx_update(input = docx_out)
#' }
#' @return the name of the produced pdf (the same value as `output`)
docx_update <- function(input) {

  if (is_osx()) {
    docx_update_osx(input = input)
  } else if (is_windows()) {
    docx_update_win(input = input)
  } else {
    stop("docx_update is only available on 'macOS' and 'Windows' systems.", call. = FALSE)
  }
}

docx_update_osx <- function(input){

  if (!is_osx()) {
    stop("docx2pdf_osx() should only be used on 'macOS' systems.", call. = FALSE)
  }

  if (!file.exists(input)) {
    stop("input does not exist", call. = FALSE)
  }

  input <- absolute_path(input)

  init_working_directory()

  script_sourcefile <- system.file(package = "doconv", "scripts", "applescripts", "docxupdate.applescript")
  script_path <- tempfile(fileext = ".applescript")
  script_str <- readLines(script_sourcefile, encoding = "UTF-8")
  script_str[1] <- sprintf(script_str[1], input)
  writeLines(script_str, script_path, useBytes = TRUE)

  res <- run("osascript", script_path)
  success <- res$status == 0

  success
}
docx_update_win <- function(input){

  if (!is_windows()) {
    stop("docx_update_win() should only be used on 'Windows' systems.", call. = FALSE)
  }

  if (!file.exists(input)) {
    stop("input does not exist", call. = FALSE)
  }

  input <- absolute_path(input)

  init_working_directory()
  default_root <- working_directory()

  script_sourcefile <- system.file(package = "doconv", "scripts", "powershell", "docxupdate.ps1")
  script_path <- tempfile(fileext = ".ps1")
  docxupdate_str <- readLines(script_sourcefile, encoding = "UTF-8")
  docxupdate_str[1] <- sprintf(docxupdate_str[1], input)
  writeLines(docxupdate_str, script_path, useBytes = TRUE)

  res <- run("powershell", args = c("-file", script_path), error_on_status = FALSE)
  success <- res$status == 0

  if(!success) stop("could not update ", input, call. = FALSE)

  success
}






#' @export
#' @title Convert pptx to pdf
#' @description Convert pptx to pdf directly using "Microsoft PowerPoint".
#' This function will not work if "Microsoft PowerPoint" is not available
#' on your machine.
#'
#' The calls to "Microsoft PowerPoint" are made differently depending on
#' the operating system. On "Windows", a "PowerShell" script using COM
#' technology is used to control "Microsoft PowerPoint". On macOS, an
#' "AppleScript" script is used to control "Microsoft PowerPoint".
#' @section Macos manual authorizations:
#' On macOS the call is happening into a working
#' directory managed with function [working_directory()].
#'
#' Manual interventions are necessary to authorize
#' 'PowerPoint' applications to write in a single directory: the working directory.
#' These permissions must be set manually, this is required by the macOS security
#' policy. We think that this is not a problem because it is unlikely that you will
#' use a Mac machine as a server.
#'
#' You must also click "allow" two times to:
#'
#' 1. allow R to run 'AppleScript' scripts that will control PowerPoint
#' 2. allow PowerPoint to write to the working directory.
#'
#' This are one-time operations.
#' @param input,output file input and optional file output (default
#' to input with pdf extension).
#' @examples
#' library(locatexec)
#' if (exec_available('powerpoint')) {
#'   file <- system.file(package = "doconv",
#'     "doc-examples/example.pptx")
#'
#'   out <- pptx2pdf(input = file,
#'     output = tempfile(fileext = ".pdf"))
#'
#'   if (file.exists(out)) {
#'     message(basename(out), " is existing now.")
#'   }
#' }
#' @return the name of the produced pdf (the same value as `output`)
pptx2pdf <- function(input, output = gsub("\\.pptx$", ".pdf", input)) {

  if (is_osx()) {
    pptx2pdf_osx(input = input, output = output)
  } else if (is_windows()) {
    pptx2pdf_win(input = input, output = output)
  } else {
    stop("pptx2pdf is only available on 'macOS' and 'Windows' systems.", call. = FALSE)
  }
}

#' @importFrom locatexec is_osx
#' @importFrom processx run
pptx2pdf_osx <- function(input, output = gsub("\\.pptx$", ".pdf", input)){

  if (!is_osx()) {
    stop("pptx2pdf_osx() should only be used on 'macOS' systems.", call. = FALSE)
  }

  if (!file.exists(input)) {
    stop("input does not exist", call. = FALSE)
  }

  input <- absolute_path(input)
  output <- absolute_path(output)

  init_working_directory()
  default_root <- working_directory()

  output_name <- file.path(default_root, gsub("\\.pptx$", ".pdf", basename(input)))

  script_sourcefile <- system.file(package = "doconv", "scripts", "applescripts", "pptx2pdf.applescript")
  script_path <- tempfile(fileext = ".applescript")
  script_str <- readLines(script_sourcefile, encoding = "UTF-8")
  script_str[1] <- sprintf(script_str[1], input)
  script_str[2] <- sprintf(script_str[2], output_name)
  writeLines(script_str, script_path, useBytes = TRUE)

  res <- run("osascript", script_path, error_on_status = FALSE)

  success <- res$status == 0

  if(success) {
    success <- file.copy(from = output_name, to = output, overwrite = TRUE)
  }

  rm_working_directory()

  if(!success) stop("could not convert ", input, call. = FALSE)

  output
}

#' @importFrom locatexec is_windows
pptx2pdf_win <- function(input, output = gsub("\\.pptx$", ".pdf", input)){

  if (!is_windows()) {
    stop("pptx2pdf_win() should only be used on 'Windows' systems.", call. = FALSE)
  }

  if (!file.exists(input)) {
    stop("input does not exist", call. = FALSE)
  }

  input <- absolute_path(input)
  output <- absolute_path(output)

  script_sourcefile <- system.file(
    package = "doconv", "scripts", "powershell", "pptx2pdf.ps1")
  script_path <- tempfile(fileext = ".ps1")
  script_str <- readLines(script_sourcefile, encoding = "UTF-8")
  script_str[1] <- sprintf(script_str[1], input)
  script_str[2] <- sprintf(script_str[2], normalizePath(output, mustWork = FALSE))
  writeLines(script_str, script_path, useBytes = TRUE)

  res <- run("powershell", args = c("-file", script_path), error_on_status = FALSE)
  success <- res$status == 0

  if(!success) stop("could not convert ", input, call. = FALSE)

  success


}

