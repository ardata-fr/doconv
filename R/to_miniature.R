#' @import magick
#' @export
#' @title Thumbnail of a document
#' @description Convert a file into an image (magick image) where the
#' pages are arranged in rows, each row can contain one to several pages.
#'
#' The result can be saved as a png file.
#' @param filename input filename, supported documents are 'Microsoft Word',
#' 'Microsoft PowerPoint', 'RTF' and 'PDF' document.
#' @param row row index for every pages. 0 are to be used to drop
#' the page from the final minature. If both `row` and `ncol` are
#' provided, `row` takes precedence and a warning is issued.
#'
#' * `c(1, 1)` is to be used to specify that a 2 pages document
#' is to be displayed in a single row with two columns.
#' * `c(1, 1, 2, 3, 3)` is to be used to specify that a 5 pages document
#' is to be displayed as: first row with pages 1 and 2, second row with page 3,
#' third row with pages 4 and 5.
#' * `c(1, 1, 0, 2, 2)` is to be used to specify that a 5 pages document
#' is to be displayed as: first row with pages 1 and 2,
#' second row with pages 4 and 5.
#'
#' @param ncol number of pages per row. When set, pages are automatically
#' arranged with `ncol` pages per row. Ignored if `row` is also provided.
#' @param ncol_landscape number of landscape-oriented pages per row.
#' When set, portrait pages use `ncol` per row and landscape pages use
#' `ncol_landscape` per row. Requires `ncol`.
#' @param width width of a single image, recommanded values are:
#'
#' * 650 for docx files
#' * 750 for pptx files
#'
#' @param border_color border color, see [magick::image_border()].
#' @param border_geometry border geometry to be added around
#' images, see [magick::image_border()].
#' @param fileout if not NULL, result is saved in a png file whose filename
#' is defined by this argument.
#' @param dpi resolution (dots per inch) to use for images, see [pdftools::pdf_convert()].
#' @param timeout timeout in seconds that libreoffice is allowed to use
#' in order to generate the corresponding pdf file, ignored if 0.
#' @param ... arguments used by webshot2 when HTML document.
#' @return a magick image object as returned by [magick::image_read()].
#' @examples
#' library(locatexec)
#' docx_file <- system.file(
#'   package = "doconv",
#'   "doc-examples/example.docx"
#' )
#' if(exec_available("word"))
#'   to_miniature(docx_file)
#'
#' pptx_file <- system.file(
#'   package = "doconv",
#'   "doc-examples/example.pptx"
#' )
#' if(exec_available("libreoffice") && check_libreoffice_export())
#'   to_miniature(pptx_file)
#' @importFrom locatexec exec_available
to_miniature <- function(filename, row = NULL, ncol = NULL,
                         ncol_landscape = NULL, width = NULL,
                         border_color = "#ccc", border_geometry = "2x2",
                         dpi = 150,
                         fileout = NULL, timeout = 120,
                         ...) {

  if (!file.exists(filename)) {
    stop("filename does not exist")
  }

  if(grepl("\\.(ppt|pptx)$", filename)){
    if(is.null(width)) width <- 750
    pptx_to_miniature(
      filename, row = row, ncol = ncol, ncol_landscape = ncol_landscape,
      width = width,
      border_color = border_color, border_geometry = border_geometry,
      fileout = fileout, dpi = dpi, timeout = timeout)
  } else if(grepl("\\.(docx|doc|rtf)$", filename)){
    if(is.null(width)) width <- 650
    docx_to_miniature(
      filename, row = row, ncol = ncol, ncol_landscape = ncol_landscape,
      width = width,
      border_color = border_color, border_geometry = border_geometry,
      fileout = fileout, timeout = timeout)
  } else if(grepl("\\.pdf$", filename)){
    if(is.null(width)) width <- 650
    pdf_to_miniature(
      filename, row = row, ncol = ncol, ncol_landscape = ncol_landscape,
      width = width,
      border_color = border_color, border_geometry = border_geometry,
      dpi = dpi,
      fileout = fileout)
  } else {
    stop("function to_miniature do support this type of file:", basename(filename))
  }

}

#' @export
#' @title Check if 'Microsoft Office' is available
#' @description The function test if 'Microsoft Office' is available.
#' @return a single logical value.
#' @examples
#' msoffice_available()
msoffice_available <- function() {
  (is_windows() || is_osx()) &&
    exec_available("word") &&
    exec_available("powerpoint")
}

pdf_to_miniature <- function(filename, row = NULL, ncol = NULL,
                             ncol_landscape = NULL, width = 650,
                             border_color = "#ccc", border_geometry = "2x2",
                             dpi = 150,
                             fileout = NULL) {
  img_list <- pdf_to_images(filename, dpi = dpi)
  x <- images_to_miniature(
    img_list = img_list,
    row = row, ncol = ncol, ncol_landscape = ncol_landscape,
    width = width * dpi/72,
    border_color = border_color, border_geometry = border_geometry
  )
  if(!is.null(fileout))
    image_write(x, path = fileout, format = "png")
  x
}

docx_to_miniature <- function(filename, row = NULL, ncol = NULL,
                              ncol_landscape = NULL, width = 650,
                              border_color = "#ccc", border_geometry = "2x2",
                              fileout = NULL, dpi = 150, timeout = 120) {
  pdf_filename <- tempfile(fileext = ".pdf")

  if (exec_available("word")) {
    docx2pdf(input = filename, output = pdf_filename)
  } else {
    to_pdf(input = filename, output = pdf_filename, timeout = timeout)
  }

  x <- pdf_to_miniature(pdf_filename,
    row = row, ncol = ncol, ncol_landscape = ncol_landscape,
    width = width, dpi = dpi,
    border_color = border_color, border_geometry = border_geometry
  )
  if(!is.null(fileout))
    image_write(x, path = fileout, format = "png")
  x
}

pptx_to_miniature <- function(filename, row = NULL, ncol = NULL,
                              ncol_landscape = NULL, width = 750,
                              border_color = "#ccc", border_geometry = "2x2",
                              dpi = 150,
                              fileout = NULL, timeout = 120) {
  pdf_filename <- tempfile(fileext = ".pdf")
  if (exec_available("powerpoint")) {
    pptx2pdf(input = filename, output = pdf_filename)
  } else {
    to_pdf(input = filename, output = pdf_filename, timeout = timeout)
  }

  x <- pdf_to_miniature(pdf_filename,
    row = row, ncol = ncol, ncol_landscape = ncol_landscape,
    width = width, dpi = dpi,
    border_color = border_color, border_geometry = border_geometry
  )
  if(!is.null(fileout))
    image_write(x, path = fileout, format = "png")
  x
}

htmlshot <- function(x, fileout = NULL, ...) {

  if (!requireNamespace("webshot2", quietly = TRUE)) {
    package_name <- "webshot2"
    stop(sprintf(
      "'%s' package should be installed to create a minature from an HTML file.",
      package_name)
    )
  }
  curr_wd <- getwd()
  path <- absolute_path(x)

  setwd(dirname(path))
  tf <- file(tempfile(fileext = ".txt"), "w")
  sink(tf)
  tryCatch(
    {
      webshot2::webshot(
        url = basename(path),
        file = fileout, ...
      )
    },
    finally = {
      sink()
      close(tf)
      setwd(curr_wd)
    }
  )

  fileout
}
