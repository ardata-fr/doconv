#' @import magick
#' @export
#' @title Thumbnail of a document
#' @description Convert a file to an image (magick image)
#' where pages are arranged in a layout. Result can be saved
#' to a png file.
#' @param filename input file
#' @param row row index for every pages
#' @param width width of a single image
#' @param border_color border color
#' @param border_geometry border geometry to be added around images
#' @param fileout if not NULL, result is saved in a png file whose filename
#' is defined by this argument.
#' @return a magick image object as returned by [image_read()].
#' @examples
#' library(locatexec)
#' docx_file <- system.file(
#'   package = "doconv",
#'   "doc-examples/bookdown.docx"
#' )
#' if(exec_available("python") && docx2pdf_available())
#'   to_miniature(docx_file)
#'
#' pptx_file <- system.file(
#'   package = "doconv",
#'   "doc-examples/first_example.pptx"
#' )
#' if(exec_available("libreoffice"))
#'   to_miniature(pptx_file)
to_miniature <- function(filename, row = NULL, width = NULL,
                         border_color = "#ccc", border_geometry = "2x2",
                         fileout = NULL) {

  if (!file.exists(filename)) {
    stop("filename does not exist")
  }

  if(grepl("\\.(ppt|pptx)$", filename)){
    if(is.null(width)) width <- 750
    pptx_to_miniature(
      filename, row = row, width = width,
      border_color = border_color, border_geometry = border_geometry,
      fileout = fileout)
  } else if(grepl("\\.(doc|docx)$", filename)){
    if(is.null(width)) width <- 650
    docx_to_miniature(
      filename, row = row, width = width,
      border_color = border_color, border_geometry = border_geometry,
      fileout = fileout)
  } else if(grepl("\\.pdf$", filename)){
    if(is.null(width)) width <- 650
    pdf_to_miniature(
      filename, row = row, width = width,
      border_color = border_color, border_geometry = border_geometry,
      fileout = fileout)
  } else {
    stop("function to_miniature do support this type of file:", basename(filename))
  }

}

pdf_to_miniature <- function(filename, row = NULL, width = 650,
                             border_color = "#ccc", border_geometry = "2x2",
                             fileout = NULL) {
  img_list <- pdf_to_images(filename)
  x <- images_to_miniature(
    img_list = img_list,
    row = row, width = width,
    border_color = border_color, border_geometry = border_geometry
  )
  if(!is.null(fileout))
    image_write(x, path = fileout, format = "png")
  x
}

docx_to_miniature <- function(filename, row = NULL, width = 650,
                              border_color = "#ccc", border_geometry = "2x2",
                              fileout = NULL) {
  pdf_filename <- tempfile(fileext = ".pdf")
  docx2pdf(input = filename, output = pdf_filename)
  x <- pdf_to_miniature(pdf_filename,
    row = row, width = width,
    border_color = border_color, border_geometry = border_geometry
  )
  if(!is.null(fileout))
    image_write(x, path = fileout, format = "png")
  x
}

pptx_to_miniature <- function(filename, row = NULL, width = 750,
                              border_color = "#ccc", border_geometry = "2x2",
                              fileout = NULL) {
  pdf_filename <- tempfile(fileext = ".pdf")
  to_pdf(input = filename, output = pdf_filename)
  x <- pdf_to_miniature(pdf_filename,
    row = row, width = width,
    border_color = border_color, border_geometry = border_geometry
  )
  if(!is.null(fileout))
    image_write(x, path = fileout, format = "png")
  x
}

