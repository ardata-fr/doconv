#' @importFrom tools R_user_dir
working_directory <- function(){
  dir <- R_user_dir(package = "doconv", which = "data")
  file.path(dir, "tmp_convert")
}

rm_working_directory <- function(force = TRUE){
  dir <- working_directory()
  unlink(dir, recursive = TRUE, force = TRUE)
}

init_working_directory <- function(force = TRUE){
  dir <- working_directory()
  if(force) rm_working_directory()
  if(!dir.exists(dir))
    dir.create(dir, showWarnings = FALSE, recursive = TRUE)
  dir
}

absolute_path <- function(x) {
  if (length(x) != 1L) {
    stop("'x' must be a single character string")
  }
  epath <- path.expand(x)

  if (file.exists(epath)) {
    epath <- normalizePath(epath, "/", mustWork = TRUE)
  } else {
    if (!dir.exists(dirname(epath))) {
      stop("directory ", x, " does not exist.", call. = FALSE)
    }
    cat("", file = epath)
    epath <- normalizePath(epath, "/", mustWork = TRUE)
    unlink(epath)
  }
  epath
}

#' @noRd
#' @importFrom pdftools pdf_convert
#' @title Convert a PDF document to images
#' @description Convert a pdf file to a list of images (magick images).
#' @param file the pdf file
#' @examples
#' infile <- tempfile(fileext = ".pdf")
#' pdf(file = infile, width = 8.3, height = 11.7)
#' barplot(1:10, col = 1:10)
#' barplot(1:10, col = 11:20)
#' dev.off()
#'
#' pdf_to_images(file = infile)
pdf_to_images <- function(file) {

  if(!file.exists(file)){
    stop("could not find ", shQuote(file, type = "sh"))
  }
  dest <- tempfile()
  dir.create(dest)

  png_file <- gsub("\\.pdf$", "%03d.png", basename(file))

  screen_copies <- pdf_convert(
    pdf = file, format = "png", verbose = FALSE,
    filenames = file.path(dest, png_file)
  )
  img_src <- list.files(dest, pattern = "\\.png$", full.names = TRUE)
  img_list <- lapply(img_src, image_read)
  img_list
}


#' @import magick
#' @noRd
#' @title Convert a set of images to a single png miniature
#' @description Convert a set of images to a png file
#' where pages are arranged in a layout.
#' @param img_list a list of magick image objects
#' @param row row index for every pages
#' @param width width of a single image
#' @param border_color border color
#' @param border_geometry border geometry to be added around images
#' @param fileout is not NULL image is saved to fileout
images_to_miniature <- function(img_list, row = NULL, width = 650,
                                border_color = "#ccc", border_geometry = "2x2",
                                fileout = NULL) {

  if (is.null(row)) {
    row <- seq_along(img_list)
  }

  geometry <- sprintf("%.0fx", max(table(row)) * width)

  img_list <- lapply(img_list, image_border, color = border_color, geometry = border_geometry)

  img_list <- split(img_list, row)
  img_stack <- list()
  for (row_imgs in img_list) {
    imginfo <- do.call(rbind, lapply(row_imgs, image_info))
    z <- lapply(row_imgs, image_extent,
                geometry = paste(
                  max(imginfo$width),
                  max(imginfo$height),
                  sep = "x"
                )
    )
    temp <- do.call(c, z)
    temp <- image_append(temp)
    temp <- image_extent(temp, geometry = geometry)
    img_stack[[length(img_stack) + 1]] <- temp
  }

  img_stack <- do.call(c, img_stack)
  img_stack <- image_append(img_stack, stack = TRUE)

  if(!is.null(fileout)){
    image_write(img_stack, path = fileout)
  }

  img_stack
}

