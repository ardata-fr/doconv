#' @importFrom tools R_user_dir
#' @export
#' @title manage docx2pdf working directory
#' @description Initialize or remove working directory used
#' when docx2pdf create the PDF.
#'
#' On 'macOS', the operation require writing rights to the directory by the Word or PowerPoint program.
#' Word or PowerPoint program must be authorized
#' to write in the directories, if the authorization does not exist, a
#' manual confirmation window is launched, thus preventing automation.
#'
#' Fortunately, users only have to do this once. The package implementation use
#' only one directory where results are saved in order to have only one
#' time to click this confirmation.
#'
#' This directory is managed by R function [R_user_dir()]. Its value can be
#' read with the `working_directory()` function. The directory can be
#' deleted with `rm_working_directory()` and created with `init_working_directory()`.
#' Each call will remove that directory when completed.
#'
#' As a user, you do not have to use these functions because they are called
#' automatically by the `docx2pdf()` function. They are provided to meet
#' the requirements of CRAN policy:
#'
#' *"\[...\] packages may store user-specific data, configuration and cache files in their respective user
#' directories \[...\], provided that by default sizes are kept as small as possible
#' and the contents are actively managed (including removing outdated material)."*
working_directory <- function(){
  dir <- R_user_dir(package = "doconv", which = "data")
  absolute_path(dir)
}

#' @export
#' @rdname working_directory
rm_working_directory <- function(){
  dir <- working_directory()
  unlink(dir, recursive = TRUE, force = TRUE)
}

#' @export
#' @rdname working_directory
init_working_directory <- function(){
  rm_working_directory()
  dir <- working_directory()
  dir.create(dir, showWarnings = FALSE, recursive = TRUE)
  dir
}

absolute_path <- function(x){

  if (length(x) != 1L)
    stop("'x' must be a single character string")
  epath <- path.expand(x)

  epath <- normalizePath(epath, "/", mustWork = file.exists(epath))
  epath
}

# Escape a file path for a PowerShell double-quoted string.
# Backtick is the PS escape character; $, ", and ` need escaping.
escape_path_ps <- function(path) {
  path <- gsub("/", "\\", path, fixed = TRUE)
  path <- gsub("`", "``", path, fixed = TRUE)
  path <- gsub('"', '`"', path, fixed = TRUE)
  path <- gsub("$", "`$", path, fixed = TRUE)
  path
}

# Escape a file path for an AppleScript double-quoted string.
escape_path_applescript <- function(path) {
  path <- gsub("\\", "\\\\", path, fixed = TRUE)
  path <- gsub('"', '\\"', path, fixed = TRUE)
  path
}

#' @noRd
#' @importFrom pdftools pdf_convert pdf_length
#' @title Convert a PDF document to images
#' @description Convert a pdf file to a list of images (magick images).
#' @param file the pdf file
#' @param dpi resolution (dots per inch) to render, see [pdf_convert()].
#' @examples
#' infile <- tempfile(fileext = ".pdf")
#' pdf(file = infile, width = 8.3, height = 11.7)
#' barplot(1:10, col = 1:10)
#' barplot(1:10, col = 11:20)
#' dev.off()
#'
#' pdf_to_images(file = infile)
pdf_to_images <- function(file, dpi = 150) {

  if(!file.exists(file)){
    stop("could not find ", shQuote(file, type = "sh"))
  }
  dest <- tempfile()
  dir.create(dest)

  png_file <- gsub("\\.pdf$", "_%d.%s", basename(file))

  screen_copies <- pdf_convert(
    pdf = file, format = "png", verbose = FALSE,
    dpi = dpi,
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
  } else if(length(img_list) != length(row)){
    stop("The length of the `row` argument (", length(row), ") does not correspond to the ",
         "number of images (", length(img_list), ") to be processed.")
  }
  img_list <- img_list[row > 0]
  row <- row[row > 0]

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

# get PS exec policy
ps_execution_policy <- function() {
  res <- run("powershell", args = c("Get-ExecutionPolicy"), error_on_status = FALSE)
  policy <- gsub("\\s+", "", res$stdout)
  policy
}


# check if policy allows PS script execution. If not advise user.
stop_on_wrong_ps_exec_policy <- function() {
  policy <- ps_execution_policy()
  stop("Conversion failed.\n",
    "Your PowerShell execution policy '", policy, "' prevents script execution.\n",
    "Run one of the following commands in PS as admin to enable script execution: \n",
    "  `Set-ExecutionPolicy Bypass`\n",
    "  `Set-ExecutionPolicy Unrestricted`\n",
    call. = FALSE
  )
}
