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

# Run `expr` while holding a lock shared by all R processes of the user.
# 'Word' and 'PowerPoint' run as a single instance: two conversions at
# the same time (e.g. parallel checks) use the same instance and one of
# them fails. They also share the same working directory.
with_office_lock <- function(expr, timeout = 600) {
  lock_dir <- R_user_dir(package = "doconv", which = "cache")
  dir.create(lock_dir, showWarnings = FALSE, recursive = TRUE)
  lock <- filelock::lock(file.path(lock_dir, "office.lock"),
                         timeout = timeout * 1000)
  if (is.null(lock)) {
    stop("could not get access to 'Microsoft Office' within ", timeout,
         " seconds, another R process is still using it.", call. = FALSE)
  }
  on.exit(filelock::unlock(lock))
  expr
}

# 'Rscript' of the running R session, not the first one found on the PATH
rscript_exec <- function() {
  file.path(R.home("bin"), if (is_windows()) "Rscript.exe" else "Rscript")
}

# Write a PowerShell script in UTF-8 with a BOM, Windows PowerShell 5.1
# reads files without BOM with the ANSI code page and non ASCII paths
# are then not found (#8).
write_ps1 <- function(x, path) {
  con <- file(path, open = "wb")
  on.exit(close(con))
  writeBin(as.raw(c(0xef, 0xbb, 0xbf)), con)
  writeLines(enc2utf8(x), con, sep = "\r\n", useBytes = TRUE)
  invisible(path)
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

  png_file <- gsub("\\.pdf$", "_%d.%s", basename(file), ignore.case = TRUE)

  # pdf_convert() returns the files in page order, list.files() would
  # sort them alphabetically (page 10 before page 2)
  screen_copies <- pdf_convert(
    pdf = file, format = "png", verbose = FALSE,
    dpi = dpi,
    filenames = file.path(dest, png_file)
  )
  img_list <- lapply(screen_copies, image_read)
  img_list
}


ggplot_to_png <- function(x, width = 9, height = 7, units = "in", res = 200) {
  if (!requireNamespace("ragg", quietly = TRUE)) {
    stop("Package 'ragg' is required for PNG rendering.")
  }
  if (inherits(x, "gg")) {
    x <- list(x)
  } else if (is.list(x)) {
    x <- Filter(function(el) inherits(el, "gg"), x)
  }
  vapply(x, function(obj) {
    path <- tempfile(fileext = ".png")
    ragg::agg_png(filename = path, width = width, height = height,
                  units = units, res = res)
    tryCatch(
      print(obj),
      error = function(e) {
        grid::grid.rect(gp = grid::gpar(col = NA, fill = "white"))
      },
      finally = grDevices::dev.off()
    )
    path
  }, character(1))
}

compute_row <- function(img_list, ncol, ncol_landscape = NULL) {
  n <- length(img_list)

  if (is.null(ncol_landscape)) {
    return(rep(seq_len(ceiling(n / ncol)), each = ncol, length.out = n))
  }

  orientations <- vapply(img_list, function(img) {
    info <- image_info(img)
    if (info$width > info$height) "landscape" else "portrait"
  }, character(1))

  row <- integer(n)
  current_row <- 1L
  col_in_row <- 0L
  current_orientation <- orientations[1]

  for (i in seq_len(n)) {
    max_cols <- if (orientations[i] == "landscape") ncol_landscape else ncol
    if (orientations[i] != current_orientation || col_in_row >= max_cols) {
      current_row <- current_row + 1L
      col_in_row <- 0L
      current_orientation <- orientations[i]
    }
    col_in_row <- col_in_row + 1L
    row[i] <- current_row
  }

  row
}

#' @import magick
#' @noRd
#' @title Convert a set of images to a single png miniature
#' @description Convert a set of images to a png file
#' where pages are arranged in a layout.
#' @param img_list a list of magick image objects
#' @param row row index for every pages
#' @param ncol number of pages per row
#' @param ncol_landscape number of landscape pages per row
#' @param width width of a single image
#' @param border_color border color
#' @param border_geometry border geometry to be added around images
#' @param fileout is not NULL image is saved to fileout
images_to_miniature <- function(img_list, row = NULL, ncol = NULL,
                                ncol_landscape = NULL, width = 650,
                                border_color = "#ccc", border_geometry = "2x2",
                                fileout = NULL) {

  if (!is.null(row) && !is.null(ncol)) {
    warning("`row` and `ncol` are both set, `ncol` is ignored.")
  }
  if (!is.null(ncol_landscape) && is.null(ncol)) {
    stop("`ncol_landscape` requires `ncol`.")
  }
  if (is.null(row) && !is.null(ncol)) {
    row <- compute_row(img_list, ncol = ncol, ncol_landscape = ncol_landscape)
  }

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

# tracked changes and comments ----

check_show_markup <- function(show_markup) {
  if (!is.logical(show_markup) || length(show_markup) != 1L || is.na(show_markup)) {
    stop("`show_markup` must be TRUE or FALSE.", call. = FALSE)
  }
  invisible(show_markup)
}

# 'LibreOffice' PDF filter. Comments are drawn in the margin with
# `show_markup = TRUE`, they are not exported otherwise.
libreoffice_pdf_filter <- function(input, show_markup = FALSE) {
  if (show_markup && grepl("\\.(docx|doc|rtf|odt)$", input, ignore.case = TRUE)) {
    'pdf:writer_pdf_Export:{"ExportNotesInMargin":{"type":"boolean","value":"true"}}'
  } else {
    "pdf:writer_pdf_Export"
  }
}

# 'LibreOffice' always shows tracked changes, unless the document settings
# ask for the final view (`w:revisionView`); 'Word' ignores this setting.
# Returns `input`, or a copy in `dir` with these settings when `input`
# contains tracked changes.
docx_hide_revisions <- function(input, dir) {
  if (!grepl("\\.docx$", input, ignore.case = TRUE)) {
    return(input)
  }
  unz <- tempfile()
  on.exit(unlink(unz, recursive = TRUE, force = TRUE))

  document <- utils::unzip(input, files = "word/document.xml", exdir = unz)
  if (!length(document)) {
    return(input)
  }
  xml <- paste(readLines(document, warn = FALSE, encoding = "UTF-8"), collapse = "\n")
  if (!grepl("<w:(ins|del|moveFrom|moveTo)[ >]", xml)) {
    return(input)
  }
  if (!requireNamespace("officer", quietly = TRUE)) {
    warning("package 'officer' is required to hide tracked changes with ",
            "'LibreOffice', they are shown in the PDF.", call. = FALSE)
    return(input)
  }

  unlink(unz, recursive = TRUE, force = TRUE)
  officer::unpack_folder(input, unz)
  settings_file <- file.path(unz, "word", "settings.xml")
  if (!file.exists(settings_file)) {
    return(input)
  }
  settings <- paste(readLines(settings_file, warn = FALSE, encoding = "UTF-8"), collapse = "\n")
  settings <- gsub("<w:revisionView[^>]*/>", "", settings)
  # first child instead of its schema position: only read by 'LibreOffice'
  settings <- sub(
    "(<w:settings[^>]*>)",
    "\\1<w:revisionView w:markup=\"false\" w:insDel=\"false\" w:formatting=\"false\"/>",
    settings
  )
  writeLines(settings, settings_file, useBytes = TRUE)

  output <- file.path(dir, basename(input))
  officer::pack_folder(unz, output)
  output
}

# 'LibreOffice' is run without the LD_LIBRARY_PATH set by R on Linux: R
# prepends its own library directories to it (see R_HOME/etc/ldpaths) and
# 'LibreOffice' then fails ("libreglo.so: cannot open shared object file")
# or aborts.
#' @importFrom locatexec is_unix
libreoffice_env <- function() {
  if (is_unix()) c("current", LD_LIBRARY_PATH = "") else "current"
}
