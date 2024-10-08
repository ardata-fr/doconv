load_package <- function(z) {
  if (!requireNamespace(z)) {
    stop("package '", z,"' is required to run that function")
  }
  suppressPackageStartupMessages(require(z, character.only = TRUE))
}

#' @title Visual test for document
#' @description This expectation can be used with 'tinytest' and 'testthat'
#' to check if a current document of type pdf, docx, doc, rtf, pptx or png
#' matches a target document. When the expectation is checked
#' for the first time, the expectation fails and a target miniature
#' of the document is saved in a folder named `_tinytest_doconv` or
#' `_snaps`.
#' @param name a string to identify the test. Each document in the test suite must have a unique name.
#' @param x file path of a document
#' @param tolerance the ratio of different pixels that is acceptable before triggering a failure.
#' @param engine test package being used in the test suite, one of "tinytest" or "testthat".
#' @return A [tinytest::tinytest()] or a [testthat::expect_snapshot_file] object.
#' @export
#' @examples
#' file <- system.file(package = "doconv",
#'   "doc-examples/example.docx")
#' \dontrun{
#' if (require("tinytest") && msoffice_available()){
#'   # first run add a new snapshot
#'   expect_snapshot_doc(x = file, name = "docx file", engine = "tinytest")
#'   # next runs compare with the snapshot
#'   expect_snapshot_doc(x = file, name = "docx file", engine = "tinytest")
#'
#'   # cleaning directory
#'   unlink("_tinytest_doconv", recursive = TRUE, force = TRUE)
#' }
#' if (require("testthat") && msoffice_available()){
#'   local_edition(3)
#'   # first run add a new snapshot
#'   expect_snapshot_doc(x = file, name = "docx file", engine = "testthat")
#'   # next runs compare with the snapshot
#'   expect_snapshot_doc(x = file, name = "docx file", engine = "testthat")
#' }
#' }
expect_snapshot_doc <- function(
    name,
    x,
    tolerance = 0.001,
    engine = c("tinytest", "testthat")
    ) {

  engine <- match.arg(engine)
  load_package(engine)

  if (inherits(x, "rdocx")) {
    x <- print(x, target = tempfile(fileext = ".docx"))
  } else if (inherits(x, "rpptx")) {
    x <- print(x, target = tempfile(fileext = ".pptx"))
  } else if (inherits(x, "rtf")) {
    x <- print(x, target = tempfile(fileext = ".rtf"))
  }

  if ("testthat" %in% engine) {
    expect_snapshot_testthat(
      name = name, x = x,
      tolerance = tolerance)
  } else {
    expect_snapshot_tinytest(
      name = name, current = x,
      tolerance = tolerance)
  }

}


#' @title Visual test for an HTML document
#' @description This expectation can be used with 'tinytest' and 'testthat'
#' to check if a current document of type HTML
#' matches a target document. When the expectation is checked
#' for the first time, the expectation fails and a target miniature
#' of the document is saved in a folder named `_tinytest_doconv` or
#' `_snaps`.
#' @param name a string to identify the test. Each document in the test suite must have a unique name.
#' @param x file path of an HTML document
#' @param tolerance the ratio of different pixels that is acceptable before triggering a failure.
#' @param engine test package being used in the test suite, one of "tinytest" or "testthat".
#' @param ... arguments used by `webshot::webshot2()`.
#' @return A [tinytest::tinytest()] or a [testthat::expect_snapshot_file] object.
#' @export
#' @examples
#' file <- tempfile(fileext = ".html")
#' html <- paste0("<html><head><title>hello</title></head>",
#'        "<body><h1>Hello World</h1></body></html>\n")
#' cat(html, file = file)
#'
#' \dontrun{
#' if (require("tinytest") && require("webshot2")){
#'   # first run add a new snapshot
#'   expect_snapshot_html(x = file, name = "html file",
#'     engine = "tinytest")
#'   # next runs compare with the snapshot
#'   expect_snapshot_html(x = file, name = "html file",
#'     engine = "tinytest")
#'
#'   # cleaning directory
#'   unlink("_tinytest_doconv", recursive = TRUE,
#'     force = TRUE)
#' }
#' if (require("testthat") && require("webshot2")){
#'   local_edition(3)
#'   # first run add a new snapshot
#'   expect_snapshot_html(x = file, name = "html file",
#'     engine = "testthat")
#'   # next runs compare with the snapshot
#'   expect_snapshot_html(x = file, name = "html file",
#'     engine = "testthat")
#' }
#' }
expect_snapshot_html <- function(
    name,
    x,
    tolerance = 0.001,
    engine = c("tinytest", "testthat"),
    ...
    ) {

  engine <- match.arg(engine)
  load_package(engine)
  x <- htmlshot(x, fileout = tempfile(fileext = ".png"), ...)

  if ("testthat" %in% engine) {
    expect_snapshot_testthat(
      name = name, x = x,
      tolerance = tolerance)
  } else {
    expect_snapshot_tinytest(
      name = name, current = x,
      tolerance = tolerance)
  }

}


expect_snapshot_testthat <- function(name, x, tolerance = 0.001) {
  name <- paste0(name, ".png")
  testthat::announce_snapshot_file(name = name)
  png_out <- tempfile(fileext = ".png")

  file_type <- gsub("(.*)\\.(pdf|docx|pptx|rtf|png)$", "\\2", x)
  exec <- guess_exec(file_type)

  if (!is.null(exec) && !exec_available(exec)) {
    return(
      testthat::expect_snapshot_file(
        png_out, name,
        compare = function(ref, new) {
          FALSE
        }
      )
    )
  }

  if (file_type %in% c("pdf", "docx", "doc", "rtf")) {
    to_miniature(x, fileout = png_out, width = 1000)
  } else if ("pptx" %in% file_type) {
    to_miniature(x, fileout = png_out, width = 1200)
  } else if ("png" %in% file_type) {
    file.copy(x, png_out, overwrite = TRUE)
  } else {
    stop("`expect_snapshot_doc()` only supports docx, pptx, rtf, pdf and png files.")
  }

  testthat::expect_snapshot_file(
    path = png_out, name = name,
    compare = function(ref, new) {
      abs(diff_image(ref, new)) < tolerance
    }
  )
}

expect_snapshot_tinytest <- function(current, name, tolerance = 0.001) {
  file_type <- gsub("(.*)\\.(pdf|docx|pptx|rtf|png)$", "\\2", current)
  exec <- guess_exec(file_type)

  if (!is.null(exec) && !exec_available(exec)) {
    return(
      tinytest::tinytest(
        result = TRUE,
        call = sys.call(sys.parent(1)),
        diff = "NA",
        info = paste0(
          shQuote(exec),
          " must be available to run this function"
        )
      )
    )
  }

  if (file_type %in% c("pdf", "doc", "docx", "rtf")) {
    expect_office_doc_diff_tinytest(
      current,
      name,
      tolerance = tolerance,
      width = 1000
    )
  } else if ("pptx" %in% file_type) {
    expect_office_doc_diff_tinytest(
      current,
      name,
      tolerance = tolerance,
      width = 1200
    )
  } else if (grepl("\\.png$", current)) {
    expect_png_diff(current, name, tolerance = tolerance)
  } else {
    stop("`expect_snapshot_tinytest()` only supports docx, doc, pptx, rtf, pdf and png files.")
  }
}

expect_office_doc_diff_tinytest <- function(current,
                                            name,
                                            tolerance = sqrt(.Machine$double.eps),
                                            width = 1000) {
  # portable test names
  name <- gsub(" ", "_", name)

  tmp <- tempfile()
  tmp_current <- file.path(tmp, "current")

  dir.create(tinytest_dir, showWarnings = FALSE)
  dir.create(tmp_current, showWarnings = FALSE, recursive = TRUE)

  current_miniature <- file.path(tmp_current, paste0(name, ".png"))
  to_miniature(current, fileout = current_miniature, width = width)

  target_miniature <- file.path(tinytest_dir, basename(current_miniature))
  if (!file.exists(target_miniature)) {
    file.copy(current_miniature, target_miniature, overwrite = TRUE)
    msg <- sprintf("new miniature was saved to: %s", target_miniature)
    flag <- TRUE
    results <- "100%"
  } else {
    results <- diff_image(img1 = target_miniature, img2 = current_miniature)
    if (results < 0) {
      msg <- "difference detected - unequal dimensions"
      results <- "N/A"
      flag <- FALSE
    } else if (results < tolerance) {
      msg <- "no difference detected"
      results <- "0%"
      flag <- TRUE
    } else {
      flag <- FALSE
      msg <- "difference detected - miniature changed"
      results <- paste0(formatC(results * 100, format = "f", digits = 2), "%")
    }
  }

  tinytest::tinytest(
    result = flag,
    call = sys.call(sys.parent(2)),
    diff = as.character(results),
    info = msg
  )
}

expect_png_diff <- function(current,
                            name,
                            tolerance = sqrt(.Machine$double.eps)) {
  # portable test names
  name <- gsub(" ", "_", name)

  tmp <- tempfile()
  tmp_current <- file.path(tmp, "current")

  dir.create(tinytest_dir, showWarnings = FALSE)
  dir.create(tmp_current, showWarnings = FALSE, recursive = TRUE)

  current_miniature <- file.path(tmp_current, paste0(name, ".png"))
  file.copy(current, current_miniature, overwrite = TRUE)

  target_miniature <- file.path(tinytest_dir, basename(current_miniature))

  if (!file.exists(target_miniature)) {
    file.copy(current_miniature, target_miniature, overwrite = TRUE)
    msg <- sprintf("new image was saved to: %s", target_miniature)
    flag <- TRUE
    results <- "0%"
  } else {
    results <- diff_image(img1 = target_miniature, img2 = current_miniature)
    if (results < 0) {
      msg <- "difference detected - unequal dimensions"
      results <- "??"
      flag <- FALSE
    } else if (results < tolerance) {
      msg <- "no difference detected"
      results <- "0%"
      flag <- TRUE
    } else {
      flag <- FALSE
      msg <- "difference detected - image changed"
      results <- formatC(results * 100, format = "f", digits = 2)
    }
  }

  tinytest::tinytest(
    result = flag,
    call = sys.call(sys.parent(2)),
    diff = as.character(results),
    info = msg
  )
}

# utils -----

guess_exec <- function(file_type) {
  exec <- NULL
  if (file_type %in% c("pptx", "docx", "doc", "rtf")) {
    exec <- c(
      "pptx" = "powerpoint",
      "docx" = "word",
      "doc" = "word",
      "rtf" = "word")[file_type]
  }
  exec
}

#' @importFrom magick image_raster image_read
diff_image <- function(img1, img2) {
  img.reference <- image_raster(image_read(img1), tidy = FALSE)
  img.empty <- image_raster(image_read(img2), tidy = FALSE)
  if (length(img.empty) != length(img.reference)) {
    return(-1.0)
  }
  sum(img.reference != img.empty) / length(img.empty)
}

tinytest_dir <- "_tinytest_doconv"

.onLoad = function(libname, pkgname) {
  if ("tinytest" %in% loadedNamespaces() && requireNamespace("tinytest")) {
    tinytest::register_tinytest_extension(
      "doconv",
      c("expect_snapshot_doc"))
  }
}
