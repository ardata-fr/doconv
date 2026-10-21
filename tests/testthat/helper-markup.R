# a docx with a comment, a tracked deletion and a tracked insertion
markup_docx <- function(path = tempfile(fileext = ".docx")) {
  doc <- officer::read_docx()
  doc <- officer::body_add_fpar(doc, officer::fpar(
    officer::ftext("Paragraph with a comment: "),
    officer::run_comment(
      cmt = officer::block_list(officer::fpar("THE COMMENT TEXT")),
      run = officer::ftext("commented words"), author = "Reviewer"
    )
  ))
  doc <- officer::body_add_par(doc, "MARKERPARA")
  base <- tempfile(fileext = ".docx")
  print(doc, target = base)

  unz <- tempfile()
  officer::unpack_folder(base, unz)
  document <- file.path(unz, "word", "document.xml")
  xml <- paste(readLines(document, warn = FALSE, encoding = "UTF-8"), collapse = "\n")
  xml <- sub(
    "<w:r>\\s*<w:t[^>]*>MARKERPARA</w:t>\\s*</w:r>",
    paste0(
      '<w:r><w:t xml:space="preserve">Kept text </w:t></w:r>',
      '<w:del w:id="901" w:author="Reviewer" w:date="2026-10-21T00:00:00Z">',
      "<w:r><w:delText>DELETEDTEXT</w:delText></w:r></w:del>",
      '<w:ins w:id="902" w:author="Reviewer" w:date="2026-10-21T00:00:00Z">',
      "<w:r><w:t>INSERTEDTEXT</w:t></w:r></w:ins>"
    ),
    xml
  )
  writeLines(xml, document, useBytes = TRUE)
  officer::pack_folder(unz, path)
  unlink(c(base, unz), recursive = TRUE)
  path
}

pdf_text_of <- function(pdf) paste(pdftools::pdf_text(pdf), collapse = " ")
