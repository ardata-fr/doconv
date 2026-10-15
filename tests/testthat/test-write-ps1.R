test_that("PowerShell scripts are written in UTF-8 with a BOM", {
  path <- tempfile(fileext = ".ps1")
  doconv:::write_ps1(c('$indocx = "C:\\tmp\\\u6d4b\u8bd5.docx"', "Write-Output $indocx"), path)
  bytes <- readBin(path, "raw", file.size(path))

  expect_identical(bytes[1:3], as.raw(c(0xef, 0xbb, 0xbf)))
  # the path is kept in UTF-8
  expect_identical(
    rawToChar(bytes[-(1:3)]),
    '$indocx = "C:\\tmp\\\u6d4b\u8bd5.docx"\r\nWrite-Output $indocx\r\n'
  )
})
