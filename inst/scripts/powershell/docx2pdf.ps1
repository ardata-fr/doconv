$indocx = "%s"
$opdf = "%s"

$Word = New-Object -ComObject Word.Application
$Word.Visible = $False
try {
  $doc = $Word.Documents.Open($indocx)

  $doc.Fields.Update()
  foreach ($toc in $doc.TablesOfContents) { $toc.Update() }
  foreach ($toc in $doc.TablesOfFigures) { $toc.Update() }

  $doc.ExportAsFixedFormat(
    $opdf,
    [Microsoft.Office.Interop.Word.WdExportFormat]::wdExportFormatPDF,
    $false,
    [Microsoft.Office.Interop.Word.WdExportOptimizeFor]::wdExportOptimizeForPrint,
    [Microsoft.Office.Interop.Word.WdExportRange]::wdExportAllDocument,
    0, 0,
    [Microsoft.Office.Interop.Word.WdExportItem]::wdExportDocumentContent,
    $true, $false,
    [Microsoft.Office.Interop.Word.WdExportCreateBookmarks]::wdExportCreateWordBookmarks,
    $true, $false, $true
  )

  $doc.Close([Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges)
} finally {
  $Word.Quit()
}
