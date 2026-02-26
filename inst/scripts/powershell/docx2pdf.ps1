$indocx = "%s"
$opdf = "%s"
$showMarkup = "%s"

$Word = New-Object -ComObject Word.Application
$Word.Visible = $False
try {
  $doc = $Word.Documents.Open($indocx)

  $doc.Fields.Update()
  foreach ($toc in $doc.TablesOfContents) { $toc.Update() }
  foreach ($toc in $doc.TablesOfFigures) { $toc.Update() }

  if ($showMarkup -eq "True") {
    $Word.ActiveWindow.View.ShowRevisionsAndComments = $true
    $Word.ActiveWindow.View.RevisionsView = [Microsoft.Office.Interop.Word.WdRevisionsView]::wdRevisionsViewFinal
    $exportItem = [Microsoft.Office.Interop.Word.WdExportItem]::wdExportDocumentWithMarkup
  } else {
    $exportItem = [Microsoft.Office.Interop.Word.WdExportItem]::wdExportDocumentContent
  }

  $doc.ExportAsFixedFormat(
    $opdf,
    [Microsoft.Office.Interop.Word.WdExportFormat]::wdExportFormatPDF,
    $false,
    [Microsoft.Office.Interop.Word.WdExportOptimizeFor]::wdExportOptimizeForPrint,
    [Microsoft.Office.Interop.Word.WdExportRange]::wdExportAllDocument,
    0, 0,
    $exportItem,
    $true, $false,
    [Microsoft.Office.Interop.Word.WdExportCreateBookmarks]::wdExportCreateWordBookmarks,
    $true, $false, $true
  )

  $doc.Close([Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges)
} finally {
  $Word.Quit()
}
