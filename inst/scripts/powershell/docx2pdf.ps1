$indocx = "%s"
$opdf = "%s"

$Word = New-Object -ComObject Word.Application
$doc = $Word.Documents.Open($indocx)

$doc.Fields.Update()
$tocs = $doc.TablesOfContents
foreach ($toc in $tocs)
{
  $toc.Update()
}
$tocs = $doc.TablesOfFigures
foreach ($toc in $tocs)
{
  $toc.Update()
}
$Word.Visible = $False

$result = $doc.ExportAsFixedFormat(
               $opdf,
               [Microsoft.Office.Interop.Word.WdExportFormat]::wdExportFormatPDF,
               $false,
               [Microsoft.Office.Interop.Word.WdExportOptimizeFor]::wdExportOptimizeForPrint,
               [Microsoft.Office.Interop.Word.WdExportRange]::wdExportAllDocument,
               0,
               0,
               [Microsoft.Office.Interop.Word.WdExportItem]::wdExportDocumentContent,
               $true,
               $false,
               [Microsoft.Office.Interop.Word.WdExportCreateBookmarks]::wdExportCreateWordBookmarks,
               $true,
               $false,
               $true
          )


$doc.Close()
$word.Quit()
