$indocx = "%s"

$Word = New-Object -ComObject Word.Application
$Word.Visible = $False
try {
  $doc = $Word.Documents.Open($indocx)

  $doc.Fields.Update()
  foreach ($toc in $doc.TablesOfContents) { $toc.Update() }
  foreach ($toc in $doc.TablesOfFigures) { $toc.Update() }

  $doc.Save()
  $doc.Close()
} finally {
  $Word.Quit()
}
