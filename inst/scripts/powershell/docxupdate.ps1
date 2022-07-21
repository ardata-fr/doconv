$indocx = "%s"

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

$result = $doc.Save()


$doc.Close()
$word.Quit()
