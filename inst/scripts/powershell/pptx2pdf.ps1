$infile = "%s"
$outfile = "%s"

$ppt = New-Object -ComObject PowerPoint.Application
try {
  $pres = $ppt.Presentations.Open($infile, $true, $false, $false)

  $opt = [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPDF
  $pres.SaveAs($outfile, $opt)
  $pres.Close()
} finally {
  $ppt.Quit()
}
