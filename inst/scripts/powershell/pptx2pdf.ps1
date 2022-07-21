$infile = "%s"
$outfile = "%s"

$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open($infile)

$opt= [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPDF
$pres.SaveAs($outfile, $opt)
$pres.Close()
$ppt.Quit()
