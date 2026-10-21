$infile = "%s"
$outfile = "%s"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $False
$excel.DisplayAlerts = $False
try {
  # read only, the page setup changes below are not saved
  $wb = $excel.Workbooks.Open($infile, 0, $true)

  # one page wide per sheet
  foreach ($ws in $wb.Worksheets) {
    $ws.PageSetup.Zoom = $false
    $ws.PageSetup.FitToPagesWide = 1
    $ws.PageSetup.FitToPagesTall = $false
  }

  $xlTypePDF = 0
  $wb.ExportAsFixedFormat($xlTypePDF, $outfile)
  $wb.Close($false)
} finally {
  $excel.Quit()
}
