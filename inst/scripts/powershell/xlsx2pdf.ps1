$infile = "%s"
$outfile = "%s"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  $wb = $excel.Workbooks.Open($infile)
  $wb.ExportAsFixedFormat(
    [Microsoft.Office.Interop.Excel.XlFixedFormatType]::xlTypePDF,
    $outfile
  )
  $wb.Close($false)
} finally {
  $excel.Quit()
}
