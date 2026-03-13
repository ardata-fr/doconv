set infile to POSIX file "%s"
set opdf to POSIX file "%s"

set wasRunning to false

if application "Microsoft Excel" is running then
  set wasRunning to true
end if

try
  tell application "Microsoft Excel"
    open infile
    save active workbook in opdf as PDF file format
    close active workbook saving no
  end tell
on error errMsg
  try
    tell application "Microsoft Excel" to close active workbook saving no
  end try
  if not wasRunning then
    tell application "Microsoft Excel" to quit
  end if
  error errMsg
end try

if not wasRunning then
  tell application "Microsoft Excel" to quit
end if
