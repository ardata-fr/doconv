set indocx to POSIX file "%s"
set opdf to POSIX file "%s"

set wasRunning to false

if application "Microsoft PowerPoint" is running then
  set wasRunning to true
end if

try
  tell application "Microsoft PowerPoint"
    open indocx
    save active presentation in opdf as save as PDF
    close active presentation
  end tell
on error errMsg
  try
    tell application "Microsoft PowerPoint" to close active presentation
  end try
  if not wasRunning then
    tell application "Microsoft PowerPoint" to quit
  end if
  error errMsg
end try

if not wasRunning then
  tell application "Microsoft PowerPoint" to quit
end if
