set indocx to POSIX file "%s"

set wasRunning to false

if application "Microsoft Word" is running then
  set wasRunning to true
end if

try
  tell application "Microsoft Word"
    set theActiveDoc to open file name indocx
    repeat with aField in (get fields of theActiveDoc)
      update field aField
    end repeat
    repeat with toc in (get tables of contents of theActiveDoc)
      update toc
    end repeat
    repeat with toc in (get tables of figures of theActiveDoc)
      update toc
    end repeat
    close active window of theActiveDoc saving yes
  end tell
on error errMsg
  try
    tell application "Microsoft Word" to close active window saving no
  end try
  if not wasRunning then
    tell application "Microsoft Word" to quit
  end if
  error errMsg
end try

if not wasRunning then
  tell application "Microsoft Word" to quit
end if
