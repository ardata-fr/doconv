set indocx to POSIX file "%s"
set opdf to POSIX file "%s"
set showMarkup to %s

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
    -- the PDF shows tracked changes and comments as the window does
    set theView to view of active window of theActiveDoc
    set show revisions and comments of theView to showMarkup
    set revisions view of theView to revisions view final
    save as theActiveDoc file format format PDF file name opdf
    close active window of theActiveDoc saving no
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
