set indocx to POSIX file "%s"
set opdf to POSIX file "%s"

set myVariable to false

if application "Microsoft Word" is running then
  set myVariable to true
end if

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
	save as theActiveDoc file format format PDF file name opdf
	close active window of theActiveDoc saving no
end tell

if myVariable is equal to false then
  tell application "Microsoft Word" to quit
end if
