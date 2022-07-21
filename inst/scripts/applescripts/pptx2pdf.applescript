set indocx to POSIX file "%s"
set opdf to POSIX file "%s"

set myVariable to false

if application "Microsoft PowerPoint" is running then
  set myVariable to true
end if

tell application "Microsoft PowerPoint"
	open indocx
	save active presentation in opdf as save as PDF
	close active presentation
end tell

if myVariable is equal to false then
  tell application "Microsoft PowerPoint" to quit
end if
