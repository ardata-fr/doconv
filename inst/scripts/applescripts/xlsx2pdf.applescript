set inxlsx to POSIX file "%s"
set opdf to POSIX file "%s"

set wasRunning to false

if application "Microsoft Excel" is running then
  set wasRunning to true
end if

try
  tell application "Microsoft Excel"
    open inxlsx
    set wb to active workbook
    -- one page wide per sheet; sheets must be addressed by index,
    -- 'repeat with ws in worksheets of wb' fails with error -50
    repeat with i from 1 to count of worksheets of wb
      set ps to page setup object of worksheet i of wb
      set zoom of ps to false
      set fit to pages wide of ps to 1
      -- no limit on the height: false and 0 are refused (error -2700)
      set fit to pages tall of ps to 9999
    end repeat
    save workbook as wb filename (opdf as text) file format PDF file format
    close wb saving no
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
