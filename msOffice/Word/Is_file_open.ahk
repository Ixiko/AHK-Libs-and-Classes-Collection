; By kon
; Taken from - https://github.com/ahkon/MS-Office-COM-Basics/blob/master/Examples/Word/Is_file_open.ahk
; License - https://github.com/ahkon/MS-Office-COM-Basics/blob/master/LICENSE

﻿FilePath := A_ScriptDir "\New Microsoft Word Document.docx"  ; Path to a Word document.
if FileOpen(FilePath, "rw") ; FileOpen fails if the file is already open.
    FileStatus := "not open"
else
    FileStatus := "already open"
MsgBox, 64, File Status, % "The file is " FileStatus ".`n`n(" FilePath ")"
