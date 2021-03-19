#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.



Run "Report.pptx"

Sleep, 3000
Send {Left}
Sleep, 3000
Send {Enter}
Sleep, 8000
Send {Alt}
Sleep, 1000
Send {f}
Sleep, 1000
Send {s}
Sleep, 1000
Send !{F4}

