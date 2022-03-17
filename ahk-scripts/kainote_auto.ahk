#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

if A_Args.Length() < 1
{
    MsgBox % "This script requires at least 1 parameters but it only received " A_Args.Length() "."
    ExitApp
}

MsgBox % A_Args[1] " is set as a path to working directory."

WorkDir := A_Args[1]
_WorkDir := WorkDir "\*.mkv" ; Work only with .mkv files
Array := [] ; Initialize to be blank
FileList := ""  ; Initialize to be blank
Loop, %_WorkDir% 
    FileList .= A_LoopFileName "`n"
Sort, FileList

x := 01
Ep := ; Initialize to be blank
Loop, parse, FileList, `n
{
    if (A_LoopField = "")  ; Ignore the blank item at the end of the list
        continue
    FileName := A_LoopField
    Ep := SubStr("0" x, -1)
    if ( RegExMatch(FileName, "(?<![0-9])" Ep "(?![0-9])") != 0)
        Array.Push(FileName)
    x++
}

if Array.Length() < 1
{
    MsgBox, No files detected! Is path set correctly?
    ExitApp
}


for index, element in Array
{
    ;MsgBox % "Element number" . index . "is" . element
    ;MsgBox, 4,, Search pattern: (?<![0-9]) %Ep% (?![0-9]). Element number %index% is %element%.  Continue?
    ;IfMsgBox, No
    ;    ExitApp
}

Esc::ExitApp ; Press ESC to stop script
^j:: ; Press Ctrl+J to start script

i := 1 ; Start with first element of Files Array
FileName := 
Loop, % Array.MaxIndex() ; Loop until last element of Files Array
{ 
    FileName := Array[i]
    Sleep, 1000
    WinWaitActive, ahk_class Kainote_main_windowNR ; Wait for Kainote
    Send, ^+o ; Open video (Ctrl+Shift+O)
    WinWaitActive, Choose video file ; Wait for "Choose video file" dialog to pop-up
    Send, ^l ; Windows Explorer shortcut (CTRL+L) for selecting address bar
    SendRaw, %WorkDir% ; Send working dir path to address bar
    Sleep, 500
    Send, {enter} ; Open directory
    ; MouseClick, left, 320, 480
    ControlFocus, ComboBox1, Choose video file ; Select File Name bar
    SendRaw, %FileName% ; Send current filename to File Name bar
    Sleep, 500
    Send, {enter} ; Open file
    WinWaitActive, Confirmation ; Wait for "Confirmation" for loading related subtitles dialog
        Sleep, 500
        Send, {enter}
    WinWaitActive, Incompatible resolution,, 2 ; Wait 2s for "Incompatible resolution" dialog
    ;Control, Check,, Incompatible resolution ;Check "Change only subtitles resolution"
    Control, Check,, wxWindowNR3, Incompatible resolution ; Check "Resample subtitles (no stretch)"
    ControlClick, wxWindowNR4, Incompatible resolution ; Click "Change" button
    Loop ; Loop until "Style Manager" is opened, so we know when video indexing ends
    {
          If WinExist("Style manager")
        {
            Break
        }
        Send, ^m ; Open "Style Manager" (Ctrl+M)
        Sleep, 1000
        ;SoundPlay, C:\Windows\media\tada.wav, 1
    }
    ControlClick, wxWindowNR75, Style manager ; Click "Close" button on "Style Manager" panel
    WinWaitActive, ahk_class Kainote_main_windowNR
    MouseClick, left, 122, 662 ; Select randomly one subtitle line
    Send, ^a ; Select all lines
    Sleep, 500
    Send, ^] ; Open Automation - ASSWipe by custom shortcut (Ctrl+])
    WinWaitActive, ASSWipe ; Wait for "ASSwipe" panel to pop-up
    ControlClick, wxWindowNR26, ASSWipe ; Click "OK" button on "ASSwipe" panel
    Sleep, 500
    WinWaitActive, ahk_class Kainote_main_windowNR
    Send, ^p ; Run post processor by custom shortcut (Ctrl+P)
    Sleep, 1000
    Send, ^s ; Save changes (Ctrl+P)
    i++
}


