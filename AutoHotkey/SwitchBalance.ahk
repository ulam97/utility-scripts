#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#Include C:\Program Files\AutoHotkey\lib\VA.ahk

; Get the volume of the first and second channels.
volume1 := VA_GetMasterVolume(1)
volume2 := VA_GetMasterVolume(2)

bound := 2
LBVol := volume2 - bound
UBVol := volume2 + bound

if (volume1 > LBVol and volume1 < UBVol)
{
	LeftChSet := 10.0
	RightChSet := 20.0
	VA_SetMasterVolume(LeftChSet, 1)
	VA_SetMasterVolume(RightChSet, 2)
	MsgBox, % "Left channel set to: " LeftChSet
        . "`n  Right channel set to: " RightChSet
} else
{
	BothChSet := 10.0
	VA_SetMasterVolume(BothChSet, 1)
	VA_SetMasterVolume(BothChSet, 2)
	MsgBox, % "Both channels set to: " . BothChSet
}