; Git Integration for TradeStation using AutoIt
;
; TS Version: 10 or later
; AutoIT Site: https://www.autoitscript.com/
; AutoIT Version: 3.3.16.1
;
; Credits: A big thanks to moscu for creating the original AutoIT backup
;		   script. Parts of that script were used here. 01/02/2006
;		   https://community.tradestation.com/Discussions/Topic.aspx?Topic_ID=44956
;
; Author: Snaggs
; Date	: 06/10/24
;
; History:
; ---------------------------------------------------------------------------
; Date		 Version	Task
; --------	 --------	-----------------------------------------------------
; 06/10/24	 1.0.1		Snaggs - Created script to export active study
; 06/11/25	 1.1.0		Snaggs - Unified script to work with Win10 and Win11
;								 Save file using AutoIt instead of Notepad
;								 to make it work across Win10 & Win11
;								 Reformatted and cleaned up the code
;								 Localized variables in functions
;								 Added error handing
; 12/27/25	1.2.0		Snaggs - Simplified the code
;                              - Added file saved dialog box

#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <StringConstants.au3>

HotKeySet("!`", "Export")		; Alt-`     Saves the file to disk
HotKeySet("^`", "EndScript")	; Ctrl-`    Exit's the script

Global $rootDir = @UserProfileDir & "\Documents\GitHub\TradeStation\"	; Must have trailing \
Global $tdeTitle = "TradeStation Development Environment"

MsgBox($MB_ICONINFORMATION, "Starting CodeExport Script", "Starting...", 2)

While 1
	Sleep(100)
WEnd

Func EndScript()
	MsgBox($MB_ICONINFORMATION, "Ending Script", "Exiting...", 2)
	Exit
EndFunc

Func Export()
	; Find the TradeStation Development Environment window
	WinActivate("[CLASS:TSDEV.EXE TRADESTATION]")

	; Get the text from the title bar
	Local $winTitle = WinGetTitle("[ACTIVE]")

	; Get a handle to the control that has the code in it, then get the code from it
	Local $hWnd = ControlGetFocus("[ACTIVE]")
	Local $winCode = ControlGetText("[ACTIVE]", "", $hWnd)

	; Copy the code to the clipboard
	ClipPut($winCode)

	; Locate the start and end of the study name
	Local $start = StringInStr($winTitle, "-", 0) + 2   ; +2 to get past the '- '
	Local $end = StringInStr($winTitle, ":", 2, -1)

	; Extract the study name from the Window Title
	Local $studyName = StringStripWS(StringMid($winTitle, $start, $end - $start -1), 3)

	; Replace invalid filename chars with an _
	$studyName = StringRegExpReplace($studyName, '[\\/:*?"<>|]', '_')

	; Get the study type (indicator, function, etc)
	Local $studyType = StringLower(StringMid($WinTitle, $end + 2))

	; If the study is unsaved, drop the * off the end of the studyType
	If StringInStr($studyType, "*") Then
		$studyType = StringTrimRight($studyType, 1)
    ElseIf StringInStr($studyType, "(read-only)") Then
        $studyType = StringTrimRight($studyType, 12)
    EndIf

    ; Build the directory based on studyType
	Local $subDir = ""
	Switch $studyType
		Case "activitybar"
			$subDir = "ActivityBars"
		Case "function"
			$subDir = "Functions"
		Case "indicator"
			$subDir = "Indicators"
		Case "paintbar"
			$subDir = "PaintBars"
		Case "probabilitymap"
			$subDir = "ProbabilityMaps"
		Case "showme"
			$subDir = "ShowMes"
		Case "strategy"
			$subDir = "Strategies"
		Case "trading application"
			$subDir = "TradingApps"
		Case Else
			MsgBox($MB_OK, "Unknown Study Type", "Study type unknown: " & $studyType)
			Return
	EndSwitch

	; Add the trailing slash
	Local $targetDir = $rootDir & $subDir & "\"

	; If the directory doesn't exist, create it
	If Not FileExists($targetDir) Then DirCreate($targetDir)

	; Build the full filename
	Local $fullFilePath = $targetDir & $studyName & ".txt"

	; Save the contents to the file
	SaveFile($fullFilePath, $studyName)
EndFunc

; Save the the clipboard text to a file.
Func SaveFile($filePath, $studyName)
	; Get the text to save from the Clipboard
	Local $text = ClipGet()
	If @error Or $text = "" Then
		MsgBox(16, "Clipboard Error", "No code found in clipboard.")
		Return
	EndIf

	; Open the studyFile
	Local $fh = FileOpen($filePath, $FO_OVERWRITE)
	If $fh = -1 Then
		MsgBox(16, "File Error", "Failed to open file: " & $filePath)
		Return
	EndIf

	; Write the contents to the file and close it
	FileWrite($fh, $text)
	FileClose($fh)

    ; Notify the user the study was saved to a file
    MsgBox($MB_ICONINFORMATION, "File Saved", $studyName, 2)
EndFunc
