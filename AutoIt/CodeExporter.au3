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
;								 Localized variables in func's
;								 Added error handing

#include <MsgBoxConstants.au3>
#include <StringConstants.au3>

HotKeySet("!`", "Export")		; Alt-`
HotKeySet("^`", "EndScript")	; Ctrl-`

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
	WinActivate("[CLASS:TSDEV.EXE TRADESTATION]")

	; Get the text from the title bar
	Local $winTitle = WinGetTitle("[ACTIVE]")

	; Get a handle to the control that has the code in it, then get the code from it
	Local $hWnd = ControlGetFocus("[ACTIVE]")
	Local $winCode = ControlGetText("[ACTIVE]", "", $hWnd)

	; Copy the code to the clipboard
	ClipPut($winCode)

	; Split the title off the string
	Local $parts = StringSplit($winTitle, $tdeTitle & " - ", $STR_ENTIRESPLIT)

	; Make sure we have a valid title format
	If $parts[0] < 2 Then
		MsgBox($MB_ICONERROR, "Error", "Invalid TradeStation title format.")
		Return
	EndIf

	; Reverse the string to strip the last part off
	Local $rev = StringReverse($parts[2])

	; Find the index of the first colon
	Local $delimIdx = StringInStr($rev, ":")

	; Get the studyType and studyName
	Local $revStudyType = StringLeft($rev, $delimIdx - 2)
	Local $revStudyName = StringMid($rev, $delimIdx + 1)

	; If the study isn't saved, then there's an * on the end of the name, so remove it
	If StringLeft($revStudyType, 1) = "*" Then
		$revStudyType = StringMid($revStudyType, 2)

	; If the study is (read-only) then remove that part of the name
	ElseIf StringLeft($revStudyType, 11) = ")ylno-daer(" Then
		$revStudyType = StringMid($revStudyType, 13)
	EndIf

	; Reverse the studyType and studName back to normal
	Local $studyType = StringLower(StringReverse($revStudyType))
	Local $studyName = StringStripWS(StringReverse($revStudyName), 3)

	; Get the studyName
	$studyName = FormatName($studyName)

	; Build the directory
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
	SaveFile($fullFilePath)
EndFunc

; TradeStation allows these characters, but they can't be in a filename
; Convert the character to its ASCII value
Func FormatName($name)
	$name = StringReplace($name, "\", "#92")
	$name = StringReplace($name, "/", "#47")
	$name = StringReplace($name, ":", "#58")
	$name = StringReplace($name, "*", "#42")
	$name = StringReplace($name, "?", "#63")
	$name = StringReplace($name, "<", "#60")
	$name = StringReplace($name, ">", "#62")
	$name = StringReplace($name, "|", "#124")
	Return $name
EndFunc

; Save the the clipboard text to a file.
Func SaveFile($filePath)
	; If the file exists, delete it so we can overwrite it
	If FileExists($filePath) Then FileDelete($filePath)

	; Get the text to save from the Clipboard
	Local $text = ClipGet()
	If @error Or $text = "" Then
		MsgBox(16, "Clipboard Error", "No code found in clipboard.")
		Return
	EndIf

	; Open the studyFile
	Local $fh = FileOpen($filePath, 2)
	If $fh = -1 Then
		MsgBox(16, "File Error", "Failed to open file: " & $filePath)
		Return
	EndIf

	; Write the contents to the file and close it
	FileWrite($fh, $text)
	FileClose($fh)
EndFunc
