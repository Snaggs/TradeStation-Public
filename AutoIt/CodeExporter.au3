; Git Integration for TradeStation using AutoIt
;
; TS Version: 10 or later
; AutoIT Site: https://www.autoitscript.com/
; AutoIT Version: 3.3.16.1
;
; Credits: A big thanks to moscu for creating the original AutoIT backup
;          script. Parts of that script were used here. 01/02/2006
;          https://community.tradestation.com/Discussions/Topic.aspx?Topic_ID=44956
;
; Author: Snaggs
; Date  : 06/10/24
;
; History:
; ---------------------------------------------------------------------------
; Date       Version    Task
; --------   --------   -----------------------------------------------------
; 06/10/24   1.0.10     Snaggs - Created script to export active study

#include <MsgBoxConstants.au3>
#include <StringConstants.au3>

; Map hot keys to export and exit
HotKeySet("!`", "Export")
HotKeySet("^`", "EndScript")

$rootDir = @UserProfileDir & "\Documents\TradeStation\"   ; Must have trailing \

$activityBarDir = "ActivityBars"
$functionDir = "Functions"
$indicatorDir = "Indicators"
$paintBarDir = "PaintBars"
$probabilityMapDir = "ProbabilityMaps"
$showMeDir = "ShowMes"
$strategyDir = "Strategies"
$TradingAppDir = "TradingApps"

$tdeTitle = "TradeStation Development Environment"

Func EndScript()
   MsgBox($MB_ICONINFORMATION, "Ending Script", "Exiting...", 3)
   Exit
EndFunc

Func Export()
   ; Sample Title bar
   ; TradeStation Development Environment - Mov Avg 1 Line : Indicator

   WinActivate("[CLASS:TSDEV.EXE TRADESTATION]")

   ; Get the text from the title bar
   $winTitle = WinGetTitle("[ACTIVE]")

   ; Get a handle to the control that has the code in it, then get the code from it
   $hWnd = ControlGetFocus("[ACTIVE]")
   $winCode = ControlGetText("[ACTIVE]", "", $hWnd)

   ; Copy the code to the clipboard
   ClipPut($winCode)

   ; Split the title off the string
   $parts = StringSplit($winTitle, $tdeTitle & " - ", $STR_ENTIRESPLIT)

   ; Reverse the string to strip the last part off
   $parts = StringReverse($parts[2])

   ; Find the index of the first colon
   $delimIdx = StringInStr($parts, ":")

   ; Remove the colon and space
   $studyType = StringLeft($parts, $delimIdx - 2)

   ; If the study isn't saved, then there's an * on the end of the name, so remove it
   If StringLeft($studyType, 1) = "*" then
      $studyType = StringMid($studyType, 2)

   ; If the study is (read-only) then remove that part of the name
   ElseIf StringLeft($studyType, 11) = ")ylno-daer(" Then   ; read-only backwards
      $studyType = StringMid($studyType, 13)
   EndIf

   ; Reverse the studyType back to normal
   $studyType = StringReverse($studyType)

   ; Get the study name
   $studyName = StringMid($parts, $delimIdx + 1)

   ; Reverse the study name back to normal
   $studyName = StringReverse($studyName)

   ; Trim off the white spaces
   $studyName = StringStripWS(StringStripWS($studyName, 1), 2)
   $studyName = FormatName($studyName)

   ; Build the directory
   If StringLower($studyType) = "activitybar" Then
      $directory = $rootDir & $activitybarDir
   ElseIf StringLower($studyType) = "function" Then
      $directory = $rootDir & $functionDir
   ElseIf StringLower($studyType) = "indicator" Then
      $directory = $rootDir & $indicatorDir
   ElseIf StringLower($studyType) = "paintbar" Then
      $directory = $rootDir & $paintbarDir
   ElseIf StringLower($studyType) = "probabilitymap" Then
      $directory = $rootDir & $probabilitymapDir
   ElseIf StringLower($studyType) = "showme" Then
      $directory = $rootDir & $showmeDir
   ElseIf StringLower($studyType) = "strategy" Then
      $directory = $rootDir & $strategyDir
   ElseIf StringLower($studyType) = "trading application" Then
      $directory = $rootDir & $tradingAppDir
   Else
      MsgBox($MB_OK, "Unknown Study Type", "Study type unknown: " & $studyType)
      Exit
   EndIf

   ; Add a trailing slash
   $directory = $directory & "\"

   ; If the directory doesn't exist, create it
   If Not FileExists($directory) Then
      DirCreate($directory)
   EndIf

   SaveFile($directory, $studyName)
EndFunc

Func FormatName($name)
   If StringInStr($name,"\") Then
      $name= StringReplace($name, "\", "#92")
   EndIf
   If StringInStr($name,"/") Then
      $name= StringReplace($name, "/", "#47")
   EndIf
   If StringInStr($name,":") Then
      $name= StringReplace($name, ":", "#58")
   EndIf
   If StringInStr($name,"*") Then
      $name= StringReplace($name, "*", "#42")
   EndIf
   If StringInStr($name,"?") Then
      $name= StringReplace($name, "?", "#63")
   EndIf
   If StringInStr($name,"<") Then
      $name= StringReplace($name, "<", "#60")
   EndIf
   If StringInStr($name,">") Then
      $name= StringReplace($name, ">", "#62")
   EndIf
   If StringInStr($name,"|") Then
      $name= StringReplace($name, "|", "#124")
   EndIf
   return $name
EndFunc

;Save the text collected in the clipboard to a file.
Func SaveFile($dir, $name)
   If FileExists($dir & $name & ".txt") Then
      FileDelete($dir & $name & ".txt")
   EndIf
   ; Run Notepad
   Run("notepad.exe '" & $dir & $name & ".txt'")
   ; Now a screen will pop up and ask to save the changes, the window is called
   ; "Notepad" and has some text "Yes" and "No"
   Sleep(300)
   WinActivate("Notepad", "Do you want to create a new file")
   If WinActive("Notepad", "Do you want to create a new file") Then
      Send("y")
   EndIf
   ; Wait for the Notepad become active -
   ; it is titled "name - Notepad" on English systems
   WinActivate($name &".txt - Notepad")
   Sleep(100)
   If WinActive($name &".txt - Notepad") Then
      ; Now that the Notepad window is active paste code
      Send("^v")
      Sleep(100)
      Send("!{F4}")
      Sleep(100)
      Send("!S")
      Sleep(100)
   Else
      MsgBox(0, "3 Not Found", $name & ".txt - Notepad")
      Exit
   EndIf
   ; Now wait for Notepad to close before continuing
   ;MsgBox($MB_OK, "Value", $name)
   WinWaitClose($name & ".txt - Notepad", 7)
EndFunc

MsgBox($MB_ICONINFORMATION, "Starting CodeExport Script", "Starting...", 2)

While 1
   Sleep(100)
WEnd
