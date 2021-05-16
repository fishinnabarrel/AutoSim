#CS
-----------------------------------------------------------------------------
Title:            AutoGamelog.au3
Developer:        David Pyke
Date:             June 30, 2007
Modified:         January 19, 2019
Version:          0.9.0

Description:  Parse DMB boxscores (expanded style) to extract player gamelogs
              for pitchers or batters.  Output to csv.
-----------------------------------------------------------------------------
Copyright (C) 2021 David Pyke.

This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <https://www.gnu.org/licenses/>.
-----------------------------------------------------------------------------
#CE

#include <Array.au3>
#include <Date.au3>
#include <File.au3>
#include <GUIConstantsEx.au3>
#include <EditConstants.au3>
#include <WindowsConstants.au3>
#include <Misc.au3>
#include <String.au3>

Opt("GUIOnEventMode", 1)    ; 0=disable, 1=enable
Opt("MustDeclareVars", 1)   ; 0=no, 1=require pre-declare

Global Const $Title = "AutoGamelog"
Global Const $Version = "0.9"

Global Enum $Batting = 0, $Pitching = 1
Global Enum $AwayTeam = 0, $HomeTeam = 1
; Common output fields.
Global Enum $Name = 0, $Team, $GmDate, $GmNum, $GmSite, $Opp, $GmResult
; Batting log output fields.
Global Enum $Pos = 7, $BOr, $BatGS, $AB, $BatR, $BatH, $BI, $D, $T, _
            $BatHR, $BatBB, $BatK, $SB, $CS, $BatIW, $BatHBP, $SH, $SF, _
            $AVG, $PO, $A, $E, $PB, $Streak
; Pitching log output fields.
Global Enum $DR = 7, $Pitch, $InOut, $Result, $PchGS, $CG, $GF, $W, $L, $S, _
            $INN, $Part, $PchH, $PchR, $ER, $PchBB, $PchK, $PCH, $STR, $BF, _
            $PchHR, $PchIW, $PchHBP, $WP, $DP, $ERA

_Singleton(@ScriptName, 0)

#region --- GUI related code start ---
; Main form
Global $frmMain = GUICreate($Title & " v" & $Version, 400, 125, -1, -1, _
    BitOR($WS_SIZEBOX, $WS_CAPTION, $WS_MINIMIZEBOX, $WS_SYSMENU))

; Main menu
Global $mnuFileMenu = GUICtrlCreateMenu("&File")
Global $mnuFileBoxscoreItem = GUICtrlCreateMenuItem("Choose &boxscore folder ...", $mnuFileMenu)
Global $mnuFileSeperator = GUICtrlCreateMenuItem("", $mnuFileMenu, 2)
Global $mnuFileExitItem = GUICtrlCreateMenuItem("E&xit", $mnuFileMenu)
Global $mnuHelpMenu = GUICtrlCreateMenu("&Help")
Global $mnuHelpAboutItem = GUICtrlCreateMenuItem("&About AutoGamelog", $mnuHelpMenu)

GUICtrlCreateGroup("", 10, 0, 185, 60)
Global $chkBatting = GUICtrlCreateRadio("Game-by-game batting log", 20, 13, 145, 20)
GUICtrlSetState(-1, $GUI_CHECKED)
Global $chkPitching = GUICtrlCreateRadio("Game-by-game pitching log", 20, 37, 145, 18)

GUICtrlCreateGroup("", 205, 0, 185, 60)
Global $btnStart = GUICtrlCreateButton("Generate Gamelog", 220, 15, 155, 37)
GUICtrlSetFont($btnStart, 10)

GUICtrlCreateGroup("", 10, 60, 380, 40)
Global $lblInFiles = GUICtrlCreateLabel("Location of boxscores:", 20, 76, 115, 17)
Global $txtInFiles = GUICtrlCreateInput("", 130, 75, 255, 17, $ES_READONLY)
GUICtrlSetColor(-1, 0x0000FF)
Global $mnuInFilesContext = GUICtrlCreateContextMenu($lblInFiles)
Global $mnuInFilesBoxscoreItem = GUICtrlCreateMenuItem("Choose &boxscore folder ...", $mnuInFilesContext)

; Set the events.
GUICtrlSetOnEvent($mnuFileExitItem, "OnExitClick")
GUICtrlSetOnEvent($mnuFileBoxscoreItem, "OnSelectBoxscoreClick")
GUICtrlSetOnEvent($mnuInFilesBoxscoreItem, "OnSelectBoxscoreClick")
GUICtrlSetOnEvent($mnuHelpAboutItem, "OnAboutClick")
GUICtrlSetOnEvent($btnStart, "OnStartButtonClick")
GUISetOnEvent($GUI_EVENT_CLOSE, "OnExitClick")

GUISetState(@SW_SHOW)
#endregion --- GUI related code end ---

While True
  Sleep(1000) ; Wait for events.
WEnd

Exit (0)


; -----------------------------------------------------------------------------


Func OnSelectBoxscoreClick()
; -----------------------------------------------------------------------------
; Descripton:
;   Prompt user to select a folder containing boxscore files.

; Parameters:

; Return Value:
;   Success         Name of folder containg boxscore files.
;   Failure         "" (blank string) or previously selected boxscore folder, if one exists.
;                   Sets @error to 1.
; -----------------------------------------------------------------------------

  ; Prompt user to select folder.
  Local $boxfileFolder = FileSelectFolder("Select a folder containing Diamond Mind Baseball boxscore files.", "", 2)
  If @error Then
;~      MsgBox(16, $Title & " v" & $Version, "Error: No boxscore folder selected.")
    Local $boxfileFolder = GUICtrlRead($txtInFiles)
    SetError(1)
  EndIf

  GUICtrlSetData($txtInFiles, $boxfileFolder)

  Return
EndFunc   ;==>OnSelectBoxscoreClick


Func OnStartButtonClick()
; -----------------------------------------------------------------------------
; Descripton:
;   Main processing routine.

; Parameters:

; Return Value:
;   Success         None
;   Failure
; -----------------------------------------------------------------------------

  ; Only proceed if a boxscore folder has been chosen.
  If GUICtrlRead($txtInFiles) = "" Then
    MsgBox(16, $Title & " v" & $Version, "Error: No boxscore folder selected.")
    Return 1
  Else
    ; Set default file name, depending on selected log type.
    If GUICtrlRead($chkBatting) = $GUI_CHECKED Then
      Local $defaultName = "Batting_log.csv"
    Else
      Local $defaultName = "Pitching_log.csv"
    EndIf

    ; Prompt user to choose/create name of destination file.
    Local $fileSaveName = FileSaveDialog("Save output file as ...", @DesktopDir, _
        "All (*.*)", 2 + 16, $defaultName)
    If @error Then
      MsgBox(16, $Title & " v" & $Version, "Output file selection cancelled.")
      Return 1
    EndIf

    ; Verify output file can be opened/created.
    If Not _FileCreate($fileSaveName) Then
      MsgBox(16, $Title & " v" & $Version, "Error: Output file can't be opened/created.")
      Return 1
    EndIf
  EndIf

  ; Start progress bar display.
  ProgressOn($Title & " v" & $Version, "Sorting boxscore files ...", "This may take a while.", -1, -1, 16)

  ; Get list of boxscore files.
  Local $boxfileFolder = GUICtrlRead($txtInFiles)
  Local $boxfileList = GetFileList($boxfileFolder)
  If Not IsArray($boxfileList) Then
    ProgressOff() ; Progress bar stopped
    MsgBox(16, $Title & " v" & $Version, "Error: No boxscores found.")
    Return 1
  Else
    ; Sort boxscore files in ascending order.
    _ArraySort($boxfileList, 0, 1)
  EndIf

  ; Open output file for writing (overwrite existing file).
  Local $hOutFile = FileOpen($fileSaveName, 2)
  If $hOutFile = -1 Then    ; Did file open for writing OK?
    ProgressOff()       ; Progress bar stopped
    MsgBox(16, $Title & " v" & $Version, "Error: Output file can't be opened or created.")
    Return 1
  EndIf

  ; Dictionary object to store 'days rest' data for pitching logs
  ; and 'hitting streak' data for batting logs.
  Local $oDictionary = ObjCreate("Scripting.Dictionary")
  If @error Then
    ProgressOff() ; Progress bar stopped
    MsgBox(16, $Title & " v" & $Version, "Error: Unable to create COM object.")
    Return 1
  EndIf

  ; What type of gamelog?
  If GUICtrlRead($chkBatting) = $GUI_CHECKED Then
    Local $gamelogType = $Batting

    ; Save batting header row to output file.
    Local $textLine = '"Name","Team","Date","#","","Opp","GmResult","Pos","BOr",' & _
        '"GS","AB","R","H","BI","D","T","HR","BB","K","SB","CS","IW","HP","SH","SF",' & _
        '"AVG","PO","A","E","PB","HS"'
    FileWriteLine($hOutFile, $textLine)

  ElseIf GUICtrlRead($chkPitching) = $GUI_CHECKED Then
    Local $gamelogType = $Pitching

    ; Save pitching header row to output file.
    Local $textLine = '"Name","Team","Date","#","","Opp","GmResult","DR","Pitcher","InOut","Result",' & _
        '"GS","CG","GF","W","L","S","INN","Part","H","R","ER","BB","K","PCH","STR",' & _
        '"BF","HR","IW","HBP","WP","DP","ERA"'
    FileWriteLine($hOutFile, $textLine)

  EndIf

  Local $beginTime = TimerInit() ; Start timing

  ; Iterate through list of boxscore files, parsing each one.
  For $boxfile = 1 To $boxfileList[0]

    ; Update progress bar display.
    ProgressSet(Round($boxfile / $boxfileList[0] * 100), "Parsing boxscore " & $boxfile & _
        " of " & $boxfileList[0], "Working ...")

    Local $inFilename = $boxfileFolder & "\" & $boxfileList[$boxfile]
    If $gamelogType = $Batting Then
      BoxscoreExtractBatting($inFilename, $hOutFile, $oDictionary)
    Else
      BoxscoreExtractPitching($inFilename, $hOutFile, $oDictionary)
    EndIf
    If @error Then
      ProgressOff()
      MsgBox(16, $Title & " v" & $Version, "Unable to process boxscore: " & $boxfileList[$boxfile])
      Return 1
    EndIf

  Next

  ; Close output file.
  FileClose($hOutFile)

  Local $processingTime = TimerDiff($beginTime) ; Stop timing

  ; Display final progress message with a short delay.
  ProgressSet(100, "All selected boxscores have been parsed.", "Done")
  Sleep(1000)
  ProgressOff()

  ; Format file count and timer data for display in summary message.
  Local $hours = 0, $mins = 0, $secs = 0
  _TicksToTime($processingTime, $hours, $mins, $secs)
  Local $text = "Files processed: " & $boxfileList[0] & @CRLF & @CRLF & _
      "Processing time: " & StringFormat("%i mins :: %i secs", $mins, $secs)

  MsgBox(64 + 262144, $Title & " v" & $Version, $text)

  Return
EndFunc   ;==>OnStartButtonClick


Func GetFileList(Const $boxfileFolder)
; -----------------------------------------------------------------------------
; Descripton:
;   Get list of files selected by the user.

; Parameters:
;   boxfileFolder    Full path of the folder containing boxscore files.

; Return Value:
;   Success         Array containing the filenames of each boxscore file in the boxscore folder.
;                   First element of array contains number of files.
;   Failure         Sets @error to 1, no array is returned.
; -----------------------------------------------------------------------------

  Local $fileList = _FileListToArray($boxfileFolder, "*.box", 1)
  If @error Then
    $fileList = 0
    SetError(1)
  EndIf

  Return $fileList
EndFunc   ;==>GetFileList


Func BoxscoreExtractBatting(Const $inFilename, Const ByRef $hOutFile, ByRef $oDictionary)
; -----------------------------------------------------------------------------
; Descripton:
;   Parse boxscore file and extract batting line for each player.

; Parameters:
;   inFileName      Filename of the boxscore to be parsed.
;   hOutFile        Reference to file handle of output file opened for writing.
;   oDictionary     Reference to Dictionary object that stores name of batter and the
;                   length of his current hitting streak.

; Return Value:
;   Success         No return value (0)
;   Failure         Sets @error to 1 if file not opened successfully.
; -----------------------------------------------------------------------------

  Local Const $BattingSummaryHeader = "AB  R  H BI  D  T HR BB  K SB CS IW HP SH SF"
  Local Const $GameResultSummaryHeader = "R  H  E"

  Local $teamNames[2]           ; Team name abbreviations (Away team = 0, Home team = 1).
  Local $teamRuns[2]                ; Runs scored for each team (Away team = 0, Home team = 1).
  Local $batterLog[80][31]      ; Array to hold batting lines for both teams.
  Local $teamCount = 0          ; Away team = 0, Home team = 1
  Local $totalBatterCount = 0       ; Count of total batters on both teams.
  Local $gameDate, $gameOfDay

  ; Open boxscore file for reading.
  Local $hBoxscore = FileOpen($inFilename, 0)
  If $hBoxscore = -1 Then   ; Did file open for reading OK?
    SetError(1)
    Return
  EndIf

  ; Get game date and team name abbreviations from first line of boxscore file.
  ; e.g. "4/4/2008, Bos08-Tor08, Rogers Centre"
  Local $line = FileReadLine($hBoxscore)
  ParseFirstLine($line, $gameDate, $gameOfDay, $teamNames)

  ; The away team batting summary occurs first in the boxscore.
  ; $teamCount starts at 0 ($AwayTeam = 0) and is incremented by 1 after
  ; each team's batting summary is complete.  The loop condition will
  ; fail when $teamCount is greater than 1 ($HomeTeam = 1).

  While $teamCount <= $HomeTeam
    $line = FileReadLine($hBoxscore)
    If @error = -1 Then
      ; @error = -1 when EOF is reached.
      ; This should never happen, the loop should exit when teamCount > 1.
      ; If this happens the boxscore file is corrupt or not a boxscore file.
      SetError(1)
      Return 1
    EndIf

    ; Get runs scored for each team from game summary.
    Local $gameSummaryStart = StringInStr($line, $GameResultSummaryHeader)
    If $gameSummaryStart Then   ; Game summary header found.
      $line = FileReadLine($hBoxscore)
      $teamRuns[$AwayTeam] = StringMid($line, $gameSummaryStart - 1, 2)
      $line = FileReadLine($hBoxscore)
      $teamRuns[$HomeTeam] = StringMid($line, $gameSummaryStart - 1, 2)

    ElseIf StringInStr($line, $BattingSummaryHeader) Then   ; Batting summary header found
      Local $battingOrder = 0       ; Positions in batting order.
      While True
        $line = FileReadLine($hBoxscore)
        If StringIsSpace(StringLeft($line, 16)) Then ExitLoop   ; End of summary for this team

        ; Increment spot in batting order for players in the starting lineup.
        ; (In-game subs are indented in the boxscore.)
        If StringIsAlpha(StringLeft($line, 1)) Then $battingOrder += 1

        ; Exclude pitchers when DH is used.
        ; Starting pitcher is assigned to 10th spot in batting order when DH is used.
        If $battingOrder > 9 Then ExitLoop

        ; Populate fields in batting log array.
        $batterLog[$totalBatterCount][$Name]        = StringStripWS(StringLeft($line, 16), 3)
        $batterLog[$totalBatterCount][$Team]        = $teamNames[$teamCount]
        $batterLog[$totalBatterCount][$GmDate]      = $gameDate
        $batterLog[$totalBatterCount][$GmNum]       = $gameOfDay
        $batterLog[$totalBatterCount][$GmSite]      = ($teamCount = $AwayTeam) ? "@" : ""
        $batterLog[$totalBatterCount][$Opp]         = $teamNames[Not $teamCount]
        $batterLog[$totalBatterCount][$GmResult]    = (($teamRuns[$teamCount] > $teamRuns[Not $teamCount]) ? "W": "L") & _
            StringFormat(" %2d-%d", $teamRuns[$teamCount], $teamRuns[Not $teamCount])
        $batterLog[$totalBatterCount][$Pos]         = StringUpper(StringStripWS(StringMid($line, 19, 2), 3))
        $batterLog[$totalBatterCount][$BOr]         = AddSuffix($battingOrder)
        $batterLog[$totalBatterCount][$BatGS]       = (StringIsAlpha(StringLeft($line, 1))) ? 1 : 0
        $batterLog[$totalBatterCount][$AB]          = StringStripWS(StringMid($line, 22, 2), 3)
        $batterLog[$totalBatterCount][$BatR]        = StringStripWS(StringMid($line, 25, 2), 3)
        $batterLog[$totalBatterCount][$BatH]        = StringStripWS(StringMid($line, 28, 2), 3)
        $batterLog[$totalBatterCount][$BI]          = StringStripWS(StringMid($line, 31, 2), 3)
        $batterLog[$totalBatterCount][$D]           = StringStripWS(StringMid($line, 34, 2), 3)
        $batterLog[$totalBatterCount][$T]           = StringStripWS(StringMid($line, 37, 2), 3)
        $batterLog[$totalBatterCount][$BatHR]       = StringStripWS(StringMid($line, 40, 2), 3)
        $batterLog[$totalBatterCount][$BatBB]       = StringStripWS(StringMid($line, 43, 2), 3)
        $batterLog[$totalBatterCount][$BatK]        = StringStripWS(StringMid($line, 46, 2), 3)
        $batterLog[$totalBatterCount][$SB]          = StringStripWS(StringMid($line, 49, 2), 3)
        $batterLog[$totalBatterCount][$CS]          = StringStripWS(StringMid($line, 52, 2), 3)
        $batterLog[$totalBatterCount][$BatIW]       = StringStripWS(StringMid($line, 55, 2), 3)
        $batterLog[$totalBatterCount][$BatHBP]      = StringStripWS(StringMid($line, 58, 2), 3)
        $batterLog[$totalBatterCount][$SH]          = StringStripWS(StringMid($line, 61, 2), 3)
        $batterLog[$totalBatterCount][$SF]          = StringStripWS(StringMid($line, 64, 2), 3)
        $batterLog[$totalBatterCount][$AVG]         = StringStripWS(StringMid($line, 68, 4), 3)
        $batterLog[$totalBatterCount][$PO]          = StringStripWS(StringMid($line, 75, 2), 3)
        $batterLog[$totalBatterCount][$A]           = StringStripWS(StringMid($line, 78, 2), 3)
        $batterLog[$totalBatterCount][$E]           = StringStripWS(StringMid($line, 81, 2), 3)
        $batterLog[$totalBatterCount][$PB]          = StringStripWS(StringMid($line, 84, 2), 3)

        ; Calculate batter's 'hitting streak'.
        Local $key = $batterLog[$totalBatterCount][$Team] & " " & $batterLog[$totalBatterCount][$Name]
        If $batterLog[$totalBatterCount][$BatH] > 0 Then
          ; Batter had one or more hits this game, extends hitting streak.
          If $oDictionary.Exists($key) Then
            $batterLog[$totalBatterCount][$Streak] = $oDictionary.Item($key) + 1
            $oDictionary.Item($key) = $oDictionary.Item($key) + 1
          Else
            ; Batter isn't in dictionary yet.
            ; Add him (key) and update his current hitting streak (value) to 1.
            $oDictionary.Add($key, 1)
            $batterLog[$totalBatterCount][$Streak] = 1
          EndIf
        Else
          ; Batter didn't have a hit this game, ends hitting streak.
          If $oDictionary.Exists($key) Then
            $batterLog[$totalBatterCount][$Streak] = 0
            $oDictionary.Item($key) = 0
          Else
            ; Batter isn't in dictionary yet.
            $batterLog[$totalBatterCount][$Streak] = 0
          EndIf
        EndIf
        $totalBatterCount += 1
      WEnd
      $teamCount += 1
    EndIf
  WEnd

  ; Find and add notation to multi-position players
  While True
    $line = FileReadLine($hBoxscore)
    If StringInStr($line, "Temperature:") Then
      ; Temperature is always the last line in the boxscore - end when it's found.
      ExitLoop
    ElseIf StringInStr($line, StringFormat("%-3s:", $teamNames[$AwayTeam])) Then
      $teamCount = $AwayTeam
;~          MsgBox(0,"Multi-position player Testing", $line)
    ElseIf StringInStr($line, StringFormat("%-3s:", $teamNames[$HomeTeam])) Then
      $teamCount = $HomeTeam
;~          MsgBox(0,"Multi-position player Testing", $line)
    EndIf

    Local $playerNameEndIdx = StringInStr($line, "moved to")
    If $playerNameEndIdx Then
      ; Extract player name from line
      ; Player name preceeds 'moved to' phrase in boxscore.
      Local $playerName = StringStripWS(StringMid($line, 5, $playerNameEndIdx - 6), 3)
      Local $playerPosition = StringStripWS(StringMid($line, $playerNameEndIdx + 9, 2), 3)
      Local $searchIdx = _ArraySearch($batterlog, $playerName, 0, 0, 0, 1, 1)

      While $searchIdx <> -1
;~ ConsoleWrite("index: " & $index & "  name: '" & $playerName & "'  teamCount: " & $teamCount & "  position: " & $playerPosition & "  -  " & $line & @CRLF)
        If $batterLog[$searchIdx][$Name] = $playerName And $batterLog[$searchIdx][$Team] = $teamNames[$teamCount] Then
;~                  MsgBox(0,"Multi-position player Testing", $index & " : " & $name & " - " & $teamNames[$teamCount] & " - " & $position)
          $batterLog[$searchIdx][$Pos] &= "-" & StringUpper($playerPosition)
          ExitLoop
        Else
;~                  MsgBox(0,"Multi-position player Testing", $index & " : " & $name & " not found")
;~ ConsoleWrite("*** Not found *** " & $index & " : '" & $playerName & "'")
          $searchIdx = _ArraySearch($batterlog, $playerName, $searchIdx + 1, 0, 0, 1, 1)
          If $searchIdx = -1 Then
            MsgBox(0,"Player name not found.", $searchIdx & " : " & $playerName & " - " & $team)
            SetError(1)
            Return 1
          EndIf
        EndIf
      WEnd

    EndIf

  WEnd

  ; Write all batters for both teams to destination file.
  SaveTeamLogTxt($hOutFile, $batterLog, $totalBatterCount)

  FileClose($hBoxscore)

  Return
EndFunc   ;==>BoxscoreExtractBatting


Func BoxscoreExtractPitching(Const $inFilename, Const ByRef $hOutFile, ByRef $oDictionary)
; -----------------------------------------------------------------------------
; Descripton:
;   Parse boxscore file and extract pitching line for each pitcher.

; Parameters:
;   inFileName      Filename of the boxscore to be parsed.
;   hOutFile        Reference to file handle of output file opened for writing.
;   oDictionary     Reference to Dictionary object that stores name of pitcher and the
;                   date he pitched last.

; Return Value:
;   Success         No return value (0)
;   Failure         Sets @error to 1 if file not opened successfully.
; -----------------------------------------------------------------------------

  Local Const $PitchingSummaryHeader = "INN  H  R ER BB  K PCH STR"
  Local Const $GameResultSummaryHeader = "R  H  E"
  Local Const $DaysRestDefault = "-"
  Local Const $OutsPerInning = 3

  Local $teamNames[2]           ; Team name abbreviations (Away team = 0, Home team = 1).
  Local $teamRuns[2]                ; Runs scored for each team (Away team = 0, Home team = 1).
  Local $teamCount = 0          ; Away team = 0, Home team = 1
  Local $gameDate, $gameOfDay

  ; Open boxscore file for reading.
  Local $hBoxscore = FileOpen($inFilename, 0)
  If $hBoxscore = -1 Then   ; Did file open for reading OK?
    SetError(1)
    Return
  EndIf

  ; Get game date and team names from first line of boxscore file.
  Local $line = FileReadLine($hBoxscore)
  ParseFirstLine($line, $gameDate, $gameOfDay, $teamNames)

  ; The away team pitching summary occurs first in the boxscore.
  ; $teamCount starts at 0 ($AwayTeam = 0) and is incremented by 1 after
  ; each team's pitching summary is complete.  The loop condition will
  ; fail when $teamCount is greater than 1 ($HomeTeam = 1).

  While $teamCount <= $HomeTeam
    $line = FileReadLine($hBoxscore)
    If @error = -1 Then
      ; @error = -1 when EOF is reached.
      ; This should never happen, the loop should exit when teamCount > 1.
      ; If this happens the boxscore file is corrupt or not a boxscore file.
      SetError(1)
      Return 1
    EndIf

    ; Check if current line is the game summary header.
    Local $gameSummaryStart = StringInStr($line, $GameResultSummaryHeader)
    If $gameSummaryStart Then
      ; Get runs scored for each team from game summary.
      $line = FileReadLine($hBoxscore)  ; Read line with away team summary.
      $teamRuns[$AwayTeam] = StringMid($line, $gameSummaryStart - 1, 2)
      $line = FileReadLine($hBoxscore)  ; Read line with home team summary.
      $teamRuns[$HomeTeam] = StringMid($line, $gameSummaryStart - 1, 2)

    ; Check if current line is the pitching summary header.
    ElseIf StringInStr($line, $PitchingSummaryHeader) Then
      Local $pitcherLog[40][33]         ; Array to hold pitching log for team.
      Local $teamPitcherCount = 0       ; Count pitchers in order of appearance
      While True
        $line = FileReadLine($hBoxscore)
        If StringIsSpace(StringLeft($line, 10)) Then ExitLoop ; End of summary for this team

        ; Populate fields in pitching log array.
        $pitcherLog[$teamPitcherCount][$Name]       = StringStripWS(StringLeft($line, 16), 3)
        $pitcherLog[$teamPitcherCount][$Team]       = $teamNames[$teamCount]
        $pitcherLog[$teamPitcherCount][$GmDate]     = $gameDate
        $pitcherLog[$teamPitcherCount][$GmNum]      = $gameOfDay
        $pitcherLog[$teamPitcherCount][$GmSite]     = ($teamCount = $AwayTeam) ? "@" : ""
        $pitcherLog[$teamPitcherCount][$Opp]        = $teamNames[Not $teamCount]
        $pitcherLog[$teamPitcherCount][$GmResult]   = (($teamRuns[$teamCount] > $teamRuns[Not $teamCount]) ? "W" : "L") & _
            StringFormat(" %2d-%d", $teamRuns[$teamCount], $teamRuns[Not $teamCount])
        $pitcherLog[$teamPitcherCount][$DR]         = 0     ; Calculated below.
        $pitcherLog[$teamPitcherCount][$Pitch]      = ""    ; Calculated below.
        $pitcherLog[$teamPitcherCount][$InOut]      = ""    ; Calculated below.
        $pitcherLog[$teamPitcherCount][$Result]     = StringStripWS(StringMid($line, 18, 13), 8)
        $pitcherLog[$teamPitcherCount][$PchGS]      = ($teamPitcherCount = 0) ? 1 : 0
        $pitcherLog[$teamPitcherCount][$CG]         = 0     ; Calculated below.
        $pitcherLog[$teamPitcherCount][$GF]         = 0     ; Calculated below.
        $pitcherLog[$teamPitcherCount][$W]          = StringRegExp($pitcherLog[$teamPitcherCount][$Result], "W")
        $pitcherLog[$teamPitcherCount][$L]          = StringRegExp($pitcherLog[$teamPitcherCount][$Result], "L")
        $pitcherLog[$teamPitcherCount][$S]          = StringRegExp($pitcherLog[$teamPitcherCount][$Result], "\bS")
        $pitcherLog[$teamPitcherCount][$INN]        = StringStripWS(StringMid($line, 33, 2), 3)
        $pitcherLog[$teamPitcherCount][$Part]       = StringStripWS(StringMid($line, 36, 1), 3)
        $pitcherLog[$teamPitcherCount][$PchH]       = StringStripWS(StringMid($line, 38, 2), 3)
        $pitcherLog[$teamPitcherCount][$PchR]       = StringStripWS(StringMid($line, 41, 2), 3)
        $pitcherLog[$teamPitcherCount][$ER]         = StringStripWS(StringMid($line, 44, 2), 3)
        $pitcherLog[$teamPitcherCount][$PchBB]      = StringStripWS(StringMid($line, 47, 2), 3)
        $pitcherLog[$teamPitcherCount][$PchK]       = StringStripWS(StringMid($line, 50, 2), 3)
        $pitcherLog[$teamPitcherCount][$PCH]        = StringStripWS(StringMid($line, 53, 3), 3)
        $pitcherLog[$teamPitcherCount][$STR]        = StringStripWS(StringMid($line, 57, 3), 3)
        $pitcherLog[$teamPitcherCount][$BF]         = StringStripWS(StringMid($line, 63, 2), 3)
        $pitcherLog[$teamPitcherCount][$PchHR]      = StringStripWS(StringMid($line, 66, 2), 3)
        $pitcherLog[$teamPitcherCount][$PchIW]      = StringStripWS(StringMid($line, 69, 2), 3)
        $pitcherLog[$teamPitcherCount][$PchHBP]     = StringStripWS(StringMid($line, 72, 2), 3)
        $pitcherLog[$teamPitcherCount][$WP]         = StringStripWS(StringMid($line, 75, 2), 3)
        $pitcherLog[$teamPitcherCount][$DP]         = StringStripWS(StringMid($line, 78, 2), 3)
        $pitcherLog[$teamPitcherCount][$ERA]        = StringStripWS(StringMid($line, 81, 5), 3)

        ; Calculate 'days rest' field.
        Local $currentDate = FormatDate($gameDate)
        Local $key = $teamNames[$teamCount] & " " & $pitcherLog[$teamPitcherCount][$Name]
        If $oDictionary.Exists($key) Then
          ; Pitcher is in dictionary (has pitched on previous date).
          ; Compare current date with stored previous.
          ; Copy difference between dates (in days) into $daysRest.
          ; Update stored value to current date (becomes new previous date pitched).
          Local $rest = _DateDiff("d", $oDictionary.Item($key), $currentDate) - 1
          $pitcherLog[$teamPitcherCount][$DR] = ($rest < 0) ? 0: $rest
          $oDictionary.Item($key) = $currentDate
        Else
          ; Pitcher isn't in dictionary yet (hasn't pitched before this date).
          ; Add him (key) and the last date he pitched (value).
          $oDictionary.Add($key, $currentDate)
          $pitcherLog[$teamPitcherCount][$DR] = $DaysRestDefault
        EndIf

        $teamPitcherCount += 1
      WEnd

      ; Calculate InOut, Pitch CG and GF fields
      Local $inInning = 0, $outInning = 0
      Local $totalOuts = 0
      For $i = 0 To $teamPitcherCount - 1
        ; Inning(s) pitched in, from inInning to outInning.
        ; Note - The outInning part of this field may be innaccurate if a
        ;        pitcher starts an inning and is removed before recording an out.
        $inInning = Int($totalOuts / $OutsPerInning) + 1
        If $pitcherLog[$i][$INN] = 0 And $pitcherLog[$i][$Part] = 0 Then
          $outInning = $inInning
        Else
          $totalOuts += $pitcherLog[$i][$INN] * $OutsPerInning + $pitcherLog[$i][$Part]
          $outInning = Round($totalOuts / $OutsPerInning + 0.4)
        EndIf

        ; Order of appearance for each team, as # of Total.
        $pitcherLog[$i][$Pitch] = ($i + 1) & " of " & $teamPitcherCount

        If $teamPitcherCount = 1 Then
          ; Only pitcher to appear for team this game, complete game (CG = 1).
          $pitcherLog[$i][$CG] = 1
          $pitcherLog[$i][$InOut] = StringFormat(" CG %d", $outInning)
        ElseIf ($i + 1) = $teamPitcherCount Then
          ; Last pitcher to appear for team this game, game finished (GF = 1).
          $pitcherLog[$i][$GF] = 1
          $pitcherLog[$i][$InOut] = StringFormat(" %2d-%df", $inInning, $outInning)
        Else
          If $i = 0 Then
            ; First pitcher for this team, game started.
            $pitcherLog[$i][$InOut] = StringFormat(" GS-%d", $outInning)
          Else
            ; Middle relief, didn't start, didn't finish.
            $pitcherLog[$i][$InOut] = StringFormat(" %2d-%d", $inInning, $outInning)
          EndIf
        EndIf
      Next

      ; Write all pitchers in team array to destination file.
      SaveTeamLogTxt($hOutFile, $pitcherLog, $teamPitcherCount)

      $teamCount += 1
    EndIf
  WEnd

  FileClose($hBoxscore)

  Return
EndFunc   ;==>BoxscoreExtractPitching


Func ParseFirstLine(Const $line, ByRef $gameDate, ByRef $gameOfDay, ByRef $teamNames)
; -----------------------------------------------------------------------------
; Descripton:
;   Extract date, game number and home and away team abbreviations from first line of file.
;   e.g. "3/31/2009, SF08-Mil08, Miller Park"

; Parameters:
;   line            Line to parse - should be first line of boxscore file.
;   gameDate        Reference to gameDate field - stores date game was played.
;   gameOfDay       Reference to gameOfDay field - stores game number when there is one.
;   teamNames       Reference to 2-element array to store away team and home team name.

; Return Value:
;   Success         gameDate, gameOfDay and teamName by reference.
;   Failure
; -----------------------------------------------------------------------------

  $gameDate = StringLeft($line, StringInStr($line, ",") - 1)
;~   MsgBox(262144, 'Debug line ~' & @ScriptLineNumber, 'Selection:' & @CRLF & '$line' & @CRLF & @CRLF & 'Return:' & @CRLF & $line) ;### Debug MSGBOX

  If StringInStr($line, "game") > 0 Then
    $gameOfDay = StringMid($line, StringInStr($line, "game") + 5, 1)
  Else
    $gameOfDay = 0
  EndIf

;~   Local $Team = _StringBetween($line, ",", "\d", -1, 1)
  Local $Team = StringRegExp($line, ",(.*?)\d", $STR_REGEXPARRAYMATCH)
;~   MsgBox(262144, 'Debug line ~' & @ScriptLineNumber, 'Selection:' & @CRLF & '$Team' & @CRLF & @CRLF & 'Return:' & @CRLF & $Team[0]) ;### Debug MSGBOX

  $teamNames[$AwayTeam] = StringStripWS($Team[0], $STR_STRIPLEADING + $STR_STRIPTRAILING)
;~   $Team = _StringBetween($line, "-", "\d", -1, 1)
  $Team = StringRegExp($line, "-(.*?)\d", $STR_REGEXPARRAYMATCH)
  $teamNames[$HomeTeam] = StringStripWS($Team[0], $STR_STRIPLEADING + $STR_STRIPTRAILING)

  Return
EndFunc   ;==>ParseFirstLine


Func SaveTeamLogTxt(Const ByRef $hOutFile, Const ByRef $teamPlayers, Const $playerCount)
; -----------------------------------------------------------------------------
; Descripton:
;   Save contents of teamPlayers array to output file.

; Parameters:
;   hOutFile        Reference to file handle of output file opened for writing.
;   teamPlayers     Reference to 2-D array of players with log fields.
;   playerCount     Number of players saved in teamPlayers array.

; Return Value:
;   Success         No return value (0)
;   Failure         Set @error to 1 if unable to write a line.
; -----------------------------------------------------------------------------

  Local $textLine = ""
  For $row = 0 To $playerCount - 1
    ; Compose comma delimited line to save.
    $textLine = '"' & $teamPlayers[$row][$Name] & '"'
    For $col = 1 To UBound($teamPlayers, 2) - 1
      $textLine &= ',"' & $teamPlayers[$row][$col] & '"'
    Next

    ; Save line to previously opened destination file.
    FileWriteLine($hOutFile, $textLine)
    If @error Then
      SetError(1)
      ExitLoop
    EndIf
  Next

  Return
EndFunc   ;==>SaveTeamLogTxt


Func AddSuffix($position)
; -----------------------------------------------------------------------------
; Descripton:
;   Add suffix to position (e.g. 1 -> 1st, 2 -> 2nd, 3 -> 3rd, etc ...)

; Parameters:
;   position        Integer value to suffix.

; Return Value:
;   Success         Position with suffix appended.
;   Failure         Set @error = 1.
; -----------------------------------------------------------------------------

  If StringIsInt($position) Then
    Switch $position
      Case 1
        $position &= "st"
      Case 2
        $position &= "nd"
      Case 3
        $position &= "rd"
      Case 4 To 10
        $position &= "th"
      Case Else
        ; Here there be monsters!
        SetError(1)
    EndSwitch
  Else
    SetError(1)
  EndIf

  Return $position
EndFunc     ;==>AddSuffix


Func FormatDate($date)
; -----------------------------------------------------------------------------
; Descripton:
;   Parse date into ('YYYY/MM/DD') format. As required by _DateDiff.

; Parameters:
;   date            Date value to format.

; Return Value:
;   Success         Formatted date.
;   Failure
; -----------------------------------------------------------------------------

  Local $formattedDate = StringMid($date, StringInStr($date, "/", 0, 2) + 1) & "/" & _
      StringLeft($date, StringInStr($date, "/") - 1) & "/" & _
      StringMid($date, StringInStr($date, "/") + 1, _
      StringInStr($date, "/", 0, 2) - StringInStr($date, "/") - 1)

  Return $formattedDate
EndFunc     ;==>FormatDate


Func OnAboutClick()
    ; Displays an About message.
  ; Return value: None
    Local $text = $Title & @CRLF & _
    "Version: " & $Version & @CRLF & _
    "January 2019" & @CRLF&@CRLF & _
    "David Pyke"
    MsgBox(0, "About " & $Title, $text)
    Return
EndFunc     ;==>OnAboutClick


Func OnExitClick()
  ; Exits script in response to Exit Event.
  ; Return vale: None
  Exit(0)
EndFunc     ;==>OnExitClick
